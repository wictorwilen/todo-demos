import { BotDeclaration, MessageExtensionDeclaration, IBot, PreventIframe, BotCallingWebhook } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState, WaterfallDialog, WaterfallStepContext, DialogTurnResult } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, InvokeResponse, BotFrameworkAdapter, AttachmentLayoutTypes } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import TodoTeamsMessageExtension from "../todoTeamsMessageExtension/TodoTeamsMessageExtension";
import { TeamsContext, TeamsActivityProcessor, IMembersAddedEvent } from "botbuilder-teams";
import { OAuthPrompt } from "botbuilder-dialogs";
import { DefaultHttpClient, TokenCredentials, ServiceClient, RequestPrepareOptions } from "@azure/ms-rest-js";
import CreateToDoMessageExtension from "../createToDoMessageExtension/CreateToDoMessageExtension";
import * as _ from "lodash";
import express = require("express");
import { IncomingCallProcessor, IncomingCallHandler, call, resultInfo } from "botbuilder-calling-processor";


// Initialize debug logging module
const log = debug("msteams");

const loginPrompt = new OAuthPrompt("LoginDialog", {
    connectionName: process.env.MICROSOFT_APP_OAUTHSETTING as string,
    text: "Please login",
    title: "Login",
    timeout: 30000 // User has 5 minutes to login.
});

import todoCard = require("./todoCard.json");
import { getAuthToken } from "./callingApi";
import * as _Request from "request";
import uuidv1 = require("uuid/v1");


/**
 * Implementation for TodoTeamsBot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)
@PreventIframe("/todoTeamsBot/faq.html")
export class TodoTeamsBot implements IBot {
    private readonly conversationState: ConversationState;
    /**
     * Local property for CreateToDoMessageExtension
     */
    @MessageExtensionDeclaration("createToDoMessageExtension")
    // tslint:disable-next-line: variable-name
    private _createToDoMessageExtension: CreateToDoMessageExtension;
    /**
     * Local property for TodoTeamsMessageExtension
     */
    @MessageExtensionDeclaration("todoTeamsMessageExtension")
    // tslint:disable-next-line: variable-name
    private _todoTeamsMessageExtension: TodoTeamsMessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;
    private commandState: StatePropertyAccessor<any>;
    private payloadState: StatePropertyAccessor<any>;
    private readonly activityProc = new TeamsActivityProcessor();

    /**
     * The constructor
     * @param conversationState 
     */
    public constructor(conversationState: ConversationState) {
        // Message extension CreateToDoMessageExtension
        this._createToDoMessageExtension = new CreateToDoMessageExtension();

        // Message extension TodoTeamsMessageExtension
        this._todoTeamsMessageExtension = new TodoTeamsMessageExtension();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.commandState = conversationState.createProperty("commandState");
        this.payloadState = conversationState.createProperty("payloadState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));
        this.dialogs.add(loginPrompt);
        this.dialogs.add(new WaterfallDialog("GraphDialog", [
            this.promptStep.bind(this),
            this.processStep.bind(this)
        ]));

        this.onIncomingCall = this.onIncomingCall.bind(this);
        // Set up the Activity processing

        this.activityProc.messageActivityHandler = {
            // Incoming messages
            onMessage: async (context: TurnContext): Promise<void> => {
                // get the Microsoft Teams context, will be undefined if not in Microsoft Teams
                const teamsContext: TeamsContext = TeamsContext.from(context);

                // TODO: add your own bot logic in here
                switch (context.activity.type) {
                    case ActivityTypes.Message:
                        const dc = await this.dialogs.createContext(context);
                        if (context.activity.value) {
                            // we have an empty message, but a value payload
                            await dc.continueDialog();
                            if (!dc.context.responded) {
                                await dc.beginDialog("GraphDialog");
                            }
                        } else {

                            const text = teamsContext ?
                                teamsContext.getActivityTextWithoutMentions().toLowerCase() :
                                context.activity.text;

                            if (text.startsWith("hello")) {
                                await context.sendActivity("Oh, hello to you as well!");
                            } else if (text.startsWith("help")) {
                                await dc.beginDialog("help");
                            } else if (text.startsWith("signout")) {
                                const botAdapter = dc.context.adapter as BotFrameworkAdapter;
                                await botAdapter.signOutUser(dc.context, process.env.MICROSOFT_APP_OAUTHSETTING as string);
                                await dc.context.sendActivity("You are now signed out.");
                            } else {
                                await dc.continueDialog();
                                if (!dc.context.responded) {
                                    await dc.beginDialog("GraphDialog");
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }

                // Save state changes
                return this.conversationState.saveChanges(context);
            }
        };

        this.activityProc.conversationUpdateActivityHandler = {
            onTeamMembersAdded: async (event: IMembersAddedEvent): Promise<void> => {
                log("Conversation update");
                // Display a welcome card when the bot is added to a conversation
                if (event.membersAdded && event.membersAdded.length !== 0) {
                    for (const idx in event.membersAdded) {
                        if (event.membersAdded[idx].id !== event.turnContext.activity.recipient.id) {
                            const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                            await event.turnContext.sendActivity({ attachments: [welcomeCard] });
                        }
                    }
                }
            }
        }

        // Message reactions in Microsoft Teams
        this.activityProc.messageReactionActivityHandler = {
            onMessageReaction: async (context: TurnContext): Promise<void> => {
                const added = context.activity.reactionsAdded;
                if (added && added[0]) {
                    await context.sendActivity({
                        textFormat: "xml",
                        text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                    });
                }
            }
        };

        
        this.activityProc.invokeActivityHandler = {
            onInvoke: async (turnContext: TurnContext): Promise<InvokeResponse> => {
                // Sanity check the Activity type and channel Id.
                if (turnContext.activity.type === ActivityTypes.Invoke && turnContext.activity.channelId !== "msteams") {
                    throw new Error("The Invoke type is only valid on the MS Teams channel.");
                };
                log("invokeactivityhandler");
                const dc = await this.dialogs.createContext(turnContext);
                await dc.continueDialog();
                if (!turnContext.responded) {
                    await dc.beginDialog("GraphDialog");
                }
                return Promise.resolve({
                    status: 200
                });
            }
        }
    }

    /**
     * The Bot Framework `onTurn` handlder.
     * The Microsoft Teams middleware for Bot Framework uses a custom activity processor (`TeamsActivityProcessor`)
     * which is configured in the constructor of this sample
     */
    public async onTurn(context: TurnContext): Promise<any> {
        // transfer the activity to the TeamsActivityProcessor
        await this.activityProc.processIncomingActivity(context);
    }

    // tslint:disable-next-line: member-ordering
    private callProcessor: IncomingCallProcessor;
    @BotCallingWebhook("/api/calling")
    public async onIncomingCall(req: express.Request, res: express.Response, next: express.NextFunction) {
        log("/api/calling ping");
        if (!this.callProcessor) {
            const handler: IncomingCallHandler = {
                onIncomingCall: async (resource: string, resourceData: call): Promise<string> => {
                    log("Incoming call");
                    log(`https://graph.microsoft.com/beta${resource}/answer`)
                    return new Promise<string>(async (resolve, reject) => {
                        const token = await getAuthToken();
                        const clientContext = uuidv1();
                        // answer the call
                        _Request({
                            method: "POST",
                            uri: `https://graph.microsoft.com/beta${resource}/answer`,
                            headers: {
                                "content-type": "application/json",
                                "authorization": `Bearer ${token}`
                            },
                            body: JSON.stringify({
                                callbackUri: `https://${process.env.HOSTNAME}/api/calling`,
                                acceptedModalities: ["audio", "video"],
                                mediaConfig: {
                                    "@odata.type": "#microsoft.graph.serviceHostedMediaConfig",
                                    "preFetchMedia": [
                                        {
                                            uri: `https://${process.env.HOSTNAME}/assets/audio1.wav`,
                                            resourceId: "1D6DE2D4-CD51-4309-8DAA-70768651088E"
                                        }
                                    ] 
                                }
                            })
                        }, async (error: any, response: any, body: any) => {
                            if (response.statusCode === 202) {
                                log(`Call answered!`);
                                resolve(clientContext); 
                            } else {
                                log(error);
                                log(response.statusCode);
                                reject(`Invalid response from Microsoft Graph: ${response.status}`);
                            }
                        }); 
                    });
                },
                onCallTerminated: (resource: string, result: resultInfo): Promise<void> => {
                    log("Call termindated");
                    return Promise.resolve();
                }
            };
            this.callProcessor = new IncomingCallProcessor(handler);
        }
        this.callProcessor.process(req, res);
    }

    private async promptStep(step: WaterfallStepContext<any>): Promise<DialogTurnResult> {
        const activity = step.context.activity;

        if (activity.type === ActivityTypes.Message && !(/\d{6}/).test(activity.text)) {
            await this.commandState.set(step.context, activity.text);
            await this.payloadState.set(step.context, activity.value);
            await this.conversationState.saveChanges(step.context);
        }
        return await step.beginDialog("LoginDialog");
    }


    private async processStep(step: WaterfallStepContext<any>): Promise<DialogTurnResult> {
        // We do not need to store the token in the bot. When we need the token we can
        // send another prompt. If the token is valid the user will not need to log back in.
        // The token will be available in the Result property of the task.
        const tokenResponse = step.result;

        // If the user is authenticated the bot can use the token to make API calls.
        if (tokenResponse !== undefined) {
            let parts = await this.commandState.get(step.context);
            const value = await this.payloadState.get(step.context);
            if (value) {
                const credentials = new TokenCredentials(tokenResponse.token);
                const client = new ServiceClient(credentials, {});
                const request: RequestPrepareOptions = {
                    url: `https://graph.microsoft.com/beta/me/outlook/tasks/${value.id}/complete`,
                    method: "POST"
                };
                const result = await client.sendRequest(request);
                await step.context.sendActivity(`You're working hard! Good job!`);
            } else {
                if (!parts) {
                    parts = step.context.activity.text;
                }
                const command = parts.split(" ")[0].toLowerCase();
                if (command === "me") {
                    const credentials = new TokenCredentials(tokenResponse.token);
                    const client = new ServiceClient(credentials, {});
                    const request: RequestPrepareOptions = {
                        url: "https://graph.microsoft.com/beta/me?$select=displayName",
                        method: "GET"
                    };
                    const result = await client.sendRequest(request);
                    log(result.parsedBody);
                    await step.context.sendActivity(`You are ${result.parsedBody.displayName}`);
                } else if (command === "tasks") {
                    const credentials = new TokenCredentials(tokenResponse.token);
                    const client = new ServiceClient(credentials, {});
                    const request: RequestPrepareOptions = {
                        url: "https://graph.microsoft.com/beta/me/outlook/tasks?$select=subject,status,id,importance&$filter=status eq 'notStarted'",
                        method: "GET"
                    };
                    const result = await client.sendRequest(request);
                    log(result.parsedBody);
                    const cards = result.parsedBody.value.map((r: { subject: string; status: string; importance: string; id: string; }) => {
                        const c: any = _.cloneDeep(todoCard);
                        c.body[0].text = r.subject;
                        c.body[1].text = r.status;
                        c.body[2].text = r.importance;
                        c.actions[0].data.id = r.id;
                        return CardFactory.adaptiveCard(c);
                    });
                    console.log(JSON.stringify({
                        text: `Your tasks`,
                        attachmentLayout: AttachmentLayoutTypes.List,
                        attachments: cards
                    }));
                    await step.context.sendActivity({
                        text: `Your tasks`,
                        attachmentLayout: AttachmentLayoutTypes.List,
                        attachments: cards
                    });
                } else if (command === "token") {
                    await step.context.sendActivity(`Your token is: ${tokenResponse.token}`);

                } else {
                    await step.context.sendActivity(`I have no idea what you're talking about...`);
                }
            }
        } else {
            // Ask the user to try logging in later as they are not logged in.
            await step.context.sendActivity(`We couldn't log you in. Please try again later.`);
        }
        return await step.endDialog();
    };
}
