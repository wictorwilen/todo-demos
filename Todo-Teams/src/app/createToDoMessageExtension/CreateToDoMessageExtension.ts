import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor, ITaskInfo } from "botbuilder-teams-messagingextensions";


// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/createToDoMessageExtension/config.html")
export default class CreateToDoMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        const card = CardFactory.heroCard("Test", "Test", [`https://${process.env.HOSTNAME}/assets/icon.png`]);

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run

            return Promise.resolve({
                type: "result",
                attachmentLayout: "grid",
                attachments: [
                    card
                ]
            } as MessagingExtensionResult);
        } else {
            // the rest
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    card
                ]
            } as MessagingExtensionResult);
        }
    }

    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Create To-Do Configuration",
            value: `https://${process.env.HOSTNAME}/createToDoMessageExtension/config.html`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

    public async onSubmitAction(context: TurnContext, value: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        const card = CardFactory.heroCard("Test", "Test", [`https://${process.env.HOSTNAME}/assets/icon.png`]);
        log("Create To-Do");
        log(value);
        return Promise.resolve({
            type: "result",
            attachmentLayout: "grid",
            attachments: [
                card
            ]
        } as MessagingExtensionResult);
    }

    public async onFetchTask(context: TurnContext, value: { commandContext: any, context: any, messagePayload: any }): Promise<ITaskInfo> {
        return Promise.resolve({
            title: "Task Module",
            card: CardFactory.adaptiveCard({
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                type: "AdaptiveCard",
                version: "1.0",
                body: [
                    {
                        type: "TextBlock",
                        text: "Please enter your e-mail"
                    },
                    {
                        type: "Input.Text",
                        id: "myEmail",
                        placeholder: "youremail@example.com",
                        style: "email"
                    },
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "OK",
                        data: { id: "unique-id" }
                    }
                ]
            })
        });
    }
    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        log(value);
        return Promise.resolve();
    }
}

