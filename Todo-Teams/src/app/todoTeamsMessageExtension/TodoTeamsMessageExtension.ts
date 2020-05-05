import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/todoTeamsMessageExtension/config.html")
export default class TodoTeamsMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        // const card = CardFactory.heroCard(
        //     "Test",
        //     "Test",
        //     ["https://todo-teams.ngrok.io/assets/icon.png"],
        //     [{
        //         type: "Action.Submit",
        //         title: "More details",
        //         value: "unique-id"
        //     }]);
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: "Headline"
                    },
                    {
                        type: "TextBlock",
                        text: "Description"
                    },
                    {
                        type: "Image",
                        url: "https://todo-teams.ngrok.io/assets/icon.png"
                    }
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "More details",
                        data: {
                            action: "moreDetails",
                            id: "1234-5678"
                        }
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        const preview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: "Headline",
                text: "Description",
                images: [
                    {
                        url: "https://todo-teams.ngrok.io/assets/icon.png"
                    }
                ],
                tap: { type: "invoke", title: "Option 1", value: { option: "opt1" } }
            }
        };


        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    { ...card, preview }
                ]
            } as MessagingExtensionResult);
        } else {
            // the rest
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    { ...card, preview }
                ]
            } as MessagingExtensionResult);
        }
    }


    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Todo Teams Configuration",
            value: "https://todo-teams.ngrok.io/todoTeamsMessageExtension/config.html"
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        if (value.action === "moreDetails") {
            log(`I got this ${value.id}`);
        }
        return Promise.resolve();
    }

    public async onSelectItem(context: TurnContext, value: any): Promise<MessagingExtensionResult> {
        log("onSelectIteam");
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: "A completley different card"
                    },
                    {
                        type: "Image",
                        url: "https://todo-teams.ngrok.io/assets/icon.png"
                    }
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "More details",
                        data: {
                            action: "moreDetails",
                            id: "1234-5678"
                        }
                    },
                    {
                        type: "Action.Submit",
                        title: "Second action",
                        data: {
                            action: "evenMoreDetails",
                            id: "1234-5678"
                        }
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        const preview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: "Headline 222",
                text: "Description 222",
                images: [
                    {
                        url: "https://todo-teams.ngrok.io/assets/icon.png"
                    }
                ],
                tap: { type: "invoke", title: "Option 1", value: { option: "opt1" } }
            }
        };

        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [
                {...card, preview}
            ]
        } as MessagingExtensionResult);
    }

}
