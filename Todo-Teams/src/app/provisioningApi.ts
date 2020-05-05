import express = require("express");
import debug = require("debug");

import { DefaultHttpClient, TokenCredentials, ServiceClient, RequestPrepareOptions } from "@azure/ms-rest-js";
import request = require("request");

// Initialize debug logging module
const log = debug("msteams");


// AAD v2 token, required here, cannot use ADAL with the BotId
async function getAuthToken(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
        request({
            method: "POST",
            uri: `https://login.microsoftonline.com/wictordev.onmicrosoft.com/oauth2/v2.0/token?`,
            headers: {
                "content-type": "application/x-www-form-urlencoded"
            },
            body: `client_id=${process.env.CLIENT_APP_ID}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret=${encodeURIComponent(process.env.CLIENT_APP_SECRET as string)}&grant_type=client_credentials`
        }, (error: any, response: any, body: any) => {
            if (error) {
                reject(error);
            } else {
                resolve(JSON.parse(body).access_token);
            }
        });
    });
}

export const provisionApi = (options: any): express.Router => {
    const router = express.Router();
    router.post("/", async (req: express.Request, res: express.Response, next: express.NextFunction) => {
        log("start team creation");
        const url = `https://graph.microsoft.com/beta/teams`;

        const credentials = new TokenCredentials(await getAuthToken());
        const client = new ServiceClient(credentials, {});
        const requestOptions: RequestPrepareOptions = {
            url,
            method: "POST",
            body: {
                "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
                "displayName": req.body.displayName,
                "description": req.body.description,
                "owners@odata.bind": [
                    "https://graph.microsoft.com/v1.0/users/f4b157b0-3e3e-410c-9648-b9bd5d53a689"
                ],
                "channels": [
                    {
                        displayName: "Announcements 📢",
                        isFavoriteByDefault: true,
                        description: "This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements.",
                        tabs: [
                            {
                                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('4e099bcc-9c5a-4f6e-80a1-ad7f599b480c')",
                                "name": "Todo",
                                "configuration": {
                                    contentUrl: "https://todo-teams.ngrok.io/todoTeamsTab/?noOfItems=10"
                                }
                            },
                            {
                                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.youtube')",
                                "name": "A Pinned YouTube Video",
                                "configuration": {
                                    contentUrl: "https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ",
                                    websiteUrl: "https://www.youtube.com/watch?v=X8krAMdGvCQ"
                                }
                            }
                        ]
                    }
                ],
                "memberSettings": {
                    allowCreateUpdateChannels: true,
                    allowDeleteChannels: true,
                    allowAddRemoveApps: true,
                    allowCreateUpdateRemoveTabs: true,
                    allowCreateUpdateRemoveConnectors: true
                },
                "installedApps": [
                    {
                        "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('4e099bcc-9c5a-4f6e-80a1-ad7f599b480c')"
                    }
                ]
            }
        };
        const result = await client.sendRequest(requestOptions);
        log(result);
        res.status(202).send();
    });

    return router;
};
