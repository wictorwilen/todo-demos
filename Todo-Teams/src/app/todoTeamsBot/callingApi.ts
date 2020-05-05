import express = require("express");
import debug = require("debug");
import { IncomingCall, ParticiapntsResponse, IncomingCallResponse } from "./calldefs";
import request = require("request");
import uuidv1 = require("uuid/v1");
import * as fs from 'fs';

// Initialize debug logging module
const log = debug("msteams");


// AAD v2 token, required here, cannot use ADAL with the BotId
export async function getAuthToken(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
        request({
            method: "POST",
            uri: `https://login.microsoftonline.com/wictordev.onmicrosoft.com/oauth2/v2.0/token?`,
            headers: {
                "content-type": "application/x-www-form-urlencoded"
            },
            body: `client_id=${process.env.MICROSOFT_APP_ID}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret=${encodeURIComponent(process.env.MICROSOFT_APP_PASSWORD as string)}&grant_type=client_credentials`
        }, (error: any, response: any, body: any) => {
            if (error) {
                reject(error);
            } else {
                resolve(JSON.parse(body).access_token);
            }
        });
    });
}

async function participants(call: string, token: string, clientContext: string): Promise<any[]> {
    return new Promise<any[]>((resolve, reject) => {

        request({
            method: "GET",
            uri: `https://graph.microsoft.com/beta${call}/participants`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            }
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) {
                const users: ParticiapntsResponse = JSON.parse(body);
                log(body);
                users.value.forEach(p => {
                    log(`\tParticiapnt: ${p.info.identity.user.displayName}, ${p.isInLobby}, ${p.isMuted}`);
                });
                resolve(users.value);
            } else {
                log(`Participants failed: ${response.statusCode}, ${response.body}`);
                reject(response.statusCode);
            }
        });
    });
}

async function participant(call: string, user: string, token: string, clientContext: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {
        log(`https://graph.microsoft.com/beta${call}/participants/${user}`);
        request({
            method: "GET",
            uri: `https://graph.microsoft.com/beta${call}/participants/${user}`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            }
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) {
                resolve(JSON.parse(body));
            } else {
                log(`Participant failed: ${response.statusCode}, ${response.body}`);
                reject(response.statusCode);
            }
        });
    });
}

async function unmute(call: string, token: string, clientContext: string): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        request({
            method: "POST",
            uri: `https://graph.microsoft.com/beta${call}/unmute`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            },
            body: JSON.stringify({
                clientContext,
            })
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) {
                log(`Unmuted self`);
                resolve();

            } else {
                log(`Unmuting failed: ${response.statusCode}, ${response.body}`);
                reject(response.statusCode);
            }
        });
    });
}

async function mute(call: string, user: string, token: string, clientContext: string): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        request({
            method: "POST",
            uri: `https://graph.microsoft.com/beta${call}/participants/${user}/mute`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            },
            body: JSON.stringify({
                clientContext,
            })
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) {
                log(`Muted`);
                resolve();

            } else {
                log(`Muting failed: ${response.statusCode}, ${response.body}`);
                reject(response.statusCode);
            }
        });
    });
}

async function playPrompt(call: string, token: string, clientContext: string): Promise<string> {
    return new Promise<string>((resolve, reject) => {

        request({
            method: "POST",
            uri: `https://graph.microsoft.com/beta${call}/playPrompt`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            },
            body: JSON.stringify({
                clientContext,
                prompts: [
                    {
                        "@odata.type": "#microsoft.graph.mediaPrompt",
                        "mediaInfo": {
                            uri: `https://${process.env.HOSTNAME}/assets/audio1.wav`,
                            resourceId: "1D6DE2D4-CD51-4309-8DAA-70768651088E"
                        },
                        "loop": 1
                    }
                ]
            })
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) {
                log(`Audio sent`);
                log(body);
                resolve(body);

            } else {
                log(`Sending audio failed: ${response.statusCode}, ${response.body}`);
                reject(response.statusCode);
            }
        });
    });
}


async function record(call: string, token: string, clientContext: string): Promise<string> {
    return new Promise<string>((resolve, reject) => {

        request({
            method: "POST",
            uri: `https://graph.microsoft.com/beta${call}/record`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            },
            body: JSON.stringify({
                bargeInAllowed: true,
                clientContext,
                maxRecordDurationInSeconds: 20,
                recordingFormat: "wav",
                playBeep: true,
                streamWhileRecording: true,
                stopTones: ["#", "11", "*"]
            })
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) { // DOCBUG: says 2020
                log(`Recording started`);
                log(body);
                // DOCBUG: we should have the lcoation in the headers
                resolve(call + "/operations/" + JSON.parse(body).id);
                // play media

            } else {
                log(`Recording failed: ${response.statusCode}, ${response.body}`);
                reject(response.statusCode);
            }
        });
    });
}

async function getCall(call: string, token: string, clientContext: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {

        request({
            method: "GET",
            uri: `https://graph.microsoft.com/beta${call}`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
            headers: {
                "content-type": "application/json",
                "authorization": `Bearer ${token}`
            }
        }, (error: any, response: any, body: any) => {
            if (response.statusCode === 200) { // DOCBUG: says 2020
                resolve(JSON.parse(body));                // play media
            } else {
                reject(response.statusCode);
            }
        });
    });
}



const currentCalls: any = {};
const recordings: string[] = [];

export const callingApi = async (req: express.Request, res: express.Response, next: express.NextFunction) => {
    const incomingCall: IncomingCall = (req.body as IncomingCallResponse).value[0];
    const token = await getAuthToken();
    log(`Incoming call of type=${incomingCall.changeType}...`);

    switch (incomingCall.changeType) {
        case "created":
            const clientContext = uuidv1();
            log(`This is context: ${clientContext}`);
            log(`Incoming call from ${incomingCall.resourceData.source.identity.user.id} to ${incomingCall.resourceData.targets.map(t => t.identity.application.id).join(",")}`);


            // const t = incomingCall.resourceData.meetingInfo ? incomingCall.resourceData.meetingInfo.token : token;
            // https://docs.microsoft.com/en-us/graph/api/call-answer?view=graph-rest-beta
            request({
                method: "POST",
                uri: `https://graph.microsoft.com/beta${incomingCall.resource}/answer`, // "/app/calls/2d1a0f00-56ae-4991-9b04-5aac475e7055",
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
                    currentCalls[incomingCall.resource] = {
                        state: incomingCall.resourceData.state,
                        clientContext,
                        source: incomingCall.resourceData.source
                    };
                    res.status(200).send(clientContext);

                } else {
                    res.status(response.statusCode).send(response.status);
                }
            });
            break;
        case "deleted":
            // TODO: clean up any persisted stuff
            if (currentCalls[incomingCall.resource] !== undefined) {
                currentCalls[incomingCall.resource].state = "deleted";
            } else
                if (recordings.indexOf(incomingCall.resource) !== -1) {
                    // download the recording
                    log("Downloading recording...");
                    request({
                        method: "GET",
                        uri: incomingCall.resourceData.recordResourceLocation as string,
                        headers: {
                            authorization: "Bearer " + incomingCall.resourceData.recordResourceAccessToken
                        },
                        encoding: null
                    }, async (error: any, response: any, body: any) => {
                        const buffer: Buffer = new Buffer(body, "binary");

                        fs.open("download.wav", "w", (err, fd) => {
                            if (err) {
                                log(`Could not write file: ${err}`);
                            } else {
                                // write the contents of the buffer, from position 0 to the end, to the file descriptor returned in opening our file
                                fs.write(fd, buffer, 0, buffer.length, null, (err2) => {
                                    if (err2) { throw new Error("error writing file: " + err2); }
                                    fs.close(fd, () => {
                                        log("wrote the file successfully");
                                    });
                                });
                            }
                        });
                    });
                }
            res.status(200).send();
            break;
        case "updated":
            if (currentCalls[incomingCall.resource]) {
                currentCalls[incomingCall.resource].state = incomingCall.resourceData.state;

                if (incomingCall.resourceData.state === "established") {
                    try {
                        const call = await getCall(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                        log(JSON.stringify(call));
                        const p2 = await participant(incomingCall.resource, call.myParticipantId, token, currentCalls[incomingCall.resource].clientContext);
                        log(JSON.stringify(p2));
                        const p = await participant(incomingCall.resource, call.source.identity.user.id, token, currentCalls[incomingCall.resource].clientContext);
                        log(JSON.stringify(p));

                    } catch (e) {
                        log(e);
                    }
                    try {
                        await unmute(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                    } catch (e) {
                        log(e);
                    }

                    try {
                        await participants(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                    } catch (e) {
                        log(e);
                    }

                    try {
                        await playPrompt(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                    } catch (e) {
                        log(e);
                    }

                    try {
                        const op = await record(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                        recordings.push(op);
                    } catch (e) {
                        log(e);
                    }
                    setTimeout(async () => {
                        await playPrompt(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                        const parts = await participants(incomingCall.resource, token, currentCalls[incomingCall.resource].clientContext);
                        if (parts) {
                            parts.forEach(async (p) => {
                                await mute(incomingCall.resource, p.id, token, currentCalls[incomingCall.resource].clientContext);
                            });
                        }
                    }, 8000);
                }
            }
            res.status(200).send();
            break;
        default:

    }


}
