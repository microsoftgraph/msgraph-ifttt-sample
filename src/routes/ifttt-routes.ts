/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client } from "@microsoft/microsoft-graph-client";
import { Event, Message, User, Group } from "@microsoft/microsoft-graph-types";
import express from 'express';
import fetch from "node-fetch";
import { GraphClient } from "../helpers/graph-client";
import { Middleware } from "../helpers/middleware";
import { BirthdayResponseData } from "../models/birthday-response-model";
import { EventResponseData } from "../models/event-response-model";
import { MessageResponseData } from "../models/message-response-model";
import { OptionsResponseData } from "../models/options-response-model";

// Routes for ifttt endpoint tests (and some misc graph items)
export class IftttRoutes {

    private _router;
    public getRouter() {
        return this._router;
    }

    constructor() {

        this._router = express.Router();

        // endpoints for IFTTT setup and tests
        this._router.get('/status', Middleware.serviceKeyCheck, this.getStatus);
        this._router.get('/user/info', Middleware.authorizationCheck, this.getUserInfo);
        this._router.post('/test/setup', Middleware.serviceKeyCheck, this.postTestSetup);

        // triggers for mail, calendar, and user data
        this._router.post('/triggers/message_mention', Middleware.authorizationCheck, this.postTrigger_Message_Mention);
        this._router.post('/triggers/event_created', Middleware.authorizationCheck, this.postTrigger_EventCreated);
        this._router.post('/triggers/group_member_birthday', Middleware.authorizationCheck, Middleware.triggerFieldsCheck, this.postTrigger_Birthday);
        this._router.post('/triggers/group_member_birthday/fields/group_id/options', Middleware.authorizationCheck, this.getParam_Group);
    }

    // Check status
    private getStatus(req, res) {

        res.status(200).send();
    }

    // Get user info (for ifttt auth)
    private async getUserInfo(req, res) {
        try {
            const client:Client = GraphClient.getClient(req.headers.authorization);
            const user:User = await client.api("/me").get().catch(e=>{
                return Promise.reject(e);
            })

            res.json({
                data: {
                    id: user.id,
                    name: user.userPrincipalName
                }
            });
        } catch (e) {
            console.log(e);
            res.status(401).send({
                "errors": [{
                    "message": `Retrieving Graph data failed : ${e.message}`
                }]
            });
        }
    }

    // The test/setup endpoint
    private async postTestSetup(req, res) {

        try {

            const formData = new URLSearchParams();
            formData.append('scope', 'Notes.ReadWrite.All Calendars.Read identityriskevent.read.all directory.read.all group.readwrite.all user.readwrite.all mail.readwrite');
            formData.append('client_id', process.env.CLIENT_ID);
            formData.append('client_secret', process.env.CLIENT_SECRET);
            formData.append('grant_type', 'password');
            formData.append('username', process.env.TEST_USER);
            formData.append('password', process.env.TEST_PWD);

            const tokenUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
            const tokenResponse = await fetch(tokenUrl, {
                method: 'POST',
                body: formData
            });

            const tokenJson = await tokenResponse.json();

            if (!tokenResponse.ok) {
                throw tokenJson.error_description;
            }

            const token = tokenJson.access_token;
            console.log(`received authtoken: ${token}`);

            const testData =  {
                data: {
                    samples: {
                        triggers: {
                            group_member_birthday : {
                                group_id: `4e336cbe-360b-4feb-9011-6c7bf25a3d70`
                            }
                        },
                        actions: {
                            create_message: {
                                team_id: 'null',
                                team_name: 'Test team',
                                channel_name: 'Test channel',
                                message: `Test message`
                            }
                            // ,
                            // create_channel: {
                            //     team_id: 'null',
                            //     team_name: 'Test team',
                            //     channel_name: `Channel_${generateUniqueId()}`
                            // },
                            // create_team: {
                            //     team_name: `Team_${generateUniqueId()}`
                            // }
                        }
                    },
                    accessToken: token
                }
            };

            res.status(200).json(testData);
        }
        catch(e) {
            res.status(401).send({
                "errors": [{
                    "message": `Test setup error: ${e.message}`
                }]
            });
        }
    }

    // Trigger endpoint: New event created
    private async postTrigger_EventCreated(req, res) {

        const authHeader = req.headers.authorization;
        const numOfItems = (req.body.limit !== undefined) ? req.body.limit : 50;

        // get the next N events for the user
        const startDateFilter = new Date();

        try {
            const url = `/me/events?$filter=start/dateTime ge '${startDateFilter.toISOString()}'&$orderBy=lastModifiedDateTime desc&$top=${numOfItems}`;
            const client:Client = GraphClient.getClient(authHeader);
            let graphItems:Event[] = await client.api(url).get().then(v=>{return v.value}).catch(e=> {
                return Promise.reject(e);
            });

            console.log(`found ${graphItems ? graphItems.length : 0} items`);

            // Filter to items created on same day of event
            graphItems = graphItems.filter((v) => {

                const cd = new Date(v.createdDateTime);
                const sd = new Date(v.start.dateTime);
                const hrs = (cd.getTime() - sd.getTime()) / (1000 * 60 * 60);

                return hrs < 24;
            });

            console.log(`found ${graphItems ? graphItems.length : 0} items`);

            if (graphItems.length < 3)
            {
                /// need to fake 3 items so that IFTTT will behave
                graphItems.push(
                    { id:`123`, createdDateTime:`2018-01-01`, organizer:{emailAddress:{name:'anon'}}, subject:'subj', start:{dateTime:`2018-01-01`}}
                    ,{ id:`456`, createdDateTime:`2018-01-02`, organizer:{emailAddress:{name:'anon'}}, subject:'subj', start:{dateTime:`2018-01-01`}}
                    ,{ id:`789`, createdDateTime:`2018-01-03`, organizer:{emailAddress:{name:'anon'}}, subject:'subj', start:{dateTime:`2018-01-01`}}
                );
            }

            // convert graph events to IFTTT Events
            const responseItems = EventResponseData.fromGraph(graphItems, numOfItems);

            res.status(200).json(responseItems);
        } catch (e) {
            res.status(401).send({
                "errors": [{
                    "message": `Retrieving Graph data failed : ${e.message}`
                }]
            });
        }
    }

    /**
     * returns a set of messages that have the user @mentioned in them
     */
    private async postTrigger_Message_Mention(req, res){
        const authHeader = req.headers.authorization;
        const numOfItems = (req.body.limit !== undefined) ? req.body.limit : 50;

        const client:Client = GraphClient.getClient(authHeader, 'beta');
        const request = client.api(`/me/messages?$filter=MentionsPreview/IsMentioned%20eq%20true&$select=Subject,Sender,ReceivedDateTime,MentionsPreview`);
        let graphItems:Message[];

        // get the next N events for the user
        graphItems = await request.get().then(v=>{return v.value}).catch(e => {
                console.log(e);
                res.status(401).send({
                    "errors": [{
                        "message": `Retrieving Graph data failed : ${e}`
                    }]
                });
                return;
            });

        console.log(`found ${graphItems ? graphItems.length : 0} items`);

        if (graphItems.length < 3)
        {
            console.log(`Faking 3 items for ifttt`);
            /// need to fake 3 items so that IFTTT will behave
            graphItems.push(
                 { id:`123`, subject:`fake message`, sender:{emailAddress:{name:`Fake`}}, receivedDateTime:`2019-01-01`}
                ,{ id:`456`, subject:`fake message`, sender:{emailAddress:{name:`Fake`}}, receivedDateTime:`2019-01-02`}
                ,{ id:`789`, subject:`fake message`, sender:{emailAddress:{name:`Fake`}}, receivedDateTime:`2019-01-03`}
            );
        }

        // convert graph events to IFTTT Events
        const responseItems = MessageResponseData.fromGraph(graphItems, numOfItems);

        res.status(200).json(responseItems);
    }

    /**
     * returns a set of users in a select group that happen to have birthdays today
     */
    private async postTrigger_Birthday(req, res){
        const authHeader = req.headers.authorization;
        const numOfItems = (req.body.limit !== undefined) ? req.body.limit : 50;
        const groupId = req.body.triggerFields.group_id;

        const client:Client = GraphClient.getClient(authHeader);
        const graphItems:User[] = [];
        try{
            if (groupId === undefined)  {
                res.status(400).send({
                    "errors": [{
                        "message": `Missing triggerFields : group_id`
                    }]
                    });
                    return;
            }

            // note: not possible to list birthday in group list today so will cycle each for birthday
            const queryMembers = await client.api(`groups/${groupId}/members`).get().catch(e=>{throw new Error(e.message);});
            const groupUsers:User[] = queryMembers.value;// .then(v=>{return v.value})
            // note: odata query filter is not possible at this time to filter to birthdays

            const now:Date = new Date();
            for (const user of groupUsers)
            {
                const gUser:User = await client.api(`users/${user.id}?$select=id,displayName,birthday`).get().catch(e=>{throw new Error(e.message);});
                const birthday = new Date(gUser.birthday);
                if ( birthday.getDate() === now.getDate() && birthday.getMonth() === now.getMonth())
                {
                    // it's this person's birthday!
                    gUser.id = gUser.id + now.getFullYear().toString();
                    gUser.birthday = gUser.birthday.replace( new RegExp("/d{4}/"),now.getFullYear().toString());
                    graphItems.push(gUser);
                    console.log(`${gUser.displayName}'s birthday!`);
                }
                console.log(`${gUser.birthday} for ${gUser.displayName}`);
            }
            // res.status(200).json(graphItems);
        }
        catch(e)
        {
            console.log(e);
            res.status(401).send({
                "errors": [{
                    "message": `Retrieving Graph data failed : ${e.message}`
                }]
            });
            return;
        }

        console.log(`found ${graphItems ? graphItems.length : 0} items`);

        if (graphItems.length < 3)
        {
            console.log(`Faking 3 items for ifttt`);
            /// need to fake 3 items so that IFTTT will behave
            graphItems.push(
                 { id:`123`, displayName:`fake`, birthday:`2019-01-01`}
                ,{ id:`456`, displayName:`fake`, birthday:`2019-01-02`}
                ,{ id:`789`, displayName:`fake`, birthday:`2019-01-03`}
            );
        }

        // convert graph events to IFTTT Events
        const responseItems = BirthdayResponseData.fromGraph(graphItems, numOfItems);

        res.status(200).json(responseItems);
    }

    private async getParam_Group(req, res) {
        const authHeader = req.headers.authorization;

        // get the Teams joined by the user
        const url = `/me/memberOf`;
        const client:Client = GraphClient.getClient(authHeader);
        const request = client.api(url);
        let graphItems:Group[];

        // get the next N events for the user
        graphItems = await request.get().then(v=>{return v.value}).catch(e => {
            console.log(e);
            res.status(401).send({
                "errors": [{
                    "message": `Retrieving Graph data failed : ${e}`
                }]
            });
            return;
        });

        console.log(`found ${graphItems ? graphItems.length : 0} items`);

        // convert graph items to IFTTT items
        const responseItems = OptionsResponseData.fromGraph(graphItems);  // need eh on this?
        res.status(200).json(responseItems);
    }
}
