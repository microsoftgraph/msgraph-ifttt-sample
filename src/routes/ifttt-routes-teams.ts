/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client } from "@microsoft/microsoft-graph-client";
import { Channel, Group } from "@microsoft/microsoft-graph-types";
import express from 'express';
import { Guid } from "guid-typescript";
import { isUndefined } from "util";
import { GraphClient } from "../helpers/graph-client";
import { Middleware } from "../helpers/middleware";
import { TeamsHelper } from '../helpers/teams-helper';
import { OptionsResponseData } from "../models/options-response-model";

// Helper function for generating new GUIDs
function createGuid() {
    return Guid.create().toString();
}

// Routes for IFTTT Triggers and Actions endpoints related to OneNote
export class IftttRouterTeams {

    private _router;

    // setup router and add all routes
    constructor() {
        this._router = express.Router();

        // actions
        this._router.post('/actions/create_team', Middleware.authorizationCheck, Middleware.actionFieldsCheck, this.postAction_CreateTeam);
        this._router.post('/actions/create_channel', Middleware.authorizationCheck, Middleware.actionFieldsCheck, this.postAction_CreateChannel);
        this._router.post('/actions/create_channel/fields/team_id/options', Middleware.authorizationCheck, this.getParam_Team);
        this._router.post('/actions/create_channel/fields/channel_id/options', Middleware.authorizationCheck, this.getParam_Channel);
        this._router.post('/actions/create_message', Middleware.authorizationCheck, Middleware.actionFieldsCheck, this.postAction_CreateMessage);
        this._router.post('/actions/create_message/fields/team_id/options', Middleware.authorizationCheck, this.getParam_Team);
    }

    public getRouter() {
        return this._router;
    }

    // Action endpoint: Create new Teams team
    private async postAction_CreateTeam(req, res) {

        const teamName = req.body.actionFields.team_name;

        if (teamName) {

            const authHeader = req.headers.authorization;
            const teamId = await TeamsHelper.createTeamAsync(authHeader, teamName).catch(e=>{
                res.status(401).json({
                    errors: [{
                        message: `error attempting to create or get team from graph : ${e}`
                    }]
                });
            });

            if (teamId) {
                res.status(200).json({
                    data: [{
                        id: teamId
                    }]
                });
            }
        }
        else {

            res.status(400).json({
                errors: [{
                    message: "incomplete data sent, please supply a team_name field"
                }]
            });
        }
    }

    // Action endpoint: Create a new Teams channel
    private async postAction_CreateChannel(req, res) {
        const authHeader = req.headers.authorization;
        const teamName = req.body.actionFields.team_name;

        try {
            const teamId = await TeamsHelper.createTeamAsync(authHeader, teamName).catch(e=>{
                return Promise.reject(`error attempting to create or get team from graph : ${e}`);
            });
            const channelName = req.body.actionFields.channel_name;

            if (teamId && channelName) {

                const channelId = await TeamsHelper.createChannelAsync(authHeader, teamId, channelName).catch(e=>{
                    return Promise.reject(`error attempting to create team : ${e}`);
                });

                if (channelId){
                    res.status(200).json({
                        data: [{
                            id: channelId
                        }]
                    });
                }
            }
            else {

                res.status(400).json({
                    errors: [{
                        message: "incomplete data sent, please supply team_name and channel_name fields"
                    }]
                });
            }
        } catch (e) {
            res.status(401).json({
                errors: [{
                    message: `Error creating channel: ${e.message}`
                }]
            });
        }
    }

    // Action endpoint: Create a new Teams message in existing or created Team/Channel
    private async postAction_CreateMessage(req, res) {
        const authHeader = req.headers.authorization;
        let teamId = req.body.actionFields.team_id;
        const teamName = req.body.actionFields.team_name;
        let channelId = req.body.actionFields.channel_id;
        const channelName = req.body.actionFields.channel_name;
        const message = req.body.actionFields.message;

        if ([teamId,channelName,message].some(i=> isUndefined(i)))
        {
            /// IFTTT requires this for tests... otherwise it would be ok to have the teamId missing
            res.status(400).json({
                errors: [{
                    message: "incomplete data sent, please supply team_name, channel_name, and message fields"
                }]
            });
            return;
        }

        /// Test setup uses a string for the null items
        if (teamId === "null"){
            teamId = null;
        }
        if (channelId === "null")
        {
            channelId = null;
        }

        try{
            if (!teamId)
            {
                teamId = await TeamsHelper.createTeamAsync(authHeader, teamName).catch(e=>{
                    return Promise.reject(`error attempting to create or get team from graph : ${e}`);
                });
            }
            if (teamId && !channelId)
            {
                channelId = await TeamsHelper.createChannelAsync(authHeader, teamId, channelName).catch(e=>{
                    Promise.reject(`error attempting to create or get channel from graph : ${e}`);
                });
            }
        }
        catch (e)
        {
            res.status(401).json({
                errors: [{
                    message: `${e}`
                }]
            });
            return;
        }

        if (teamId && channelId && message) {
            try {
                await TeamsHelper.createMessageAsync(authHeader, teamId, channelId, message);

                res.status(200).json({
                    data: [{
                        id: createGuid()
                    }]
                });
            }
            catch(e)
            {
                res.status(401).json({
                    errors: [{
                        message: `Error trying to send message: ${e}`
                    }]
                });
                return;
            }
        }
        else {
            res.status(400).json({
                errors: [{
                    message: "incomplete data sent, please supply team_name, channel_name, and message fields"
                }]
            });
        }
    }

    private async getParam_Team(req, res) {
        const authHeader = req.headers.authorization;

        // get the Teams joined by the user
        const url = `/me/joinedTeams`;
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

    private async getParam_Channel(req, res) {
        const authHeader = req.headers.authorization;
        const team_id = req.headers.team_id;

        // get the Teams joined by the user
        const url = `/teams/${team_id}/channels`;
        const client:Client = GraphClient.getClient(authHeader);
        const request = client.api(url);
        let graphItems:Channel[];

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
