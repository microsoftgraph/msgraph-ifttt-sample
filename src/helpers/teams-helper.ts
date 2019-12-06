/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import { Channel, Group, Message, Team } from "@microsoft/microsoft-graph-types";
import { GraphClient } from "./graph-client";

// Common functions for interaction with Microsoft Teams
export class TeamsHelper {

    // Get team id by team name
    public static async getTeamIdAsync(authToken: string, teamName: string) {

        const url = `/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=displayName,id`;
        const teams:Group[] = await GraphClient.getClient(authToken, 'beta').api(url).get()
            .then(v=>{return v.value})
            .catch(e=>{
                console.log(e);
                return Promise.reject(e);
            });

        const foundTeam = teams.find((team)=> team.displayName === teamName);

        if (foundTeam) {

            console.log(`team found ${foundTeam.id}`);
            return foundTeam.id;
        }
        else {

            console.log(`team not found`);
            return null;
        }
    }

    // Get a channel id by team id and channel name
    public static async getChannelIdAsync(authToken: string, teamId: string, channelName: string) {

        const url = `/teams/${teamId}/channels?$filter=displayName eq '${channelName}'&$select=displayName,id`;
        const channels:Channel[] = await GraphClient.getClient(authToken, 'beta').api(url).get()
            .then(v=>{return v.value})
            .catch(e=>{
                console.log(e);
                return Promise.reject(e);
            });

        const foundChannel = channels.find((channel)=> channel.displayName === channelName);

        if (foundChannel) {

            console.log(`channel found ${foundChannel.id}`);
            return foundChannel.id;
        }
        else {

            console.log(`channel not found`);
            return null;
        }
    }

    // Create a new Teams team
    public static async createTeamAsync(authToken:string, teamName:string) {

        // Check if it already exists by that name.
        let teamId = await this.getTeamIdAsync(authToken, teamName).catch(e=>{
            return Promise.reject(e);
        });

        // if not exists, try to create.
        if (!teamId) {
            const url = `/teams`;
            const newTeam = {
                "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
                "displayName": teamName,
                "description": "Team to handle all IFTTT notifications and communications"
            };

            const client:Client = GraphClient.getClient(authToken,'beta');
            let teamCreated:Response;
            try {
                teamCreated = await client.api(url).post(newTeam);

                if (teamCreated !== undefined){
                    teamId = await this.getTeamIdAsync(authToken, teamName);
                }
            } catch (error) {
                console.log(error);
                return Promise.reject(error);
            }
        }

        return teamId;
    }

    // Create a new Teams channel
    public static async createChannelAsync(authToken:string, teamId:string, channelName: string) {

        // Check if it already exists by that name.
        let channelId = await this.getChannelIdAsync(authToken, teamId, channelName).catch(e=>{
            return Promise.reject(e);
        });

        // if not exists, try to create.
        if (!channelId) {
            const url = `/teams/${teamId}/channels/`;
            const newChannel:Channel = {
                "displayName": `${channelName}`,
                "description": "Channel to handle all IFTTT notifications and communications"
            };
            const client:Client = GraphClient.getClient(authToken,'beta');
            let channelCreated:Channel;
            try {
                channelCreated = await client.api(url).post(newChannel).catch(e=>{
                    return Promise.reject(`Error creating channel: ${e}`);
                });
            } catch (e) {
                console.log(e);
                return Promise.reject(e);
            }

            channelId = channelCreated.id;
            console.log(`Channel created: ${channelId}`);
        }

        return channelId;
    }

    // Create a new Teams message
    public static async createMessageAsync(authToken:string, teamId:string, channelId:string, message:string) {
        const url = `/teams/${teamId}/channels/${channelId}/messages`;
        const msg:Message = {
            body: {
                content: message
            }
        }
        const client:Client = GraphClient.getClient(authToken,'beta');
        let msgCreated:Message;
        try {
            msgCreated = await client.api(url).post(msg).catch(e=>{
                return Promise.reject(`Error creating message: ${e}`);
            });
        } catch (e) {
            console.log(e);
            return Promise.reject(e);
        }

        if (msgCreated) {
            console.log("Message sent");
        } else {
            return Promise.reject("Could not send message");
        }
    }
}
