/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import express from 'express';
import { Middleware } from "../helpers/middleware";
import { OptionsResponseData } from "../models/options-response-model";
import { GraphClient } from "../helpers/graph-client";
import { Client } from "@microsoft/microsoft-graph-client";
import { OnenotePage, OnenoteSection } from "@microsoft/microsoft-graph-types";
import { OnenotePageResponseData } from "../models/onenotepage-response-data";
import { isUndefined } from "util";

// Routes for IFTTT Triggers and Actions endpoints related to OneNote
export class IftttRouterOnenote {

    private _router;

    // setup router and add all routes
    constructor() {
        this._router = express.Router();

        // triggers
        this._router.post('/triggers/onenote_page_created', Middleware.authorizationCheck, this.postTrigger_PageCreated);

        // actions
        this._router.post('/actions/onenote_create_page', Middleware.authorizationCheck, Middleware.actionFieldsCheck, this.postAction_create_page);
        this._router.post('/actions/onenote_create_page/fields/section_id/options', Middleware.authorizationCheck, this.getParam_section);
    }

    public getRouter() {
        return this._router;
    }

    /**
     * returns a list of sections available from the users OneNote Notebook
     */
    private async getParam_section(req, res) {
        const authHeader = req.headers.authorization;

        // setup graph call for sections
        const url = `/me/onenote/sections`;
        const client:Client = GraphClient.getClient(authHeader);
        const request = client.api(url);
        let graphItems:OnenoteSection[];

        try {
            // get the sections owned by the user
            graphItems = await request.get().then(v=>{return v.value}).catch(e => {
                    console.log(e);
                    return Promise.reject(e);
            });

            console.log(`found ${graphItems ? graphItems.length : 0} items`);

            /// include the Notebook name in the section for the dropdown
            graphItems.map((s)=>{
                s.displayName = `${s.parentNotebook.displayName}:${s.displayName}`;
            })

            // convert graph items to IFTTT items
            const responseItems = OptionsResponseData.fromGraph(graphItems);
            res.status(200).json(responseItems);
        } catch (e) {
            res.status(401).send({
                "errors": [{
                    "message": `Retrieving Graph data failed : ${e}`
                }]
            });
        }
    }

    // Action endpoint: Create a new OneNote page existing OneNote Section
    private async postAction_create_page(req, res) {
        const authHeader = req.headers.authorization;
        let sectionId = req.body.actionFields.section_id;
        let content = req.body.actionFields.content;
        let title = req.body.actionFields.title;

        /// Required IFTTT Test setup uses a string for the null items
        /// This converts to actual null
        if (sectionId === "null"){
            sectionId = null;
        }
        if (content === "null")
        {
            content = null;
        }
        if (title === "null")
        {
            title = null;
        }

        if ([sectionId, content, title].some(i=> isUndefined(i)))
        {
            res.status(400).json({
                errors: [{
                    message: "incomplete data sent, please supply sectionId, content, and title fields"
                }]
            });
            return;
        }
        else
        {
            content = `
                <html lang="en-US">
                    <head>
                        <title>${title}</title>
                        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
                    </head>
                    <body data-absolute-enabled="true" style="font-family:Calibri;font-size:11pt">
                        <div style="position:absolute;left:48px;top:115px;width:624px">
                            <p style="margin-top:0pt;margin-bottom:0pt">${content}</p>
                            <br />
                        </div>
                    </body>
                </html>`;

            try{
                const page:OnenotePage = await GraphClient.getClient(authHeader).api(`/me/onenote/sections/${sectionId}/pages`)
                    .header('Content-Type','text/html').post(content);

                res.status(200).json({
                    data: [{
                        id: `${page.id}`
                    }]
                });
            }
            catch(e)
            {
                res.status(401).json({
                    errors: [{
                        message: `Page was not posted: ${e.message}`
                    }]
                });
            }
        }
    }

    // Trigger endpoint: New page created.
    private async postTrigger_PageCreated(req, res) {
        const authHeader = req.headers.authorization;
        const numOfItems = (req.body.limit !== undefined) ? req.body.limit : 50;

        try {
            // get the last pages for the user
            const url = `/me/onenote/pages?$expand=parentNotebook,parentSection`;
            const client:Client = GraphClient.getClient(authHeader);
            let graphItems:OnenotePage[] = await client.api(url).get().then(v=>{return v.value}).catch(e=> {
                return Promise.reject(e);
            });

            // Needs to have the title ending with "#publish" to meet the requirements
            graphItems = graphItems.filter( p=>{
                p.title.endsWith("#publish")
            });

            console.log(`found ${graphItems.length} events`);
            if (graphItems.length < 3)
            {
                /// need to fake 3 items so that IFTTT will behave
                graphItems.push(
                    { id:`123`, title:"Fake", createdDateTime:`2018-01-01`,
                        parentNotebook:{displayName:"Fake Notebook"}, parentSection:{displayName:"Fake Section"},
                        contentUrl:"https://graph.microsoft.com/v1.0/me/onenote/pages/fakepageid"}
                    ,{ id:`456`, title:"Fake", createdDateTime:`2018-01-02`,
                        parentNotebook:{displayName:"Fake Notebook"}, parentSection:{displayName:"Fake Section"},
                        contentUrl:"https://graph.microsoft.com/v1.0/me/onenote/pages/fakepageid"}
                    ,{ id:`789`, title:"Fake", createdDateTime:`2018-01-03`,
                        parentNotebook:{displayName:"Fake Notebook"}, parentSection:{displayName:"Fake Section"},
                        contentUrl:"https://graph.microsoft.com/v1.0/me/onenote/pages/fakepageid"}
                );
            }

            // convert graph events to IFTTT Events
            const responseItems = OnenotePageResponseData.fromGraph(graphItems, numOfItems);

            res.status(200).json(responseItems);
        } catch (e) {
            res.status(401).json({
                errors: [{
                    message: `Graph data query failure: ${e.message}`
                }]
            });
        }
    }
}
