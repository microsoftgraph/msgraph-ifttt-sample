/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import fetch from 'node-fetch';
const IFTTT_KEY = process.env.IFTTT_KEY;

// Common functions for processing secure routes
export class Middleware {

    public static actionFieldsCheck(req, res, next) {

        if (!req.body || req.body.actionFields === undefined) {

            res.status(400).json({
                errors: [{
                    message: 'Incomplete data sent, please supply actionFields'
                }]
            });
        }
        else {
            next();
        }
    }

    public static triggerFieldsCheck(req, res, next) {

        if (!req.body || req.body.triggerFields === undefined) {

            res.status(400).json({
                errors: [{
                    message: 'Incomplete data sent, please supply triggerFields'
                }]
            });
        }
        else {
            next();
        }
    }

    // Validate the IFTTT service key value
    public static serviceKeyCheck (req, res, next) {

        console.log(req.originalUrl);
        const key = req.get("IFTTT-Service-Key");

        if (key !== IFTTT_KEY) {

            res.status(401).send({
                errors: [{
                  message: "Channel/Service key is not correct"
                }]
            });

            next('router');
        }
        else {

          next();
        }
    }

    // Validate the graph auth token in the authorization header
    public static async authorizationCheck(req, res, next)
    {
        const authHeader = req.headers.authorization;

        // Simply check auth by using the supplied key to query your own data.
        const url = "https://graph.microsoft.com/v1.0/me";
        const graphResponse = await fetch(url, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': authHeader
            }
        });

        if (graphResponse.ok) {

            next();
        }
        else {

            res.status(401).send({
              "errors": [{
                "message": `Authorization key is not correct : ${authHeader}`
              }]
            });

            next('router');
        }
    }
};