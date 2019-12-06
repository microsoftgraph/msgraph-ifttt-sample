/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import express from 'express';

// Routes for non-ifttt api endpoints
export class ApiRoutes {

    private _router;
    public getRouter() {
        return this._router;
    }

    constructor() {

        this._router = express.Router();

        this._router.get('/', this.getDefault);
        this._router.get('/auth', this.getAuth);
    }

    // Default view
    private getDefault(req, res) {

        res.render('index.ejs');
    }

    // Auth redirect endpoint
    private getAuth(req, res) {

        res.render('index.ejs');
    }
}
