/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// imports
import express from 'express';
import bodyParser from 'body-parser';
import path from 'path';
import dotenv from 'dotenv';
// setup .env config file (See Readme for more)
dotenv.config();

// express
const app = express();

// express config
app.use(bodyParser.json());
const options = {
    dotfiles: 'ignore',
    etag: false,
    extensions: ['htm', 'html', 'svg', 'PNG', 'png', 'ttf'],
    index: false,
    maxAge: '1d',
    redirect: false,
    // tslint:disable-next-line: no-shadowed-variable
    setHeaders (res, path, stat) {
      res.set('x-timestamp', Date.now())
    }
  }
app.use('/static', express.static('../public', options));

// express views
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '../views'));

/**
 * express routers separate concerns of responsibility.
 *  Most IFTTT items handled in default ifttt-routes.
 *  Others are separated by Graph provider (ex: OneNote)
 * ApiRoutes handles all UI based content.
 */
import { IftttRoutes } from './routes/ifttt-routes';
import { IftttRouterOnenote } from "./routes/ifttt-routes-onenote";
import { IftttRouterTeams } from "./routes/ifttt-routes-teams";
import { ApiRoutes } from './routes/api-routes';

app.use('/ifttt/v1', new IftttRoutes().getRouter());
app.use('/ifttt/v1', new IftttRouterOnenote().getRouter());
app.use('/ifttt/v1', new IftttRouterTeams().getRouter());
app.use('/', new ApiRoutes().getRouter());

// start the server
const port = process.env.PORT;
app.listen(port, ()=> {

    console.log(`Listening on port ${port}`);

    const IFTTT_KEY = process.env.IFTTT_KEY;
    console.log(`Service Key: ${IFTTT_KEY}`);
});