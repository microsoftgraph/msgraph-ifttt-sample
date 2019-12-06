/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MSGraph from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch"; // required for graph client

/*
Example of usage:
    private async getGraphAuth(req,res){
        let authHeader:string = req.headers['authorization'];
        let client:Client = GraphClient.getClient(authHeader);

        let reply = await client.api("/me").get().catch(e=>{console.log(e)});
        res.status(200).json(reply);
    }
*/

/**
 * Provides a pass-through client for Microsoft Graph. Using an authentication key
 * gained from another source (in this case IFTTT), provide all the functionality of
 * a graph client supplied in the official Graph SDK.
 */
export class GraphClient{
    static getClient(authHeader:string, version:string="v1.0"): MSGraph.Client
    {
        authHeader = authHeader.replace('Bearer ','');
        const authProvider:IFTTTAuthProvider = new IFTTTAuthProvider(authHeader);
        const options:MSGraph.ClientOptions = {
            authProvider,
            defaultVersion:version
        }

        return MSGraph.Client.initWithMiddleware(options);
    }
}

/**
 * Creates the required (by Graph SDK) Auth provider
 * using a authentication key passed from the IFTTT service.
 */
class IFTTTAuthProvider implements MSGraph.AuthenticationProvider{
    authToken:string = "";

    constructor(authToken:string){
        this.authToken = authToken;
    }
    async getAccessToken(authenticationProviderOptions?: MSGraph.AuthenticationProviderOptions): Promise<string>{
        return this.authToken;
    }
}