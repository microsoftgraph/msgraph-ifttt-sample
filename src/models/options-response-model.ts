/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Group } from "@microsoft/microsoft-graph-types";

/**
 * Creates the response body of any IFTTT Options item used
 *      when creating a dropdown in a trigger or action on the
 *      IFTTT website in settings when a user sets up an applet.
 */
export class OptionsResponseData {
    data: Option[];

    constructor(options:Option[])
    {
        this.data = options;
    }

    // Converts from the graph json format to the IFTTT response format
    static fromGraph(teams:Group[]): OptionsResponseData{
        const opts: Option[] = [];
        for ( const team of teams)
        {
            const opt : Option = {
                label: `${team.displayName}`,
                value: `${team.id}`
            }
            opts.push(opt);
        }
        opts.sort((a,b)=> a.label.localeCompare(b.label));
        return new OptionsResponseData(opts);
    }
}

/** Set of options for parameters,
 *  IFTTT can except multiple levels of options returned
 *  this is supported by nesting a Value array within the option object
 */
export interface Option {
    label:   string;
    value?:  string;
    values?: Value[];
}

/** single (usual) option in a parameter dropdown id and value */
export interface Value {
    label: string;
    value: string;
}
