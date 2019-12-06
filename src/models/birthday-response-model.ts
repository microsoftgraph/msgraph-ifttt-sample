import { TriggerItem, TriggerResponseData } from "./trigger-response-model";
import { User } from "@microsoft/microsoft-graph-types";

/** Creates the response body of the Birthday Event IFTTT Trigger
 *      automatically will set the json payload correctly
 *      with fomatted BirthdayTriggerItem data
 */
export class BirthdayResponseData implements TriggerResponseData<BirthdayTriggerItem> {

    data: BirthdayTriggerItem[];

    constructor (data : BirthdayTriggerItem[]) {
        this.data = data;
    }

    // Converts from the graph json format to the IFTTT response format
    static fromGraph(data:User[], numOfItems:number) : BirthdayResponseData {

        if (!data || numOfItems === 0) {
            return new BirthdayResponseData([]);
        }

        // convert the items taking only the properties required.
        let responseData = data.map(
            (dataItem) => {
                return new BirthdayTriggerItem(dataItem.id
                    , Math.floor(Date.parse(dataItem.birthday)/1000)
                    , dataItem.displayName
                    , dataItem.birthday);
            }
        )

        // Sort the items by timestamp desc (IFTTT requirement)
        responseData.sort((a, b) => b.meta.timestamp - a.meta.timestamp);

        // Filter items to only the requested max amount from IFTTT request
        if (responseData.length > numOfItems)
        {
            responseData = responseData.slice(0, numOfItems);
        }
        console.log(`filtered to ${responseData.length} items`);

        return new BirthdayResponseData(responseData);
    }
}

/** Creates birthday event item data formatted
 *      as required by IFTTT to support the new event triggers
 *      forces the IFTTT requirement of meta data id and timestamp
 *      as well as required sorting and filtering
 */
class BirthdayTriggerItem extends TriggerItem {
    // tslint:disable: variable-name
    birthday:string;
    display_name:string;
    // tslint:enable: variable-name

    constructor(id:string, timestamp:number, displayName:string, birthday:string)
    {
        super(id,timestamp);
        this.display_name = displayName;
        this.birthday = birthday;
    }
}

