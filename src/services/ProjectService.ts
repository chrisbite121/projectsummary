import {
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions,
    ISPHttpClientBatchOptions,
    ISPHttpClientBatchCreationOptions,
    SPHttpClientBatch
} from "@microsoft/sp-http";

export interface IWebProperty {
    [property: string]: any
}

import { IFieldItem } from '../models'

export class ProjectService {
    private _spHttpOptions: any = {
        getNoMetadata: <ISPHttpClientOptions>{
            headers: { 'ACCEPT': 'application/json; odata.metadata=none' }
        },
        useODataV3: <ISPHttpClientOptions>{
            headers: { 'OData-Version': '3.0' }
        }
    }

    constructor(private siteAbsoluteUrl: string, private client: SPHttpClient) { }

    //     private _makeSPHttpClientBatchRequest(): void {

    //         // Here, 'this' refers to my SPFx webpart which inherits from the BaseClientSideWebPart class.
    //         // Since I am calling this method from inside the class, I have access to 'this'.
    //         const spHttpClient: SPHttpClient = this.context.spHttpClient;
    //         const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;

    //         const spBatchCreationOpts: ISPHttpClientBatchCreationOptions = { webUrl: currentWebUrl };

    //         const spBatch: SPHttpClientBatch = spHttpClient.beginBatch(spBatchCreationOpts);

    //         // Queue a request to get current user's userprofile properties
    //         const getMyProperties: Promise<SPHttpClientResponse> = spBatch.get(`${currentWebUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClientBatch.configurations.v1);

    //         // Queue a request to get the title of the current web
    //         const getWebTitle: Promise<SPHttpClientResponse> = spBatch.get(`${currentWebUrl}/_api/web/title`, SPHttpClientBatch.configurations.v1);

    //         // Queue a request to create a list in the current web.
    //         const currentTime: string = new Date().toString();
    //         const batchOps: ISPHttpClientBatchOptions = {
    //           body: `{ Title: 'List created with SPFx batching at ${currentTime}', BaseTemplate: 100 }`
    //         };
    //         const createList: Promise<SPHttpClientResponse> = spBatch.post(`${currentWebUrl}/_api/web/lists`, SPHttpClientBatch.configurations.v1, batchOps);


    //         spBatch.execute().then(() => {

    //           getMyProperties.then((response: SPHttpClientResponse) => {
    //             response.json().then((props: any) => {
    //               console.log(props);
    //             });
    //           });

    //           getWebTitle.then((response: SPHttpClientResponse) => {

    //             response.json().then((webTitle: string) => {

    //               console.log(webTitle);
    //             });
    //           });

    //           createList.then((response: SPHttpClientResponse) => {

    //             response.json().then((responseJSON: any) => {

    //               console.log(responseJSON);
    //             });
    //           });
    //         });
    //       }








    getWebData(): Promise<object> {
        let query = `${this.siteAbsoluteUrl}/_api/web/allproperties`
        let promise: Promise<object> = new Promise<object>((resolve, reject) => {
            this.client.get(
                query,
                SPHttpClient.configurations.v1
            )
                .then((response: SPHttpClientResponse): Promise<object> => {
                    return response.json()
                })
                .then((response: object) => {
                    if (response.hasOwnProperty('PWAURL') &&
                        response['PWAURL'].length > 0 &&
                        response.hasOwnProperty('MSPWAPROJUID') &&
                        response['MSPWAPROJUID'].length > 0) {
                        return { url: response['PWAURL'], id: response['MSPWAPROJUID'] }
                    } else {
                        reject('project uid and/or project id lookup failed')
                    }


                })
                .then((response: { url: string, id: string }) => {
                    resolve(response)
                })
                .catch((err: any) => {
                    reject(err)
                })
        })
        return promise
    }

    getProjectFields(projectUrl, projectId): Promise<any> {
        let query = `${projectUrl}/_api/ProjectServer/Projects(guid'${projectId}')/IncludeCustomFields`
        let promise: Promise<object> = new Promise<object>((resolve, reject) => {
            this.client.get(
                query,
                SPHttpClient.configurations.v1
            )
                .then((response: SPHttpClientResponse): Promise<object> => {
                    return response.json()
                })
                .then((response) => {
                    resolve(response)
                })
                .catch((err: any) => {
                    reject(err)
                })
        })

        return promise
    }

    getCustomFields(projectUrl, projectId): Promise<any> {
        let query = `${projectUrl}/_api/projectserver/CustomFields`
        let promise: Promise<object> = new Promise<object>((resolve, reject) => {
            this.client.get(
                query,
                SPHttpClient.configurations.v1
            )
                .then((response: SPHttpClientResponse): Promise<object> => {
                    return response.json()
                })
                .then((response) => {
                    resolve(response)
                })
                .catch((err: any) => {
                    reject(err)
                })
        })

        return promise
    }

    getLookupTables(projectUrl, projectId): Promise<any> {
        let query = `${projectUrl}/_api/projectserver/LookupTables`
        let promise: Promise<object> = new Promise<object>((resolve, reject) => {
            this.client.get(
                query,
                SPHttpClient.configurations.v1
            )
                .then((response: SPHttpClientResponse): Promise<object> => {
                    return response.json()
                })
                .then((response) => {
                    resolve(response)
                })
                .catch((err: any) => {
                    reject(err)
                })
        })

        return promise
    }








    getProjectDataBatch(projectUrl, projectId) {
        const spBatchCreationOpts: ISPHttpClientBatchCreationOptions = { webUrl: projectUrl };
        const spBatch: SPHttpClientBatch = this.client.beginBatch();

        // Queue a request to get current user's userprofile properties
        const getProjectFields: Promise<SPHttpClientResponse> = spBatch.get(`${projectUrl}/_api/ProjectServer/Projects(guid'${projectId}')/IncludeCustomFields`, SPHttpClientBatch.configurations.v1, this._spHttpOptions.useODataV3);
        const getCustomFields: Promise<SPHttpClientResponse> = spBatch.get(`${projectUrl}/_api/projectserver/CustomFields`, SPHttpClientBatch.configurations.v1, this._spHttpOptions.useODataV3);
        const getLookupTables: Promise<SPHttpClientResponse> = spBatch.get(`${projectUrl}/_api/projectserver/LookupTables`, SPHttpClientBatch.configurations.v1, this._spHttpOptions.useODataV3);

        spBatch.execute().then(() => {

            getProjectFields
                .then((response: SPHttpClientResponse) => {
                    return response.json()
                })
                .then((props: any) => {
                    console.log(props);
                });


            getCustomFields.then((response: SPHttpClientResponse) => {
                response.json().then((props: any) => {
                    console.log(props);
                });
            })

            getLookupTables.then((response: SPHttpClientResponse) => {
                response.json().then((props: any) => {
                    console.log(props);
                });
            })
        })
    }



    // let query2 = `${projectUrl}/_api/ProjectServer/Projects(guid'${projectId}')`
    // let query3 = `${projectUrl}/_api/projectserver/CustomFields`
    // let query4 = `${projectUrl}/_api/projectserver/LookupTables`



    // let promise = Promise.all([
    //     // this.client.get(query2, SPHttpClient.configurations.v1, { headers: { 'OData-Version': '3.0' } })])
    //     this.client.get(query2, SPHttpClient.configurations.v1, this._spHttpOptions.getNoMetadata),
    //     this.client.get(query3, SPHttpClient.configurations.v1, this._spHttpOptions.getNoMetadata),
    //     this.client.get(query4, SPHttpClient.configurations.v1, this._spHttpOptions.getNoMetadata)
    // ])
    //     .then((response: Array<any>) => {
    //         console.log(response[0].json())
    //         console.log(response[1].json())
    //         console.log(response[2].json())
    //     })
    //     .then((response: any) => {

    //     })


    processCustomFieldNames(fields: object, customFields: Array<object>): object {
        let re = /^Custom_x005f_(.*)/
        Object.keys(fields).forEach(element => {
            //match fields that start with 're'
            if (element.match(re)) {
                let fieldData: object = this.findCustomFieldData(customFields, element.match(re)[1])
                let fieldName = fieldData['fieldName']
                let guid = fieldData['guid']
                //create a new property using the matched field name name and use the same property descriptor
                Object.defineProperty(fields, fieldName, Object.getOwnPropertyDescriptor(fields, element));
                //if lookuptable add field guid
                if (Array.isArray(fields[fieldName])) {
                    fields[fieldName].push(guid)
                }
                //delete old entry
                delete fields[element];
            }
        });
        return fields
    }

    findCustomFieldData(customFields: Array<object>, guid: string): object {
        let _fieldName = guid
        let _guid = guid
        customFields.forEach((obj, i) => {
            let re = /^Custom_(.*)/
            if (obj.hasOwnProperty('InternalName')) {
                let _fieldId: string = obj['InternalName'].match(re)[1]
                if (obj.hasOwnProperty('Id') && _fieldId == guid) {
                    obj.hasOwnProperty('Name') ?
                        //replace spaces so that it resembles internal name
                        _fieldName = obj['Name'].replace(' ', '') :
                        _fieldName = guid

                    obj.hasOwnProperty('Id') ?
                        _guid = obj['Id'] :
                        _guid = guid
                }
            }
        })
        return { fieldName: _fieldName, guid: _guid }
    }

    getFieldValues(fieldNames: Array<IFieldItem>, projectData: object): Array<IFieldItem> {
        fieldNames.forEach((fieldItem, i) => {
            if (projectData.hasOwnProperty(fieldItem.internalName)) {
                fieldNames[i].fieldValue = projectData[fieldItem.internalName]
            }
        })

        return fieldNames
    }

    getLookupEntryValues(projectUrl: string, fieldId): Promise<any> {
        let query = `${projectUrl}/_api/projectserver/CustomFields(guid'${fieldId}')/LookupEntries`
        let promise: Promise<object> = new Promise<object>((resolve, reject) => {
            this.client.get(
                query,
                SPHttpClient.configurations.v1
            )
                .then((response: SPHttpClientResponse): Promise<object> => {
                    return response.json()
                })
                .then((response) => {
                    console.log(response);
                    resolve(response)
                })
                .catch((err: any) => {
                    reject(err)
                })
        })

        return promise

    }

    getLookupEntry(entries, fieldData: { fieldName: string, fieldValueGuid: string }): string {
        console.log(entries);
        let _value = ''
        entries.forEach((entry, i) => {
            if (entry.hasOwnProperty('InternalName') && entry['InternalName'] == fieldData.fieldValueGuid) {
                entry.hasOwnProperty('Value') ?
                    _value = entry['Value'] :
                    _value = ''
            }
        })
        return _value
    }


}