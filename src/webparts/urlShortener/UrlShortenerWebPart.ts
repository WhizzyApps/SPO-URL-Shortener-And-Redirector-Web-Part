import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneLink, PropertyPaneButton, PropertyPaneButtonType, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

import UrlShortener from './components/UrlShortener';
import { IUrlShortenerProps } from './components/IUrlShortenerProps';

const packageSolution = require("../../../config/package-solution.json");
const timeStampFile = require("../../../config/timeStamp.json");
const buildTimeStamp = timeStampFile["buildTimeStamp"];

export interface IUrlShortenerWebPartProps {
    idLength: number;
    lookupList: {};
    newListName: string;
}

export default class UrlShortenerWebPart extends BaseClientSideWebPart<IUrlShortenerWebPartProps> 
{
    private isOwner = false;
    private currentUserRole = "";  
    private lookupListsOptions = [];

    // get data from SharePoint REST API
    private async _spApiGet (url: string): Promise<object> {
  
        const clientOptions: ISPHttpClientOptions = {
            headers: new Headers(),
            method: 'GET',
            mode: 'cors'
        };
        try {
            const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, clientOptions);
            //Since the fetch API only throws errors on network failure, We have to explicitly check for 404's etc.
            const responseJson = await response.json();
    
            responseJson['status'] = response.status;
    
            if (!responseJson.value) {
                responseJson['value'] = [];
            }
            return responseJson;
        } 
        catch (error) {
            return error;
        }
    } 

    private async _spApiPost (url: string, data): Promise<object> 
    {
        const clientOptions: ISPHttpClientOptions = {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': '',
            },
            body: data,
            mode: 'cors'
        };
        const response = await this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + url, SPHttpClient.configurations.v1, clientOptions);
        const responseJson = await response.json();
        responseJson['status'] = response.status;
        if (!responseJson.value) {
            responseJson['value'] = [];
        }
        return responseJson;
    } 

    // get permissions of current user
    private async _getUserPermissions () {
        let currentUserRole;
        // get permission of current users
        const urlAdmin = this.context.pageContext.web.absoluteUrl + '/_api/web/currentuser/isSiteAdmin';
        const isSiteAdminResponse = await this._spApiGet(urlAdmin);
        const isSiteAdmin = isSiteAdminResponse['value'];
        if (isSiteAdmin == true) {
            currentUserRole = "admin";
        } 
        else {
            const urlPerm = this.context.pageContext.web.absoluteUrl + `/_api/web/effectiveBasePermissions`;
            const permResponse = await this._spApiGet(urlPerm);
            let permArray = [];
            if (permResponse["Low"]) {
                 permArray = this._convertUserPermissions (permResponse['Low'], permResponse['High']);
                if (permArray.includes("ManagePermissions")) {
                    currentUserRole = "owner";
                } else if (permArray.includes("EditListItems")) {
                    currentUserRole = "member";
                } else {
                    currentUserRole = "visitor";
                }
            }
            else {
                currentUserRole = "visitor";
            }
    
        }
        return currentUserRole;
    }

    private async _getLists () 
    {
        const url = this.context.pageContext.web.absoluteUrl + '/_api/web/lists?$select=Id,Title&$filter=((Hidden%20eq%20false)and(BaseTemplate%20eq%20100))';
        const response = await this._spApiGet(url);
        if (response["status"] && response["status"].toString().startsWith("2"))
        {
            this.lookupListsOptions = response["value"].map(list => {
                return {key: list["Id"], text: list["Title"] };
            });
        }
        else {
            // ToDo Error message
        }
    }

    // convert permissions of current user
    private _convertUserPermissions (lowPermDec, highPermDec) {
        let Flags = {
        Low: [
            // Lists and Documents
            { EmptyMask: 0 },
            { ViewListItems: 1 << 0 },
            { AddListItems: 1 << 1 },
            { EditListItems: 1 << 2 },
            { DeleteListItems: 1 << 3 },
            { ApproveItems: 1 << 4 },
            { OpenItems: 1 << 5 },
            { ViewVersions: 1 << 6 },
            { DeleteVersions: 1 << 7 },
            { OverrideListBehaviors: 1 << 8 },
            { ManagePersonalViews: 1 << 9 },
            { ManageLists: 1 << 11 },
            { ViewApplicationPages: 1 << 12 },

            // Web Level
            { Open: 1 << 16 },
            { ViewPages: 1 << 17 },
            { AddAndCustomizePages: 1 << 18 },
            { ApplyThemAndBorder: 1 << 19 },
            { ApplyStyleSheets: 1 << 20 },
            { ViewAnalyticsData: 1 << 21 },
            { UseSSCSiteCreation: 1 << 22 },
            { CreateSubsite: 1 << 23 },
            { CreateGroups: 1 << 24 },
            { ManagePermissions: 1 << 25 },
            { BrowseDirectories: 1 << 26 },
            { BrowseUserInfo: 1 << 27 },
            { AddDelPrivateWebParts: 1 << 28 },
            { UpdatePersonalWebParts: 1 << 29 },
            { ManageWeb: 1 << 30 }
        ],
        High: [
            // High Bits
            { UseClientIntegration: 1 << 4 },
            { UseRemoteInterfaces: 1 << 5 },
            { ManageAlerts: 1 << 6 },
            { CreateAlerts: 1 << 7 },
            { EditPersonalUserInformation: 1 << 8 },

            // Special Permissions
            { EnumeratePermissions: 1 << 30 }
            //FullMask          :   2147483647 // Invisible in WebUI, not useful since it's always true when &'ed
        ]
        };
        // low permissions
        // Permissions: make array of objects
        const flagsLow = Flags.Low.map((objectItem) => {
            const code = (Object.values(objectItem)[0] >>> 0).toString(2);
            return { name: Object.keys(objectItem)[0], code: code };
        });
        let zeros = "";
        const permLowArr = (lowPermDec >>> 0)
        .toString(2)
        .split("")
        .reverse()
        .map((item, index) => {
            const result = item + zeros;
            zeros += "0";
            return result;
        });
        const flagsLowFiltered = flagsLow.filter((objectItem) =>
            permLowArr.includes(objectItem.code)
        );
        const lowPermArray = flagsLowFiltered.map((item) => item.name);
        let perm = lowPermArray;

        // high permissions
        // Permissions: make array of objects
        const flagsHigh = Flags.High.map((objectItem) => {
            const code = (Object.values(objectItem)[0] >>> 0).toString(2);
            return { name: Object.keys(objectItem)[0], code: code };
        });
        zeros = "";
        const permHighArr = (highPermDec >>> 0)
        .toString(2)
        .split("")
        .reverse()
        .map((item, index) => {
            const result = item + zeros;
            zeros += "0";
            return result;
        });
        const flagsHighFiltered = flagsHigh.filter((objectItem) =>
            permHighArr.includes(objectItem.code)
        );
        const highPermArray = flagsHighFiltered.map((item) => item.name);
        perm.concat(highPermArray);
        return perm;
    }
    
    protected async onInit(): Promise<void> {
        await this._getLists();
        this.properties.newListName = "";
    }

    public async render()
    {
        if (!this.currentUserRole) {
            this.currentUserRole = await this._getUserPermissions();
        }
        
        if ((this.currentUserRole == "admin") || (this.currentUserRole == "owner") ) {
            this.isOwner = true;
        }
        
        // get title of lookup list
        const lookupList = this.lookupListsOptions.filter(list => list.key == this.properties.lookupList)[0];
        let lookupListTitle;
        if (lookupList) {
            lookupListTitle = lookupList["text"];
        }
        else {
            lookupListTitle = "Undefined";
        }

        const element: React.ReactElement<IUrlShortenerProps> = React.createElement(
            UrlShortener,
            {
                context: this.context,
                idLength: this.properties.idLength,
                lookupList: {id: this.properties.lookupList, title: lookupListTitle},
            }
        );

        ReactDom.render(element, this.domElement);
    }

    private async _createList (name:String) 
    {
        const url = '/_api/web/lists';
        const data = JSON.stringify({
            '__metadata': { 'type': "SP.List" },
            "BaseTemplate": 100,
            "Description": "URL lookup list for URL-shortener-web-part",
            "Title": name,
        });
        const response = await this._spApiPost(url, data);
        return response;
    }

    private async _createListField (listName:String, FieldName:String, FieldType:number, EnforceUniqueValues:boolean) 
    {
        // create Field
        const urlCreateField = `/_api/web/lists/GetByTitle('${listName}')/Fields`;
        const dataCreateField = JSON.stringify({
            '__metadata': { 'type': "SP.Field" },
            "Title": FieldName,
            "StaticName": FieldName,
            "FieldTypeKind": FieldType,
            "Required": "true",
            "EnforceUniqueValues": EnforceUniqueValues,
            "Indexed": EnforceUniqueValues,
        });
        const responseCreateField = await this._spApiPost(urlCreateField, dataCreateField);
        if (responseCreateField["status"].toString().startsWith("2")) 
        {
            // add Field to view
            const urlAddFieldToView = `/_api/web/lists/GetByTitle('${listName}')/Views/GetByTitle('All%20Items')/ViewFields/addViewField('${FieldName}')`;
            const responseAddFieldToView = await this._spApiPost(urlAddFieldToView, "");
            return responseAddFieldToView;
        }
        else { return responseCreateField; }
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration 
    {
        let propertypaneDescription = this.isOwner ? null : 'You need to be owner to configure this web part';
        let selectListGroupFields = [];
        let createListGroupFields = [];
        let lengthOfIdGroupFields = [];

        const _checkList = async () => 
        {
            const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${this.properties.lookupList}')/fields?$select=Title&$filter=((Title eq 'Title') or (Title eq 'Key') or (Title eq 'TargetUrl')) and Hidden eq false and ReadOnlyField eq false `;
            const response = await this._spApiGet(url);
            if (response["status"] && response["status"].toString().startsWith("2"))
            {
                if (response["value"].length >= 3) 
                {
                    let allRequiredFieldsExist = true;
                    ["Title", "Key", "TargetUrl"].forEach(requiredField => {
                        let requiredFieldExists = false;
                        response["value"].forEach(listField => {
                            if (listField.Title == requiredField) { // if one listField.Title has the same name as the requiredField, it exists in the list.
                                requiredFieldExists = true;
                                return;
                            }
                        });
                        if (!requiredFieldExists) { // if one requiredField does not exist in list, the list doesn't approve the requirements
                            allRequiredFieldsExist = false;
                            return;
                        }
                    });
                    if (allRequiredFieldsExist) {
                        alert(`List '${this.properties.newListName}' is valid for use as lookup list. Columns 'Title', 'Key', TargetUrl' exist.`);
                    }
                    else {
                        alert(`List '${this.properties.newListName}' is not valid for use as lookup list. One of the columns 'Title', 'Key', TargetUrl' is missing or has the wrong type. Please create a new lookup List below.`);
                    }
                }
                else {
                    alert(`List '${this.properties.newListName}' is not valid for use as lookup list. One of the columns 'Title', 'Key', TargetUrl' is missing or has the wrong type. Please create a new lookup List below.`);
                }
            }
        };

        const _createListStack = async () => 
        {
            const createListResonse = await this._createList (this.properties.newListName);
            if (createListResonse["status"].toString().startsWith("2"))
            {
                const createListField1Resonse = await this._createListField(this.properties.newListName, "Key", 2, true);
                if (createListField1Resonse["status"].toString().startsWith("2"))
                {
                    const createListField2Resonse = await this._createListField(this.properties.newListName, "TargetUrl", 11, false);
                    if (createListField2Resonse["status"].toString().startsWith("2"))
                    {
                        alert(`List '${this.properties.newListName}' successfully created`);
                        await this._getLists();
                        this.context.propertyPane.refresh();
                    }
                    else {alert(createListField2Resonse["odata.error"].message.value);}
                }
                else {alert(createListField1Resonse["odata.error"].message.value);}
            }
            else {alert(createListResonse["odata.error"].message.value);}
        };

        // define description for web part in property pane with version, buildTimeStamp and link for doc website
        const about = [
            PropertyPaneLabel('version', {  
                text: "Version "  + packageSolution['solution'].version,
            }),
            PropertyPaneLabel('buildTimeStamp', {  
                text: buildTimeStamp,
            }),
            PropertyPaneLink('link', {  
                href: 'https://sharepoint-url-shortener-and-redirector.net',
                text: 'Documentation',
                target: '_blank',
            }),
        ];
        selectListGroupFields = [
            PropertyPaneDropdown('lookupList', {
                label: 'Select list',
                options: this.lookupListsOptions,
                disabled: !this.isOwner,
            }),
            PropertyPaneButton('lookupListButton', {
                text: "Check list",
                buttonType: PropertyPaneButtonType.Normal,
                onClick: _checkList,
                disabled: !this.isOwner,
            }),
        ];

        createListGroupFields = [
            PropertyPaneTextField('newListName', {
                label: 'Name',
                disabled: !this.isOwner,
            }),
            PropertyPaneButton('createNewListButton', {
                text: "Create",
                buttonType: PropertyPaneButtonType.Normal,
                onClick: _createListStack,
                disabled: !this.isOwner,
            }),
        ];

        lengthOfIdGroupFields = [
            PropertyPaneSlider('idLength', {
                label: 'Length of generated Id',
                min: 4,
                max: 16,
                disabled: !this.isOwner,
            }),
        ];
        
        return {
            pages: [
                {
                    header: {
                        description: propertypaneDescription
                    },
                    groups: [
                        {
                            groupName: '',
                            groupFields: about
                        },
                        {
                            groupName: "Create new URL lookup list",
                            groupFields: createListGroupFields
                        },
                        {
                            groupName: "Select URL lookup list",
                            groupFields: selectListGroupFields
                        },
                        {
                            groupName: "Set length of generated Id",
                            groupFields: lengthOfIdGroupFields
                        }
                    ]
                }
            ]   
        };
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }
}
