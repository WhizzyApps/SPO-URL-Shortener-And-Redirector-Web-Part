import * as React from 'react';
import styles from './UrlShortener.module.scss';
import { IUrlShortenerProps } from './IUrlShortenerProps';
import { TextField, DefaultButton, Icon, find} from 'office-ui-fabric-react/lib';

import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

export default class SpoUrlShortener extends React.Component<IUrlShortenerProps, {}> 
{
    public state = {
        urlIdStatus: 0, // 1.. without id, 2.. invalid id, 3.. valid id
        inputUrl: "",
        inputId: "",
        outputUrl: "",
        redirectingUrl: "",
        copyFeedback: "",
        postFeedback1: false,
        postFeedback2: false,
        lookUpFeedback: false,
        lookUpError: "",
        outputId: "",
        findId: "",
        shortenError: "",
        inputUrlError: "",
        hashLength: this.props.idLength.toString(),
    };
    private properties = {
        currentUrl: window.location.href,
    };
    
    private async _spApiGet (url: string): Promise<object> 
    {
        const clientOptions: ISPHttpClientOptions = {
            headers: new Headers(),
            method: 'GET',
            mode: 'cors'
        };
        const response = await this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + url, SPHttpClient.configurations.v1, clientOptions);
        const responseJson = await response.json();
        responseJson['status'] = response.status;
        if (!responseJson.value) {
            responseJson['value'] = [];
        }
        return responseJson;
    } 
    
    private async _spApiPost (url: string, data: string): Promise<object> 
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
        const response = await this.props.context.spHttpClient.post(this.props.context.pageContext.web.absoluteUrl + url, SPHttpClient.configurations.v1, clientOptions);
        const responseJson = await response.json();
        responseJson['status'] = response.status;
        if (!responseJson.value) {
            responseJson['value'] = [];
        }
        return responseJson;
    }
    
    private _generateRandomHash(hashLength: number) {
        const characters = "0123456789ABCDEFGHIJKLMNPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
        let string = "";
        for (let i = 0; i < hashLength; i++) {
            string += characters[Math.floor(Math.random() * characters.length)];
        }
        return string;
    }

    private async _createListItem (postUrl:string, targetUrl:string, outputUrl:string, outputId:string, isGeneratedId:boolean) 
    {
        let getListResponse = await this._spApiGet(`/_api/web/lists/GetById('${this.props.lookupList.id}')?$select=ListItemEntityTypeFullName`); 
        if (getListResponse["status"] == 200) 
        {
            let type = getListResponse["ListItemEntityTypeFullName"];
            if (targetUrl.length <= 255) 
            {
                let data = JSON.stringify({
                    '__metadata': { 'type': type },
                    "Title": outputId,
                    "Key": outputId,
                    "TargetUrl": 
                    {
                        '__metadata': { 'type': 'SP.FieldUrlValue' },
                        'Url': targetUrl,
                    },
                });
                let setItemsResponse1 = await this._spApiPost(postUrl, data);
                // handle result
                if (setItemsResponse1["status"] == 201) {
                    if (isGeneratedId) {this.setState({postFeedback1: true, outputUrl: outputUrl, outputId: outputId });}
                    else { this.setState({postFeedback2: true, outputUrl: outputUrl}); }
                }
                else if (setItemsResponse1["odata.error"].message.value.includes("Invalid URL")) 
                {
                    data = JSON.stringify({
                        '__metadata': { 'type': type },
                        "Title": outputId,
                        "Key": outputId,
                        "TargetUrl2": targetUrl
                    });
                    let setItemsResponse2 = await this._spApiPost(postUrl, data);
                    if (setItemsResponse2["status"] == 201) {
                        if (isGeneratedId) {this.setState({postFeedback1: true, outputUrl: outputUrl, outputId: outputId });}
                        else { this.setState({postFeedback2: true, outputUrl: outputUrl}); }
                    }
                    else {this.setState({shortenError: `SharePoint API error: ${setItemsResponse2["odata.error"].message.value}`}); }
                }
                else { this.setState({shortenError: `SharePoint API error: ${setItemsResponse1["odata.error"].message.value}`}); }
            }
            else // if (targetUrl.length > 255)
            {
                let data = JSON.stringify({
                    '__metadata': { 'type': type },
                    "Title": outputId,
                    "Key": outputId,
                    "TargetUrl2": targetUrl
                });
                let setItemsResponse = await this._spApiPost(postUrl, data);
                if (setItemsResponse["status"] == 201) {
                    if (isGeneratedId) {this.setState({postFeedback1: true, outputUrl: outputUrl, outputId: outputId });}
                    else { this.setState({postFeedback2: true, outputUrl: outputUrl}); }
                }
                else {this.setState({shortenError: `SharePoint API error: ${setItemsResponse["odata.error"].message.value}`}); }
            }
        }
        else {this.setState({shortenError: getListResponse["error"].message});}
    }

    private async _shorten_url_with_hash () 
    {
        // reset state
        this.setState({
            inputId: "",
            outputUrl: "",
            copyFeedback: "",
            postFeedback1: false,
            postFeedback2: false,
            lookUpFeedback: false,
            lookUpError: "",
            outputId: "",
            findId: "",
            shortenError: "",
            inputUrlError: "",
        });

        if (this.state.inputUrl) 
        {
            let hashIsUnique = false;
            let outputId = "";
            let loopCount = 0;
            let repeatLimitExceeded = false;
            let uniqueIdError = false;
            while (!hashIsUnique && !repeatLimitExceeded && !uniqueIdError) 
            {
                // generate Id
                let hash = this._generateRandomHash(this.props.idLength);
                // check if Id exists
                let getItemResponse = await this._spApiGet(`/_api/web/lists/GetById('${this.props.lookupList.id}')/items?$select=Key,TargetUrl&$filter=Key%20eq%20'${hash}'`); 
                let item = getItemResponse["value"];
                if (getItemResponse["status"] == 200) {
                    if (item.length == 0) {
                        hashIsUnique = true;
                        outputId = hash;
                    } 
                    else if (loopCount > 9) {repeatLimitExceeded = true;}
                    else {loopCount++;}
                }
                else {
                    uniqueIdError = true;
                    this.setState( {shortenError: getItemResponse["error"].message} );
                }
            }
            let outputUrl = this.properties.currentUrl + "?$id=" + outputId;

            if (hashIsUnique) {
                // post outputId and target URL to list
                if (!repeatLimitExceeded) 
                {
                    let postUrl = `/_api/web/lists/GetById('${this.props.lookupList.id}')/items`;
                    let targetUrl = this.state.inputUrl.trim();
                    await this._createListItem(postUrl, targetUrl, outputUrl, outputId, true);
                }
                else {this.setState({shortenError: `Error: Tried to generate unique Id 10 times, but all Ids already exist in list. Increase the length of generated Id in the configuration of the web part and try again.`});}
            }
        }
        else {this.setState({inputUrlError: `Please enter a valid URL`});}
    }

    private async _shorten_url_with_Id () 
    {
        // reset state
        this.setState({
            outputUrl: "", 
            outputId: "",
            findId: "",
            copyFeedback: "",
            postFeedback1: false,
            postFeedback2: false,
            lookUpFeedback: false, 
            lookUpError: "",
            shortenError: "",
            inputUrlError: "",
        });

        if (this.state.inputUrl) 
        {
            if (this.state.inputId) 
            {
                // check if Id exists
                let getItemResponse = await this._spApiGet(`/_api/web/lists/GetById('${this.props.lookupList.id}')/items?$select=Key,TargetUrl&$filter=Key%20eq%20'${this.state.inputId}'`); 
                let item = getItemResponse["value"];
                if (getItemResponse["status"] == 200) {
                    // id does not exist
                    if (item.length == 0)
                    {
                        // post entries to list
                        let postUrl = `/_api/web/lists/GetById('${this.props.lookupList.id}')/items`;
                        let outputUrl = this.properties.currentUrl + "?$id=" + this.state.inputId;
                        let outputId = this.state.inputId;
                        let targetUrl = this.state.inputUrl.trim();
                        await this._createListItem(postUrl, targetUrl, outputUrl, outputId, false);
                    } 
                    // id exists
                    else { this.setState({shortenError: `Error: The list item could not be added, because the Id already exists. Please enter a different Id.`}); }
                }
                else {this.setState( {shortenError: getItemResponse["error"].message} );}
            }
            else {this.setState({shortenError: `Please enter a valid Id`});}
        }
        else {this.setState({inputUrlError: `Please enter a valid URL`});}
    }

    private _openList () {
        window.open(`${this.props.context.pageContext.web.absoluteUrl}/Lists/${this.props.lookupList.title}`);
    }

    private _copy_url () {
        if (this.state.outputUrl) 
        {
            navigator["clipboard"].writeText(this.state.outputUrl);
            this.setState({copyFeedback:"URL copied"});
        }
    }

    private async _find_url () 
    {
        // reset state
        this.setState({
            inputUrl: "",
            outputUrl: "",
            outputId: "",
            inputId: "",
            copyFeedback: "",
            postFeedback1: false,
            postFeedback2: false,
            shortenError: "",
            inputUrlError: "",
            lookUpFeedback: false,
            lookUpError: "",
        });
        // get Id - get findId from user input
        let findId = this.state.findId;
        if (findId) {
            // search Id in list
            let getItemResponse = await this._spApiGet(`/_api/web/lists/GetById('${this.props.lookupList.id}')/items?$select=Key,TargetUrl,TargetUrl2&$filter=Key%20eq%20'`+findId+"'"); 
            let item = getItemResponse["value"];
            if (getItemResponse["status"] == 200) 
            {
                // if findId exists
                if (item.length == 1) {
                    let outputUrl = this.properties.currentUrl + "?$id=" + findId;
                    let targetUrl = item[0].TargetUrl && item[0].TargetUrl.Url;
                    if (!targetUrl) { targetUrl = item[0].TargetUrl2; }
                    this.setState({inputUrl: targetUrl, outputUrl: outputUrl, lookUpFeedback: true });
                }
                // if findId not exists
                else {this.setState({lookUpError: `Id not found in URL lookup list '${this.props.lookupList.title}'`,});}
            }
            else {this.setState({lookUpError: `SharePoint API error: ${getItemResponse["error"].message}`});}
        }
        else {this.setState({lookUpError: `Please enter a valid Id`});}
    }

    public async componentDidMount () 
    {
        if (this.properties.currentUrl.includes("?$id="))
        {
            // extract idFromUrl
            let idFromUrl = this.properties.currentUrl.split("?$id=")[1]; //63c0aad3-ed74-4773-bcce-6ef316433a58
            // check if Id valid
            let getItemResponse = await this._spApiGet(`/_api/web/lists/GetById('${this.props.lookupList.id}')/items?$select=Key,TargetUrl,TargetUrl2&$filter=Key%20eq%20'`+idFromUrl+"'"); 
            let item = getItemResponse["value"];
            if (item.length == 1) {
                let targetUrl = item[0].TargetUrl && item[0].TargetUrl.Url;
                if (!targetUrl) { targetUrl = item[0].TargetUrl2; }
                this.setState({urlIdStatus: 3, redirectingUrl: targetUrl, outputId: idFromUrl});
                window.location.href = targetUrl;
            }
            // if invalid
            else {this.setState({urlIdStatus: 2, outputId: idFromUrl});}
        }
        // no Id
        else {this.setState({urlIdStatus: 1});}
    }

    public render(): React.ReactElement<IUrlShortenerProps> 
    {
        // display message if no lookup list is selected
        let disabled = false;
        let disabledMessage = "";
        if (this.props.lookupList.title == "Undefined")
        {
            disabled = true;
            if (this.props.lookupList.id) {
                disabledMessage = "Previous selected lookup list doesn't exist. It might have been deleted. Please edit the web part and select a valid list.";
            }
            else {
                disabledMessage = "No lookup list selected. Please edit the web part and select a valid list.";
            }
        }

        return (
            <div className={ styles.urlShortener }>
                <div className={ styles.container }>
                    <div className={ styles.row }>
                        <div className={ styles.column }>
                            <div className={ styles.title }>SharePoint URL Shortener and Redirector </div>
                            {   // valid id
                                this.state.urlIdStatus == 3 && <div className={styles.inputUrl} >Redirecting to {this.state.redirectingUrl} </div>
                            }
                            {   // invalid id
                                this.state.urlIdStatus == 2 && <>
                                    <div className={styles.inputUrl}>
                                        <div style={{padding: "0.3rem 0"}}>Shortened URL with Id: '{this.state.outputId}' not found! </div>
                                        <div onClick={()=>this._openList()}>
                                            <DefaultButton disabled={disabled} text="Open list" className={styles.button} href={'#'}/>
                                        </div>
                                    </div>
                                </>
                            }
                            {   // no id
                                this.state.urlIdStatus == 1 && 
                                <>
                                    {// optional  section
                                        disabled && 
                                        <div className={styles.error}>
                                            {disabledMessage}
                                        </div>
                                    }
                                    {/* 1) Input URL section */}
                                    <TextField className={styles.inputUrl} disabled={disabled} borderless label="Input URL" value={this.state.inputUrl} onChange={(event, value) => {this.setState({inputUrl: value});}}></TextField>
                                    {// optional inputUrlError section
                                        this.state.inputUrlError && 
                                        <div className={styles.error}>
                                            {this.state.inputUrlError}
                                        </div>
                                    }
                                    {/* 2) Input ID section */}
                                    <div className={styles.inputId}>
                                        {/* I don't have an Id */}
                                        <div className={styles.col} style={{marginRight: "0.1rem"}}>
                                            <div style={{padding: "5px 0px"}}>I don't have an Id</div>
                                            <TextField borderless disabled={disabled} readOnly placeholder="Generated Id" value={this.state.outputId} ></TextField>
                                            <div style={{display:"flex", alignItems: "center"}}>
                                                <div onClick={()=>{this._shorten_url_with_hash();}}>
                                                    <DefaultButton disabled={disabled} text="Shorten and add to list" className={styles.button} href={'#'}/>
                                                </div>
                                                <div className={styles.iconContainer}>
                                                    {this.state.postFeedback1 && <Icon iconName={"CheckMark"} className={`${styles.successIcon}`} title='URL successfully added to list' /> }
                                                </div>
                                            </div>
                                        </div>
                                        {/* I have an Id */}
                                        <div className={styles.col} style={{marginLeft: "0.1rem"}}>
                                            <div style={{padding: "5px 0px"}}>I have an Id</div>
                                            <TextField borderless disabled={disabled} placeholder="Example Id: abc123" value={this.state.inputId} onChange={(event, value) => {this.setState({inputId: value});}}></TextField>
                                            <div style={{display:"flex", alignItems: "center"}}>
                                                <div onClick={()=>{this._shorten_url_with_Id();}}>
                                                    <DefaultButton disabled={disabled} text="Shorten and add to list" className={styles.button} href={'#'}/>
                                                </div>
                                                <div className={styles.iconContainer}>
                                                    {this.state.postFeedback2 && (
                                                        <Icon iconName={"CheckMark"} className={`${styles.successIcon}`} title='URL successfully added to list' /> 
                                                    )}
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    {// optional shortenError section
                                        this.state.shortenError && 
                                        <div className={styles.error}>
                                            {this.state.shortenError}
                                        </div>
                                    }
                                    {/* 3) Output URL section */}
                                    <div className={styles.output}>
                                        <TextField borderless disabled={disabled} label="Output URL" readOnly value={this.state.outputUrl} ></TextField>
                                        <div style={{display:"flex", justifyContent: "space-between", alignItems: "center"}}>
                                            <div onClick={()=>this._openList()}>
                                                <DefaultButton disabled={disabled} text="Open list" className={styles.button} href={'#'}/>
                                            </div>
                                            <div>
                                                {this.state.copyFeedback}
                                            </div>
                                            <div onClick={()=>this._copy_url()}>
                                                <DefaultButton disabled={disabled} text="Copy URL" className={styles.button} href={'#'}/>
                                            </div>
                                        </div>
                                    </div>
                                    
                                    {/* 4) Look up URL section */}
                                    <div className={styles.lookUp}>
                                        <div style={{paddingTop: "5px", textAlign: "center"}}>Look up URL for existing Id in list</div>
                                        <div style={{paddingBottom: "5px"}}>Input Id</div>
                                        <TextField borderless disabled={disabled} value={this.state.findId} onChange={(event, value) => {this.setState({findId: value});}}></TextField>
                                        <div style={{display:"flex", alignItems: "center"}}>
                                            <div onClick={()=>this._find_url()}>
                                                <DefaultButton disabled={disabled} text="Find URL" className={styles.button} href={'#'}/>
                                            </div>
                                            {
                                                this.state.lookUpFeedback &&
                                                <div className={styles.iconContainer}> 
                                                    <Icon iconName={"CheckMark"} className={`${styles.successIcon}`} title='URL found in list' />
                                                </div>
                                            }
                                        </div>
                                    </div>
                                    {// optional lookUpError section
                                        this.state.lookUpError && 
                                        <div className={styles.error}>
                                            {this.state.lookUpError}
                                        </div>
                                    }
                                </>
                            }
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
