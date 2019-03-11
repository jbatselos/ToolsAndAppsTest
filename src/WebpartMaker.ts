export class WebpartMaker {

    public static makeWebPart(payload: any) {
        //Stringify the web part object
        var alertsString = JSON.stringify(payload);
        var urlEncodedAlerts = encodeURI(alertsString);

        //Prepare the session payload
        var sessionStorageLocal: any = sessionStorage;
        var state = sessionStorageLocal.sp_webpart_workbench_state
        var parsed = (new DOMParser()).parseFromString(state, "text/xml")
        var div = parsed.querySelectorAll('div[data-sp-controldata]') // Returns an array of all the web parts on the page
        //var data = div[0].attributes["data-sp-controldata"] 
        //var controlDataObject = JSON.parse(decodeURI(data.value))
        div[0].setAttribute("data-sp-controldata", urlEncodedAlerts);
        var serializer = new XMLSerializer();
        var withWebpart = serializer.serializeToString(parsed)

        //Inject the page contents into the session
        sessionStorageLocal.sp_webpart_workbench_state = withWebpart;
        return;
    }

    /*

    public static makeListToolsApps() {

    }

    public static makeTabbedToolsApps() {

    }

    public static makeTabbedCustomizableToolsApps() {

    }


    public static makeDefaultToolsApps() {
        var payload = JSON.parse(JSON.stringify(this.baseToolsApps))
        //Stringify the web part object
        var alertsString = JSON.stringify(payload);
        var urlEncodedAlerts = encodeURI(alertsString);

        //Prepare the session payload
        var sessionStorageLocal: any = sessionStorage;
        var state = sessionStorageLocal.sp_webpart_workbench_state
        var parsed = (new DOMParser()).parseFromString(state, "text/xml")
        var div = parsed.querySelectorAll('div[data-sp-controldata]') // Returns an array of all the web parts on the page
        //var data = div[0].attributes["data-sp-controldata"] 
        //var controlDataObject = JSON.parse(decodeURI(data.value))
        div[0].setAttribute("data-sp-controldata", urlEncodedAlerts);
        var serializer = new XMLSerializer();
        var withWebpart = serializer.serializeToString(parsed)

        //Inject the page contents into the session
        sessionStorageLocal.sp_webpart_workbench_state = withWebpart;
        return;
    }

    public static makeCustomizableToolsApps() {
        var payload = JSON.parse(JSON.stringify(this.baseToolsApps));
        payload.webPartData.properties.isUserTileCustomizationAllowed = true;

        //Stringify the web part object
        var alertsString = JSON.stringify(payload);
        var urlEncodedAlerts = encodeURI(alertsString);

        //Prepare the session payload
        var sessionStorageLocal: any = sessionStorage;
        var state = sessionStorageLocal.sp_webpart_workbench_state
        var parsed = (new DOMParser()).parseFromString(state, "text/xml")
        var div = parsed.querySelectorAll('div[data-sp-controldata]') // Returns an array of all the web parts on the page
        //var data = div[0].attributes["data-sp-controldata"] 
        //var controlDataObject = JSON.parse(decodeURI(data.value))
        div[0].setAttribute("data-sp-controldata", urlEncodedAlerts);
        var serializer = new XMLSerializer();
        var withWebpart = serializer.serializeToString(parsed)

        //Inject the page contents into the session
        sessionStorageLocal.sp_webpart_workbench_state = withWebpart;
        return;
    }


    public static makeDefaultAlertsWebpart() {
        //Declarative Alerts Webpart for workbench page
        //var payload = WebpartMaker.payload = {
        var payload = JSON.parse(JSON.stringify(this.baseAlerts));

        //Stringify the web part object
        var alertsString = JSON.stringify(payload);
        var urlEncodedAlerts = encodeURI(alertsString);

        //Prepare the session payload
        var sessionStorageLocal: any = sessionStorage;
        var state = sessionStorageLocal.sp_webpart_workbench_state
        var parsed = (new DOMParser()).parseFromString(state, "text/xml")
        var div = parsed.querySelectorAll('div[data-sp-controldata]') // Returns an array of all the web parts on the page
        //var data = div[0].attributes["data-sp-controldata"] 
        //var controlDataObject = JSON.parse(decodeURI(data.value))
        div[0].setAttribute("data-sp-controldata", urlEncodedAlerts);
        var serializer = new XMLSerializer();
        var withWebpart = serializer.serializeToString(parsed)

        //Inject the page contents into the session
        sessionStorageLocal.sp_webpart_workbench_state = withWebpart;
        return;
    }

    public static readonly baseAlerts = {
        "controlType": 3,
        "displayMode": 2,
        "id": "8737c381-6459-4f26-84de-68fbbe3fb569",
        "position": {
            "zoneIndex": 1,
            "sectionIndex": 1,
            "controlIndex": 1,
            "sectionFactor": 12
        },
        "webPartId": "d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf",
        "webPartData": {
            "id": "d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf",
            "instanceId": "8737c381-6459-4f26-84de-68fbbe3fb569",
            "title": "Alerts",
            "description": "alerts description",
            "serverProcessedContent": {
                "htmlStrings": {},
                "searchablePlainTexts": {},
                "imageSources": {},
                "links": {}
            },
            "dataVersion": "1.0",
            "properties": {
                "description": "Web part for displaying temporary alerts or notifications.  All settings are editable in-line",
                "title": "Alert that is really important 1234",
                "buttonText": "Learn More",
                "href": "https://duckduckgo.com"
            }
        }
    }

    public static readonly baseToolsApps = {
        "controlType": 3,
        "displayMode": 2,
        "id": "8737c381-6459-4f26-84de-68fbbe3fb569",
        "position": {
            "zoneIndex": 1,
            "sectionIndex": 1,
            "controlIndex": 1,
            "sectionFactor": 12
        },
        "webPartId": "5fb28714-f831-431c-b5cb-6f24a558f522",
        "webPartData": {
            "id": "5fb28714-f831-431c-b5cb-6f24a558f522",
            "instanceId": "8737c381-6459-4f26-84de-68fbbe3fb569",
            "title": "Tools & Apps",
            "description": "ToolsApps description",
            "serverProcessedContent": {
                "htmlStrings": {},
                "searchablePlainTexts": {},
                "imageSources": {},
                "links": {}
            },
            "dataVersion": "1.0",
            "properties": {
                "commonSettings": {},
                "description": "ToolsApps",
                "categoryOrder": [],
                "viewMode": "tile",
                "layout": "default",
                "isUserTileCustomizationAllowed": false,
                "targetProfileProperty": "ToolsAppsCustomization",
                "manualItems": [
                    {
                        "Title": "Wallstreet Journal",
                        "Description": "The Wall Street Journal is an American business-focused, English-language international daily newspaper based in New York City.",
                        "Category": "Sample Tab 1",
                        "LinkUrl": "https://wsj.com",
                        "IconUrl": "https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67",
                        "AltText": "The Wall Street Journal",
                        "Enable": true,
                        "UserConfigurable": true
                    },
                    {
                        "Title": "CNN",
                        "Description": "Cable News Network is an American basic cable and satellite television news channel owned by the Turner Broadcasting System, a division of Time Warner. CNN was founded in 1980 by American media proprietor Ted Turner as a 24-hour cable news channel.",
                        "Category": "Sample Tab 2",
                        "LinkUrl": "https://cnn.com",
                        "IconUrl": "http://media.idownloadblog.com/wp-content/uploads/2014/09/cnn-icon.png",
                        "AltText": "CNN",
                        "Enable": true,
                        "UserConfigurable": true
                    },
                    {
                        "Title": "Boston Globe",
                        "Description": "The Boston Globe is an American daily newspaper founded and based in Boston, Massachusetts,",
                        "Category": "Sample Tab 3",
                        "LinkUrl": "https://bostonglobe.com",
                        "IconUrl": "http://earthtones.org/wp-content/uploads/2012/12/boston-globe-icon.jpg",
                        "AltText": "The Boston Globe",
                        "Enable": true,
                        "UserConfigurable": true
                    }
                ],
                "contentSourceConfig": {
                    "selectedSites": [],
                    "dataSourceList": "",
                    "mappedFields": [
                        {
                            "key": "Title",
                            "label": "Title",
                            "types": [
                                "Text"
                            ],
                            "internalName": "Title"
                        },
                        {
                            "key": "Category",
                            "label": "Category",
                            "types": [
                                "Text"
                            ],
                            "internalName": "Category"
                        },
                        {
                            "key": "Description",
                            "label": "Description",
                            "types": [
                                "Text",
                                "Note"
                            ],
                            "internalName": "Description"
                        },
                        {
                            "key": "Enable",
                            "label": "Enable",
                            "types": [
                                "Boolean"
                            ]
                        },
                        {
                            "key": "IconUrl",
                            "label": "Icon Url",
                            "types": [
                                "Text",
                                "Note",
                                "URL"
                            ],
                            "isPictureField": 1
                        },
                        {
                            "key": "LinkUrl",
                            "label": "Target Url",
                            "types": [
                                "Text",
                                "Note",
                                "URL"
                            ],
                            "internalName": ""
                        },
                        {
                            "key": "AltText",
                            "label": "Alternative text",
                            "types": [
                                "Text",
                                "Note"
                            ],
                            "internalName": ""
                        },
                        {
                            "key": "OrderBy",
                            "label": "Sort by",
                            "types": null,
                            "internalName": "Title"
                        }
                    ],
                    "preselectCurrentSiteForTheFirstTime": 1,
                    "preselectListForTheFirstTime": 1,
                    "defaultListTitle": "Tools & Apps"
                },
                "targetingConfig": {
                    "enableTargeting": 0,
                    "termSetName": "Topics",
                    "selectedTerms": [],
                    "upsFieldKey": "MatchUPSProfileFields",
                    "mappedManagedPropertyName": "",
                    "mappedFieldInternalName": "",
                    "isUserSpecificContent": 0,
                    "showItemsWithEmptyTag": 0,
                    "isAndTermsJoin": 0,
                    "enableContentTypeFilter": 0,
                    "contentTypeFilter": ""
                }
            }

        }
    }
       */
}