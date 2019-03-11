
//Declarative Alerts Webpart for workbench page
var alerts = {
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

//Stringify the web part object
var alertsString = JSON.stringify(alerts);
var urlEncodedAlerts = encodeURI(alertsString);

//Prepare the session payload
var state = sessionStorage.sp_webpart_workbench_state
var parsed = (new DOMParser()).parseFromString(state, "text/xml")
var div = parsed.querySelectorAll('div[data-sp-controldata]') // Returns an array of all the web parts on the page
//var data = div[0].attributes["data-sp-controldata"] 
//var controlDataObject = JSON.parse(decodeURI(data.value))
div[0].setAttribute("data-sp-controldata", urlEncodedAlerts);
var serializer = new XMLSerializer();
var withWebpart = serializer.serializeToString(parsed)

//Inject the page contents into the session
sessionStorage.sp_webpart_workbench_state = withWebpart;

//Reload the page quickly so the session doesn't get overwritten by the chron job
location.reload();



// Works (taken straight from debug window)
//sessionStorage.sp_webpart_workbench_state = "<div><div data-sp-canvascontrol=\"\" data-sp-controldata=\"%7B%22controlType%22:3,%22displayMode%22:2,%22id%22:%22bfbed9c0-c0e5-4e7f-b654-a5d7b1138ab2%22,%22position%22:%7B%22zoneIndex%22:1,%22sectionIndex%22:1,%22controlIndex%22:1,%22sectionFactor%22:12%7D,%22webPartId%22:%22d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf%22,%22webPartData%22:%7B%22id%22:%22d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf%22,%22instanceId%22:%22bfbed9c0-c0e5-4e7f-b654-a5d7b1138ab2%22,%22title%22:%22Alerts%22,%22description%22:%22alerts%20description%22,%22serverProcessedContent%22:%7B%22htmlStrings%22:%7B%7D,%22searchablePlainTexts%22:%7B%7D,%22imageSources%22:%7B%7D,%22links%22:%7B%7D%7D,%22dataVersion%22:%221.0%22,%22properties%22:%7B%22description%22:%22Web%20part%20for%20displaying%20temporary%20alerts%20or%20notifications.%20%20All%20settings%20are%20editable%20in-line%22,%22title%22:%22asdf%22,%22buttonText%22:%22asdfasdf%22,%22href%22:%22https://duckduckgo.com%22%7D%7D%7D\"><div data-sp-componentid=\"\">d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf</div><div data-sp-htmlproperties=\"\"></div></div></div>"

// Doesn't work (from serializer)
//sessionStorage.sp_webpart_workbench_state = "<div><div data-sp-canvascontrol=\"\" data-sp-controldata=\"{&quot;controlType&quot;:3,&quot;displayMode&quot;:2,&quot;id&quot;:&quot;8737c381-6459-4f26-84de-68fbbe3fb569&quot;,&quot;position&quot;:{&quot;zoneIndex&quot;:1,&quot;sectionIndex&quot;:1,&quot;controlIndex&quot;:1,&quot;sectionFactor&quot;:12},&quot;webPartId&quot;:&quot;d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf&quot;,&quot;webPartData&quot;:{&quot;id&quot;:&quot;d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf&quot;,&quot;instanceId&quot;:&quot;8737c381-6459-4f26-84de-68fbbe3fb569&quot;,&quot;title&quot;:&quot;Alerts&quot;,&quot;description&quot;:&quot;alerts description&quot;,&quot;serverProcessedContent&quot;:{&quot;htmlStrings&quot;:{},&quot;searchablePlainTexts&quot;:{},&quot;imageSources&quot;:{},&quot;links&quot;:{}},&quot;dataVersion&quot;:&quot;1.0&quot;,&quot;properties&quot;:{&quot;description&quot;:&quot;Web part for displaying temporary alerts or notifications.  All settings are editable in-line&quot;,&quot;title&quot;:&quot;this is a test&quot;,&quot;buttonText&quot;:&quot;go&quot;,&quot;href&quot;:&quot;https://duckduckgo.com&quot;}}}\"><div data-sp-componentid=\"\">d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf</div><div data-sp-htmlproperties=\"\"/></div></div>"