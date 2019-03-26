import {
    WebDriver,
    Session,
    Capabilities,
    Executor,
    WebElement,
    By,
    Builder,
    Key,
    IWebDriverOptionsCookie
} from 'selenium-webdriver';
import * as fs from 'fs';
import { TestingUtil } from "./TestingUtil";
import { WebpartMaker } from "./WebpartMaker";
import { WatchIgnorePlugin } from 'webpack';
import { delay } from 'bluebird';
import symmetricDifferenceWith from 'ramda/es/symmetricDifferenceWith';




jest.setTimeout(20000);

var capabilities = new Capabilities(Capabilities.chrome());
var driver = new Builder().withCapabilities(capabilities).build();
var testingUtil = new TestingUtil(driver);
let standardWait = 3000;
let defaultAlerts = { "controlType": 3, "displayMode": 2, "id": "8737c381-6459-4f26-84de-68fbbe3fb569", "position": { "zoneIndex": 1, "sectionIndex": 1, "controlIndex": 1, "sectionFactor": 12 }, "webPartId": "d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf", "webPartData": { "id": "d3d7317d-b35c-40bf-b9c2-ac0ad847dbdf", "instanceId": "8737c381-6459-4f26-84de-68fbbe3fb569", "title": "Alerts", "description": "alerts description", "serverProcessedContent": { "htmlStrings": {}, "searchablePlainTexts": {}, "imageSources": {}, "links": {} }, "dataVersion": "1.0", "properties": { "description": "Web part for displaying temporary alerts or notifications. All settings are editable in-line", "title": "Alert that is really important 1234", "buttonText": "Learn More", "href": "https://duckduckgo.com" } } };
let defaultToolsApps = { "controlType": 3, "displayMode": 2, "id": "8737c381-6459-4f26-84de-68fbbe3fb569", "position": { "zoneIndex": 1, "sectionIndex": 1, "controlIndex": 1, "sectionFactor": 12 }, "webPartId": "5fb28714-f831-431c-b5cb-6f24a558f522", "webPartData": { "id": "5fb28714-f831-431c-b5cb-6f24a558f522", "instanceId": "8737c381-6459-4f26-84de-68fbbe3fb569", "title": "Tools & Apps", "description": "ToolsApps description", "serverProcessedContent": { "htmlStrings": {}, "searchablePlainTexts": {}, "imageSources": {}, "links": {} }, "dataVersion": "1.0", "properties": { "commonSettings": {}, "description": "ToolsApps", "categoryOrder": [], "viewMode": "tile", "layout": "default", "isUserTileCustomizationAllowed": false, "targetProfileProperty": "ToolsAppsCustomization", "manualItems": [{ "Title": "Wallstreet Journal", "Description": "The Wall Street Journal is an American business-focused, English-language international daily newspaper based in New York City.", "Category": "Sample Tab 1", "LinkUrl": "https://wsj.com", "IconUrl": "https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67", "AltText": "The Wall Street Journal", "Enable": true, "UserConfigurable": true }, { "Title": "CNN", "Description": "Cable News Network is an American basic cable and satellite television news channel owned by the Turner Broadcasting System, a division of Time Warner. CNN was founded in 1980 by American media proprietor Ted Turner as a 24-hour cable news channel.", "Category": "Sample Tab 2", "LinkUrl": "https://cnn.com", "IconUrl": "http://media.idownloadblog.com/wp-content/uploads/2014/09/cnn-icon.png", "AltText": "CNN", "Enable": true, "UserConfigurable": true }, { "Title": "Boston Globe", "Description": "The Boston Globe is an American daily newspaper founded and based in Boston, Massachusetts,", "Category": "Sample Tab 3", "LinkUrl": "https://bostonglobe.com", "IconUrl": "http://earthtones.org/wp-content/uploads/2012/12/boston-globe-icon.jpg", "AltText": "The Boston Globe", "Enable": true, "UserConfigurable": true }], "contentSourceConfig": { "selectedSites": [], "dataSourceList": "", "mappedFields": [{ "key": "Title", "label": "Title", "types": ["Text"], "internalName": "Title" }, { "key": "Category", "label": "Category", "types": ["Text"], "internalName": "Category" }, { "key": "Description", "label": "Description", "types": ["Text", "Note"], "internalName": "Description" }, { "key": "Enable", "label": "Enable", "types": ["Boolean"] }, { "key": "IconUrl", "label": "Icon Url", "types": ["Text", "Note", "URL"], "isPictureField": 1 }, { "key": "LinkUrl", "label": "Target Url", "types": ["Text", "Note", "URL"], "internalName": "" }, { "key": "AltText", "label": "Alternative text", "types": ["Text", "Note"], "internalName": "" }, { "key": "OrderBy", "label": "Sort by", "types": null, "internalName": "Title" }], "preselectCurrentSiteForTheFirstTime": 1, "preselectListForTheFirstTime": 1, "defaultListTitle": "Tools & Apps" }, "targetingConfig": { "enableTargeting": 0, "termSetName": "Topics", "selectedTerms": [], "upsFieldKey": "MatchUPSProfileFields", "mappedManagedPropertyName": "", "mappedFieldInternalName": "", "isUserSpecificContent": 0, "showItemsWithEmptyTag": 0, "isAndTermsJoin": 0, "enableContentTypeFilter": 0, "contentTypeFilter": "" } } } };

let driverJest;


it('waits for the driver to start', () => {
    return driver.then(_d => {
        driverJest = _d
    })
})

it('Can login to enviorment', async () => {await SetupEnviorment() })
it('Can set up a webpart', async () => { await TestWebpartSetup() })
it('Can turn on user customization', async () => { await TestUserCustomization() })
it('Targerting turns on when using Sharepoint data', async () => { await TestTargeting() })
//it('Can load list from Sharepoint Site', async () => { await TestSharepointList() })
//it('Can set item Limit', async () => { await TestItemLimit() })
//it('Manage items can change order', async () => { await TestManageItems() })
it('Layouts load correctly', async () => { await TestAllLayouts() })
/*
it('can hide links', async () => { await TestTAUserCustomization() })
it('renders list view mode', async () => { await TestListMode() })
it('renders tabbed layout', async () => { await TestTabbedLayout() })
it('doesn\'t break when removing all tiles from the tabbed layout', async () => { await TestRemovingAllLinks() })
it('warns about illegal profile properties', async () => { await TestIllegalProfileProperty() })
*/
it('closes the driver', async () => { await driver.close() })


async function SetupEnviorment() {
    jest.setTimeout(60000);
    await driver.manage().window().maximize();
    await (driver.get("https://owdevelop.sharepoint.com/sites/OWv2-Develop/SitePages/AutomatedTestPage.aspx"))
    await (timeout(300));
    await driver.findElement(By.name("loginfmt")).sendKeys("owdeveloper@owdevelop.onmicrosoft.com");
    await driver.findElement(By.id("idSIButton9")).click();
    await (timeout(300));
    await driver.findElement(By.name("passwd")).sendKeys("&LEM3TF'wkNe4cCh");
    
    await driver.findElement(By.id("idSIButton9")).click();
    await (timeout(300));
    await driver.findElement(By.id("idSIButton9")).click();
    await (timeout(300));
    await driver.findElement(By.name("Edit")).click();
    await (timeout(4000));
    expect(await driver.getCurrentUrl()).toBe("https://owdevelop.sharepoint.com/sites/OWv2-Develop/SitePages/AutomatedTestPage.aspx?Mode=Edit");
    
}
async function TestWebpartSetup() {

    await SetupWebpartLocal();

    var element = await (driver.findElement(By.css('div[data-sp-feature-tag="ToolsAppsWebPart web part (Tools & Apps)"]')));
    var attrib = await (element.getAttribute("class"));

    expect(attrib).not.toBeNull;
    await remove();
}
async function TestUserCustomization(){
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    //set up webpart
    await SetupWebpartLocal();
    //go to edit webpart
    
    let editButtons = await driver.findElements(By.css('button[aria-label="Edit web part"]'));
    await editButtons[1].click();
    
    //go to Layout
    await timeout(1000);

    await driver.findElement(By.xpath('//*[@id="spPropertyPaneContainer"]/div/div/div[2]/div/div[2]/div/div[1]/div/div[3]/div/div/button')).click();
    
    //turn on user cutomization

    let switches = await driver.findElements(By.css('button[aria-checked="false"]'));   
    await switches[0].click();

   
    //apply
    await driver.findElement(By.css('button[data-automation-id="propertyPaneApplyButton"]')).click();
    await timeout(800);
    //click on new user customization button
    await driver.findElement(By.xpath('//*[@id="spPageCanvasContent"]/div/div/div/div[2]/div/div/div/div/div[1]/div/div[1]/div[1]/div/div[1]/div/div/div[2]/section/div/div/div/button')).click();
    //cancel
    let cancel = await driver.findElements(By.css('i[data-icon-name="Cancel"]'));

    await cancel[2].click();
    //cancel again
    await timeout(500);
    await cancel[0].click();
    
    await remove();
}

async function TestTargeting() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    //set up webpart
    await SetupWebpartLocal();
    //go to edit webpart
        
    let editButtons = await driver.findElements(By.css('button[aria-label="Edit web part"]'));
    await editButtons[1].click();
    await timeout(1000);

    //go to Data

    await driver.findElement(By.xpath('//*[@id="spPropertyPaneContainer"]/div/div/div[2]/div/div[2]/div/div[1]/div/div[2]/div/div/button/div')).click();

    //Switch to Manual Source
    let switches = await driver.findElements(By.css('button[aria-checked="false"]'));   
    await switches[0].click();

    // see if Targeting tab appears
    await timeout(500);
    
    let targetTab = await driver.findElement(By.xpath('//*[@id="spPropertyPaneContainer"]/div/div/div[2]/div/div[2]/div/div[1]/div/div[4]/div/div/button/div'));
    var attrib = await (targetTab.getAttribute("class"));
    expect(attrib).not.toBeNull;
    let cancel = await driver.findElements(By.css('i[data-icon-name="Cancel"]'));
    await cancel[0].click();
    await remove();
}

async function TestAllLayouts() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    //set up webpart
    await SetupWebpartLocal();
    //go to edit webpart
    let editButtons = await driver.findElements(By.css('button[aria-label="Edit web part"]'));
    await editButtons[1].click();
    
    //go to Layout
    await timeout(1000);
    await driver.findElement(By.xpath('//*[@id="spPropertyPaneContainer"]/div/div/div[2]/div/div[2]/div/div[1]/div/div[3]/div/div/button')).click();
    // click lists
    
    await layoutTests('Sierra',"tile_3f281b04");
    //await layoutTests("KPI","kpiTitle_3d0ba172");
    await layoutTests('Default',"tile_a7acc3cc");
    // await  layoutTests('Site Owner',"siteOwnerColumn_3a7ff792");
    await  layoutTests('Tabbed',"tile_75ddc177" );
    await  layoutTests('Santiago',"tile_518b2c33");
    await  layoutTests('Savona',"tile_f9cf8a80");
    await  layoutTests('Idaho',"tile_2d33fb31");
    await  layoutTests('Accordion',"tile_060a83f7");
    
    await remove();
    
}

async function TestJulietRendering() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    webpart.webPartData.properties.style = "Juliet";
    await SetupWebpartLocal();

    await testingUtil.togglePublishMode();

    let alert = await driver.findElements(By.css("div[class^='alertViewMode'] > div"))
    expect(await alert[0].getText()).toBe("Alert that is really important 1234");
    let button = alert[1].findElement(By.css("a"))
    expect(await button.getText()).toBe("Learn More");
    expect(await button.getAttribute("href")).toBe("https://duckduckgo.com/");
    takeScreenShot("AlertsJuliet");
    await remove();
}

async function TestChicagoRendering() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    webpart.webPartData.properties.style = "Chicago";
    webpart.webPartData.properties.imageUrl = "http://icons.iconarchive.com/icons/graphicloads/100-flat/256/home-icon.png";
    await SetupWebpartLocal();

    await testingUtil.togglePublishMode();
    await timeout(standardWait);

    let alert = await driver.findElements(By.css("div[class^='alertViewMode'] > div"))
    let alertContent = alert[0]
    expect(await alertContent.findElement(By.css("div[class^='alertTextCell']")).getText()).toBe("Alert that is really important 1234");
    let button = alertContent.findElement(By.css("a[class^='alertButtonCell']"))
    expect(await button.getText()).toBe("Learn More");
    expect(await button.getAttribute("href")).toBe("https://duckduckgo.com/");

    let image = await alert[1].findElement(By.css("img"));
    expect(await image.getAttribute("src")).toBe("http://icons.iconarchive.com/icons/graphicloads/100-flat/256/home-icon.png");
    takeScreenShot("AlertsChicago");
    await remove();
}

async function TestToolsAppsWebpartSetup() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    await SetupWebpartLocal();

    var element = await (driver.findElement(By.xpath("//div[starts-with(@class, 'wsbThings')]")));
    var attrib = await (element.getAttribute("class"));

    expect(attrib).not.toBeNull;
    await remove();
}

async function TestTADefaultMode() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    await SetupWebpartLocal();

    let buttons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let button = buttons[0];
    expect(await button.getAttribute("href")).toBe("https://wsj.com/");
    expect(await button.getAttribute("class")).toMatch(/tile_.*/);
    expect(await button.findElement(By.css("h1")).getText()).toBe("Wallstreet Journal");
    expect(await button.findElement(By.css("img")).getAttribute("src")).toBe("https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67");
    takeScreenShot("ToolsAppsDefault");
    await remove();
}

async function TestTAButtons() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    await SetupWebpartLocal();

    let buttons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let button = buttons[0];
    await button.click();

    await pageLoaded();
    expect(await driver.getCurrentUrl()).toBe("https://www.wsj.com/");
    await remove();
}

async function TestTAUserCustomization() {
    await SetupWebpartLocal();


    let buttonsContainer = await (driver.findElement(By.css("div[class^='toolsAppsFlexContainer']")));
    let customizationButton = await buttonsContainer.findElement(By.css("button[class^='ms-Toggle-background']"));
    await customizationButton.click();

    let hideButtons = await buttonsContainer.findElements(By.css("div[class^='deleteButton'] i"));
    expect(await hideButtons[0].getAttribute("data-icon-name")).toBe("RemoveFilter");
    await hideButtons[0].click();
    expect(await hideButtons[0].getAttribute("data-icon-name")).toBe("AddTo");

    await customizationButton.click();
    await (timeout(standardWait));
    let updatedButtons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let secondButton = updatedButtons[0]
    expect(await secondButton.getAttribute("href")).toBe("https://cnn.com/");


    takeScreenShot("ToolsAppsUserCustomization");
    await remove();
}

async function TestListMode() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.viewMode = "list";
    await SetupWebpartLocal();

    let buttons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let button = buttons[0];
    expect(await button.getAttribute("href")).toBe("https://wsj.com/");
    expect(await button.getAttribute("class")).toMatch(/list_.*/);
    expect(await button.findElement(By.css("h1")).getText()).toBe("Wallstreet Journal");
    expect(await button.findElement(By.css("img")).getAttribute("src")).toBe("https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67");
    takeScreenShot("ToolsAppsListMode");
    await remove();
}

async function TestTabbedLayout() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.layout = "tabbed";
    await SetupWebpartLocal();

    let bigContainer = await (driver.findElement(By.css("div[class^='wsbThings']")));

    let button = await (bigContainer.findElement(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    expect(await button.getAttribute("href")).toBe("https://wsj.com/");
    expect(await button.getAttribute("class")).toMatch(/tile_.*/);
    expect(await button.findElement(By.css("h1")).getText()).toBe("Wallstreet Journal");
    expect(await button.findElement(By.css("img")).getAttribute("src")).toBe("https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67");

    let tabButtons = await bigContainer.findElements(By.css("div[role='presentation'] > ul > button"));
    await tabButtons[1].click();

    button = await (bigContainer.findElement(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    expect(await button.getAttribute("href")).toBe("https://cnn.com/");
    expect(await button.getAttribute("class")).toMatch(/tile_.*/);
    expect(await button.findElement(By.css("h1")).getText()).toBe("CNN");
    expect(await button.findElement(By.css("img")).getAttribute("src")).toBe("http://media.idownloadblog.com/wp-content/uploads/2014/09/cnn-icon.png");
    takeScreenShot("ToolsAppsTabbed");
    await remove();
}

async function TestRemovingAllLinks() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    webpart.webPartData.properties.layout = "tabbed";
    await SetupWebpartLocal();
    let bigContainer = await (driver.findElement(By.css("div[class^='wsbThings']")));

    let buttonsContainer = await (bigContainer.findElement(By.css("div[class^='toolsAppsFlexContainer']")));
    let customizationButton = await buttonsContainer.findElement(By.css("button[class^='ms-Toggle-background']"));
    await customizationButton.click();

    let tabButtons = await bigContainer.findElements(By.css("div[role='presentation'] > ul > button"));
    for (let tabButton of tabButtons) {
        await tabButton.click();
        let hideButton = await driver.findElement(By.css("div[class^='toolsAppsFlexContainer'] div[class^='deleteButton'] i"));
        await hideButton.click();
    }

    customizationButton = await driver.findElement(By.css("div[class^='toolsAppsFlexContainer'] button[class^='ms-Toggle-background']"));
    await customizationButton.click();

    let updatedButtons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    expect(await updatedButtons).toEqual([]);
    let tabButton = await bigContainer.findElement(By.css("div[role='presentation'] > ul > button span[class^='ms-Pivot-text']"));
    expect(await tabButton.getText()).toBe("Uncategorized");
    takeScreenShot("ToolsAppsNoLinks");
    await remove();
}

async function TestIllegalProfileProperty() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    await SetupWebpartLocal();

    let editButton = await driver.findElement(By.css("button[data-automation-id='configureButton']"));
    await editButton.click();
    await timeout(standardWait);
    let layoutHeaderButton = await driver.findElement(By.xpath("//button[starts-with(@class, 'propertyPaneGroupHeader')]//div[contains(text(),'Layout')]"));
    await layoutHeaderButton.click();
    let propertyField = await driver.findElement(By.css("input[value='ToolsAppsCustomization']"));
    await propertyField.sendKeys(Key.chord(Key.CONTROL, 'a'), "Title");
    await timeout(standardWait);
    let errorField = await driver.findElement(By.css("span[data-automation-id='error-message']"));
    expect(await errorField.getText()).toBe("Prohibited: the property Title is a member of a built in group.");
    takeScreenShot("ToolsAppsIllegalProfileProperty");
    await remove();
}


async function SetupWebpartLocal(){

    await timeout(1000);
    var add = await driver.findElements(By.css('i[data-icon-name="Add"]'));
    await add[1].click();
    await timeout(1000);
   
    await driver.findElement(By.css('button[data-automation-id="5fb28714-f831-431c-b5cb-6f24a558f522_0"]')).click();
   
}


function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

//TODO: implement this better...
function pageLoaded() {
    return new Promise(resolve => setTimeout(resolve, 2500));
}
async function DiscardChanges(){

}
async function takeScreenShot(fileName: string) {
    var screenshot = await (driver.takeScreenshot());
    fs.writeFile("img/" + fileName + ".png", screenshot, 'base64', function (err) {
        if (!!err) { console.log(err) };
    });
}
async function remove() {
    var element = await (driver.findElement(By.xpath("//button[starts-with(@data-automation-id, 'deleteButton')]")));
    await element.click();
    await timeout(1000);
    element = await (driver.findElement(By.xpath("//button[starts-with(@data-automation-id, 'confirmButton')]")));
    await element.click();
    await timeout(1000);
}

async function layoutTests( title: string, layout: string){
    let lists = await driver.findElements(By.css('div[role="listbox"]'));
    await lists[0].click();
    await driver.findElement(By.css('button[title="List"]')).click();
    await timeout(500);
    await driver.findElement(By.css('button[data-automation-id="propertyPaneApplyButton"]')).click();
    
    
    //click on list
    await timeout(1000);
    
    lists[1].click();
    await timeout(500);
    //click on certain layout
    let dataIndex = 'button[title="'+title+'"]';
    await driver.findElement(By.css(dataIndex)).click();
    //apply
    await driver.findElement(By.css('button[data-automation-id="propertyPaneApplyButton"]')).click();
    await timeout(500);

    
    await lists[0].click();
    await driver.findElement(By.css('button[title="Tile"]')).click();
    await timeout(500);
    await driver.findElement(By.css('button[data-automation-id="propertyPaneApplyButton"]')).click();

    
    //check if results match
    let element = await driver.findElement(By.css('a[href="https://wsj.com"]'));
    element.getAttribute('class').then( attr =>  {
        expect(attr).toMatch(layout)
    })



}