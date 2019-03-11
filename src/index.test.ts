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

it('can set up a webpart', async () => { await TestWebpartSetup() })
it('edits fields in alerts', async () => { await TestFieldEditing() })
it('tests alerts default rendering', async () => { await TestDefaultRendering() })
it('tests alerts Juliet rendering', async () => { await TestJulietRendering() })
it('tests alerts Chicago rendering', async () => { await TestChicagoRendering() })
it('can setup tools and apps', async () => { await TestToolsAppsWebpartSetup() })
it('renders default mode', async () => { await TestTADefaultMode() })
it('can click t&a buttons', async () => { await TestTAButtons() })
it('can hide links', async () => { await TestTAUserCustomization() })
it('renders list view mode', async () => { await TestListMode() })
it('renders tabbed layout', async () => { await TestTabbedLayout() })
it('doesn\'t break when removing all tiles from the tabbed layout', async () => { await TestRemovingAllLinks() })
it('warns about illegal profile properties', async () => { await TestIllegalProfileProperty() })
it('closes the driver', async () => { await driver.close() })

async function TestWebpartSetup() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    var element = await (driver.findElement(By.xpath("//div[starts-with(@class, 'alerts')]")));
    var attrib = await (element.getAttribute("class"));

    expect(attrib).not.toBeNull;
}


async function TestFieldEditing() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    let interactionArea: WebElement[] = await driver.findElements(By.css("div[class^='alerts'] textarea"))
    await interactionArea[0].sendKeys(Key.chord(Key.CONTROL, 'a'), "here's a new test title");
    await interactionArea[1].sendKeys(Key.chord(Key.CONTROL, 'a'), "and some new button text");
    await interactionArea[2].sendKeys(Key.chord(Key.CONTROL, 'a'), "http://www.anewlink.com");
    await (timeout(standardWait));
    //We do this bc for some reason the dom isn't updated when you type in this webpart (even in a nomral browser) idk why. it's weird man.
    await (driver.navigate().refresh());
    await (timeout(standardWait));
    interactionArea = await driver.findElements(By.css("div[class^='alerts'] textarea"))
    expect(await interactionArea[0].getText()).toBe("here's a new test title");
    expect(await interactionArea[1].getText()).toBe("and some new button text");
    expect(await interactionArea[2].getText()).toBe("http://www.anewlink.com");

    takeScreenShot("AlertsFieldEditing");
}

async function TestDefaultRendering() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    await testingUtil.togglePublishMode();

    let alert = await driver.findElements(By.css("div[class^='alertViewMode'] > div"))
    expect(await alert[0].getText()).toBe("Alert that is really important 1234");
    let button = alert[1].findElement(By.css("a"))
    expect(await button.getText()).toBe("Learn More");
    expect(await button.getAttribute("href")).toBe("https://duckduckgo.com/");
    takeScreenShot("AlertsDefault");
}

async function TestJulietRendering() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    webpart.webPartData.properties.style = "Juliet";
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    await testingUtil.togglePublishMode();

    let alert = await driver.findElements(By.css("div[class^='alertViewMode'] > div"))
    expect(await alert[0].getText()).toBe("Alert that is really important 1234");
    let button = alert[1].findElement(By.css("a"))
    expect(await button.getText()).toBe("Learn More");
    expect(await button.getAttribute("href")).toBe("https://duckduckgo.com/");
    takeScreenShot("AlertsJuliet");
}

async function TestChicagoRendering() {
    let webpart = JSON.parse(JSON.stringify(defaultAlerts));
    webpart.webPartData.properties.style = "Chicago";
    webpart.webPartData.properties.imageUrl = "http://icons.iconarchive.com/icons/graphicloads/100-flat/256/home-icon.png";
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

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
}

async function TestToolsAppsWebpartSetup() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    var element = await (driver.findElement(By.xpath("//div[starts-with(@class, 'wsbThings')]")));
    var attrib = await (element.getAttribute("class"));

    expect(attrib).not.toBeNull;
}

async function TestTADefaultMode() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    let buttons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let button = buttons[0];
    expect(await button.getAttribute("href")).toBe("https://wsj.com/");
    expect(await button.getAttribute("class")).toMatch(/tile_.*/);
    expect(await button.findElement(By.css("h1")).getText()).toBe("Wallstreet Journal");
    expect(await button.findElement(By.css("img")).getAttribute("src")).toBe("https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67");
    takeScreenShot("ToolsAppsDefault");
}

async function TestTAButtons() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    let buttons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let button = buttons[0];
    await button.click();

    await pageLoaded();
    expect(await driver.getCurrentUrl()).toBe("https://www.wsj.com/");
}

async function TestTAUserCustomization() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

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
}

async function TestListMode() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.viewMode = "list";
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

    let buttons = await (driver.findElements(By.css("div[class^='toolsAppsFlexContainer'] > a")));
    let button = buttons[0];
    expect(await button.getAttribute("href")).toBe("https://wsj.com/");
    expect(await button.getAttribute("class")).toMatch(/list_.*/);
    expect(await button.findElement(By.css("h1")).getText()).toBe("Wallstreet Journal");
    expect(await button.findElement(By.css("img")).getAttribute("src")).toBe("https://camo.githubusercontent.com/3288d22efd14f228d106509b2b1e0d7ca28ce4e9/687474703a2f2f73666572696b2e6769746875622e696f2f77736a2f69636f6e2e706e67");
    takeScreenShot("ToolsAppsListMode");
}

async function TestTabbedLayout() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.layout = "tabbed";
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

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
}

async function TestRemovingAllLinks() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    webpart.webPartData.properties.layout = "tabbed";
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);
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
}

async function TestIllegalProfileProperty() {
    let webpart = JSON.parse(JSON.stringify(defaultToolsApps));
    webpart.webPartData.properties.isUserTileCustomizationAllowed = true;
    await SetupWebpartLocal(WebpartMaker.makeWebPart, webpart);

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
}





async function SetupWebpartLocal(creationScript: (any) => void, args: any) {
    await (driver.get("https://owdevelop.sharepoint.com/sites/OWv2-Develop/SitePages/AutomatedTestPage.aspx?Mode=Edit"))
    await driver.findElement(By.name("loginfmt")).sendKeys("owdeveloper@owdevelop.onmicrosoft.com");
    await driver.findElement(By.id("idSIButton9")).click();
    await driver.findElement(By.name("passwd")).sendKeys("&LEM3TF'wkNe4cCh");
    
    await driver.findElement(By.id("idSIButton9")).click();
    await driver.findElement(By.id("idSIButton9")).click();
    
    await (timeout(standardWait));
    await (driver.executeScript(creationScript, args));
    await (driver.navigate().refresh());
    
    await driver.findElement(By.name("Add")).click();
  
    
}

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

//TODO: implement this better...
function pageLoaded() {
    return new Promise(resolve => setTimeout(resolve, 2500));
}

async function takeScreenShot(fileName: string) {
    var screenshot = await (driver.takeScreenshot());
    fs.writeFile("img/" + fileName + ".png", screenshot, 'base64', function (err) {
        if (!!err) { console.log(err) };
    });
}