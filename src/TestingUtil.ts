import { ThenableWebDriver, By } from 'selenium-webdriver';
import { AutoConfigurable } from "./Enums";

export class TestingUtil {

    driver: ThenableWebDriver;

    constructor(driver: ThenableWebDriver) {
        this.driver = driver
    }

    //TODO
    public async Make(webpart: AutoConfigurable): Promise<void> {
        var element = await (this.driver.findElement(By.xpath("//button[starts-with(@data-automation-id, 'toolboxHintButton')]")));
        await element.click();
    }

    //TODO
    public async Configure(webpart: AutoConfigurable, ...args: any[]): Promise<void> {

    }

    public async makeAndConfigure(webpart: AutoConfigurable, ...args: any[]): Promise<void> {
        await this.Make(webpart);
        return this.Configure(webpart, ...args);
    }

    public async remove(webpart: AutoConfigurable, index: number = 0) {
        var element = await (this.driver.findElement(By.xpath("//button[starts-with(@data-automation-id, 'deleteButton')]")));
        await element.click();
        await this.pause(1000);
        element = await (this.driver.findElement(By.xpath("//button[starts-with(@data-automation-id, 'confirmButton')]")));
        await element.click();
        await this.pause(1000);
    }

    public async togglePublishMode() {
        var element = await (this.driver.findElement(By.css("button#CommandBar1preview")));
        return element.click();
    }

    private async pause(time: number) {
        return new Promise(resolve => setTimeout(resolve, time));
    }
}