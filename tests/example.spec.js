import { test, chromium } from "@playwright/test";
import ExcelJS from 'exceljs';

// Initialize Excel workbook and worksheet outside the loop to ensure they are defined
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Test Results');

test.describe.serial('Playwright Tests with Excel Logging', () => {

    const testNames = [
        'Url validation',
        'User_Type'
    ];

    console.log(testNames);
    const length = testNames.length;
    console.log(length);

    // Write test names to the first column of the worksheet
    testNames.forEach((name, index) => {
        worksheet.getCell(`A${index + 1}`).value = name;
    });

    // Browser and page variables
    let browser, page;

    // Setup and teardown for browser
    test.beforeAll(async () => {
        browser = await chromium.launch();
        const context = await browser.newContext();
        page = await context.newPage();
    });

    test.afterAll(async () => {
        await browser.close();
        // Save the workbook to a file after all tests
        const path = 'Testexample.xlsx';
        await workbook.xlsx.writeFile(path);
        console.log(`Results written to ${path}`);
    });

    // Loop to define and execute tests
    for (let index = 0; index < length; index++) {
        const element = testNames[index];
        console.log(`Element at index ${index}: ${element}`);

        test(`Url launch for ${element}`, async () => {
            try {
                const URL = 'https://staging.gravitywrite.com/login'; // Replace with the actual URL for testing
                const response = await page.goto(URL);
                const responseCode = response?.status();
                console.log(`Response Code: ---${responseCode}`);
                console.log("Page opened successfully");
                worksheet.getCell(`B${index + 1}`).value = "hero"; 
            } catch (error) {
                worksheet.getCell(`B${index + 1}`).value = "hero1"; 
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Login Page Test for ${element}`, async () => {
            try {
                const Username = 'test@example.com'; // Replace with actual username
                const Password = 'password'; // Replace with actual password

                const emailLocator = page.locator("#text");
                await emailLocator.click();
                await emailLocator.fill(Username);
                const passwordLocator = page.locator("#password");
                await passwordLocator.click();
                await passwordLocator.fill(Password);
                const button = page.locator("//button[@type='submit']");
                await button.click();

                const loginOtherPageInfo = page.locator("//div[contains(@class, 'w-11/12') and contains(@class, 'mx-auto') and contains(@class, 'bg-white') and contains(@class, 'rounded-lg') and contains(@class, 'md:w-auto')]").first();
                if (await loginOtherPageInfo.isVisible()) {
                    const continueButton = page.locator("//button[contains(text(),'Continue here')]");
                    await continueButton.click();

                    const [response] = await Promise.all([
                        page.waitForResponse(
                            (res) =>
                                res.status() == 200 &&
                                res.url() === 'https://dev.gravitywrite.com/api/auth/login'
                        ),
                    ]);
                    const responseData = await response.json();
                    console.log('User Subscription:', responseData.data.user.user_subscription.plan.title);
                    let usertype = responseData.data.user.user_subscription.plan.title;

                    const offerbutton = page.locator("//p[contains(text(),'No thanks!')]");
                    if (usertype !== "Pro" && await offerbutton.isVisible()) {
                        await offerbutton.click();
                    }

                    const skipSurveyButton = page.locator("//span[contains(text(),'Skip survey')]");
                    await skipSurveyButton.click();
                    const tryNowButton = page.locator("(//button[contains(text(),'Try Now')])[2]");
                    await tryNowButton.click();
                    worksheet.getCell(`B${index + 1}`).value = "fax"; 
                } else {
                    worksheet.getCell(`B${index + 1}`).value = "fail"; 
                }
            } catch (error) {
                worksheet.getCell(`B${index + 1}`).value = "fax1"; 
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });
    }
});
