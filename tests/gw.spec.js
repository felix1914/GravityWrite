import { test, chromium } from '@playwright/test';
import data from '../tests/data.json';
const ExcelJS = require('exceljs');

// Initialize the Excel workbook and worksheet outside of the loop
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Test Results');

// Define headers in the first column (column A)
worksheet.getCell('A1').value = 'URL';
worksheet.getCell('A2').value = 'Username & Password Function';
worksheet.getCell('A3').value = 'URL Redirection';
worksheet.getCell('A4').value = 'Title Input Result';
worksheet.getCell('A5').value = 'Create Outline Button Result';

// Iterate over each data item and perform tests
data.data.forEach((item, index) => {
    
    let browser;
    let context;
    let page;
    let Urlredirection;
    let loginpage;
    let url = item.URL;
    let username = item.Username;
    let password = item.Password;
    let titleinputresult;
    let createoutlinebtn;

    const highlightElement = async (element) => {
        await element.evaluate(el => {
            el.style.border = '5px solid';
            el.style.borderImage = 'linear-gradient(to right, red, orange, yellow, green, blue, indigo, violet) 1';
        }); // Highlight the element
        await page.waitForTimeout(1000); // Wait for 1 second
        await element.evaluate(el => el.style.border = ''); // Remove highlight
    };

    test.describe(`Tests for URL: ${item.URL}`, () => {

        test.beforeAll(async () => {
            try {
                console.log("Starting browser and opening page");
                browser = await chromium.launch();
                context = await browser.newContext();
                page = await context.newPage();
                await page.goto(url);
                console.log("Page opened successfully");

            } catch (error) {
                console.error(`Error during setup: ${error.message}`);
                test.fail(`Setup failed for URL: ${url}`);
            }
        });

        test(`Login Page Test`, async () => {
            try {
                console.log("Filling in login details");
                const emailLocator = page.locator("#text");
                await emailLocator.click();
                await emailLocator.fill(username);

                const passwordLocator = page.locator("#password");
                await highlightElement(passwordLocator);
                await passwordLocator.click();
                await passwordLocator.fill(password);

                const button = page.locator("//button[@type='submit']");
                await button.click();

                const loginOtherPageInfo = await page.locator("//div[contains(@class, 'w-11/12') and contains(@class, 'mx-auto') and contains(@class, 'bg-white') and contains(@class, 'rounded-lg') and contains(@class, 'md:w-auto')]").first();

                if (await loginOtherPageInfo) {
                    console.log("Handling login from another browser");
                    const continueButton = page.locator("//button[contains(text(),'Continue here')]");
                    await continueButton.click();
                    console.log("Page logged in from another browser");
                    const skipSurveyButton = page.locator("//span[contains(text(),'Skip survey')]");
                    await skipSurveyButton.click();

                    const tryNowButton = page.locator("(//button[contains(text(),'Try Now')])[2]");
                    await tryNowButton.click();
                    loginpage = "User name password successfully entered";
                } else {
                    loginpage = "Username password section issue";
                    console.log("Username password section issue");
                    test.fail(`Username password section issue: ${error.message}`);
                }

            } catch (error) {
                loginpage = `Username and password section issue: ${error.message}`;
                console.error(`Login test failed: ${error.message}`);
                test.fail(`Login test failed: ${error.message}`);
            }
        });

        test('URL checking ', async () => {
            try {
                let currentURL = page.url();
                let actualUrl = "https://react-dev.gravitywrite.com/blog-writer";
                console.log(currentURL);
                console.log(actualUrl);
                console.log(`Final URL: ${currentURL}`);

                if (currentURL == actualUrl) {
                    Urlredirection = "Page logged in successfully with blogwriter url";
                    console.log("Login test successful");
                } else {
                    Urlredirection = "Error: Unexpected URL after login";
                    console.error("Error: Unexpected URL after login");
                    test.fail("Unexpected URL after login");
                }

            } catch (error) {
                Urlredirection = `Login test failed: ${error.message}`;
                console.error(`Fail: ${error.message}`);
                test.fail(`Blog writer test failed: ${error.message}`);
            }
        });

        test(`AI Blog Writer Title Input Test`, async () => {
            try {
                console.log("Entering blog title");
                let blogtitle = item.Blog_Title;
                let titleInput = page.locator("#title");
                await titleInput.click();
                await titleInput.fill(blogtitle);
                titleinputresult = "The input box value successfully entered";
                console.log("Blog title input successful");

            } catch (error) {
                titleinputresult = `Fail: ${error.message}`;
                console.error(`Fail: ${error.message}`);
                test.fail(`AI Blog Writer Title Input Test: ${error.message}`);
            }
        });

        test(`AI Blog Writer CreateOutline button Test`, async () => {
            try {
                console.log("Clicking Create Outline button");
                await page.locator("//button[contains(text(),'Create Outline')]").click();
                createoutlinebtn = "Create outline button successfully clicked";
                console.log("Create outline button clicked successfully");
            } catch (error) {
                createoutlinebtn = `Fail: ${error.message}`;
                console.error(`Fail: ${error.message}`);
                test.fail(`AI Blog Writer CreateOutline button Test: ${error.message}`);
            }
        });

        test.afterAll(async () => {
            try {
                console.log("Saving results and closing browser");
                console.log(`Login page Result: ${loginpage}, ${Urlredirection}, ${titleinputresult}, ${createoutlinebtn}`);

                // Insert the results starting from column B for each test run
                const column = String.fromCharCode(66 + index); // Convert 66 (ASCII code for 'B') + index to a letter for the column

                worksheet.getCell(`${column}1`).value = url;
                worksheet.getCell(`${column}2`).value = loginpage;
                worksheet.getCell(`${column}3`).value = Urlredirection;
                worksheet.getCell(`${column}4`).value = titleinputresult;
                worksheet.getCell(`${column}5`).value = createoutlinebtn;

                console.log("Results saved successfully");
                const path = 'TestResults.xlsx';
                await workbook.xlsx.writeFile(path);
                console.log(`Results written to ${path}`);

                // Close the browser after all tests are done
                if (browser) {
                    await browser.close();
                    console.log("Browser closed");
                }

            } catch (error) {
                console.error(`Error during cleanup: ${error}`);
                test.fail(`Cleanup failed: ${error.message}`);
            }
        });
    });
});
