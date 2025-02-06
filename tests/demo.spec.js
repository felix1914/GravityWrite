import { test, chromium } from '@playwright/test';
import data from '../tests/data.json';
import ExcelJS from 'exceljs';

// Initialize the Excel workbook and worksheet outside of the loop
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Test Results');

// Add headers to the Excel sheet
worksheet.addRow(['URL', 'Username & Password Function', 'URL Redirection', 'Title Input Result', 'Create Outline Button Result', 'Add new sub heading']);

// Helper function to highlight an element
const highlightElement = async (page, element) => {
    await element.evaluate(el => {
        el.style.border = '5px solid';
        el.style.borderImage = 'linear-gradient(to right, red, orange, yellow, green, blue, indigo, violet) 1';
    });
    await page.waitForTimeout(1000); // Wait for 1 second
    await element.evaluate(el => el.style.border = ''); // Remove highlight
};

// Helper function to save results to Excel
const saveResults = async (rowValues) => {
    try {
        worksheet.addRow(rowValues);
        const path = 'demo.xlsx';
        await workbook.xlsx.writeFile(path);
        console.log(`Results written to ${path}`);
    } catch (error) {
        console.error(`Error during saving results: ${error}`);
    }
};

// Iterate over each data item and perform tests
data.data.forEach((item) => {
    const { Username, Password, URL, Newsubheading, Blog_Title } = item;

    test.describe(`Tests for URL: ${URL}`, () => {
        let browser, context, page;
        let loginResult, urlRedirectionResult, titleInputResult, createOutlineResult, addNewSubResult;

        test.beforeAll(async () => {
            try {
                browser = await chromium.launch();
                context = await browser.newContext();
                page = await context.newPage();
                await page.goto(URL);
                console.log("Page opened successfully");
            } catch (error) {
                console.error(`Error during setup: ${error.message}`);
                await saveResults([URL, 'Setup Error']);
            }
        });

        test(`Login Page Test`, async () => {
            await test.step('Performing Login', async () => {
                try {
                    const emailLocator = page.locator("#text");
                    await emailLocator.click();
                    await emailLocator.fill(Username);

                    const passwordLocator = page.locator("#password");
                    await highlightElement(page, passwordLocator);
                    await passwordLocator.click();
                    await passwordLocator.fill(Password);

                    const button = page.locator("//button[@type='submit']");
                    await button.click();

                    const loginOtherPageInfo = page.locator("//div[contains(@class, 'w-11/12') and contains(@class, 'mx-auto') and contains(@class, 'bg-white') and contains(@class, 'rounded-lg') and contains(@class, 'md:w-auto')]").first();

                    if (await loginOtherPageInfo) {
                        const continueButton = page.locator("//button[contains(text(),'Continue here')]");
                        await continueButton.click();

                        const skipSurveyButton = page.locator("//span[contains(text(),'Skip survey')]");
                        await skipSurveyButton.click();

                        const tryNowButton = page.locator("(//button[contains(text(),'Try Now')])[2]");
                        await tryNowButton.click();
                        loginResult = "Username and password successfully entered";
                    } else {
                        loginResult = "Username password section issue";
                    }
                    const currentURL = page.url();  // Get the current URL after login attempt
                    await saveResults([currentURL, loginResult]);
                } catch (error) {
                    const currentURL = page.url();  // Get the current URL even if there's an error
                    loginResult = `Username and password section issue: ${error.message}`;
                    await saveResults([currentURL, loginResult]);
                }
            });
        });

        test('URL Checking Test', async () => {
            await test.step('Checking URL Redirection', async () => {
                try {
                    const currentURL = page.url();
                    const expectedURL = "https://react-dev.gravitywrite.com/blog-writer";

                    if (currentURL === expectedURL) {
                        urlRedirectionResult = "Page logged in successfully with blogwriter URL";
                    } else {
                        urlRedirectionResult = "Error: Unexpected URL after login";
                    }
                    await saveResults([currentURL, loginResult, urlRedirectionResult]);
                } catch (error) {
                    const currentURL = page.url();  // Get the current URL even if there's an error
                    urlRedirectionResult = `Login test failed: ${error.message}`;
                    await saveResults([currentURL, loginResult, urlRedirectionResult]);
                }
            });
        });

        test('AI Blog Writer Title Input Test', async () => {
            await test.step('Entering Blog Title', async () => {
                try {
                    const titleInput = page.locator("#title");
                    await titleInput.click();
                    await titleInput.fill(Blog_Title);
                    titleInputResult = "Title input box value successfully entered";
                    const currentURL = page.url();  // Get the current URL
                    await saveResults([currentURL, loginResult, urlRedirectionResult, titleInputResult]);
                } catch (error) {
                    const currentURL = page.url();  // Get the current URL even if there's an error
                    titleInputResult = `Fail: ${error.message}`;
                    await saveResults([currentURL, loginResult, urlRedirectionResult, titleInputResult]);
                }
            });
        });

        test('AI Blog Writer Create Outline Button Test', async () => {
            await test.step('Clicking Create Outline Button', async () => {
                try {
                    await page.locator("//button[contains(text(),'Create Outline')]").click();
                    createOutlineResult = "Create outline button successfully clicked";
                    const currentURL = page.url();  // Get the current URL
                    await saveResults([currentURL, loginResult, urlRedirectionResult, titleInputResult, createOutlineResult]);
                } catch (error) {
                    const currentURL = page.url();  // Get the current URL even if there's an error
                    createOutlineResult = `Fail: ${error.message}`;
                    await saveResults([currentURL, loginResult, urlRedirectionResult, titleInputResult, createOutlineResult]);
                }
            });
        });

        test('AI Blog Writer Outline Section Add New Subheading', async () => {
            await test.step('Adding New Subheading', async () => {
                try {
                    await page.locator('(//div[@class="relative inline pr-[50px]"])[1]').hover();
                    await page.locator('(//div[@aria-label="Add sub-heading"])[1]').click();

                    const inputField = await page.locator('(//div[@class="relative inline pr-[50px]"])[2]');
                    await highlightElement(page, inputField);
                    await inputField.scrollIntoViewIfNeeded();
                    await inputField.waitFor({ state: 'visible' });

                    await inputField.type(Newsubheading);
                    addNewSubResult = "Subheading entered successfully";
                    const currentURL = page.url();  // Get the current URL
                    await saveResults([currentURL, loginResult, urlRedirectionResult, titleInputResult, createOutlineResult, addNewSubResult]);
                } catch (error) {
                    const currentURL = page.url();  // Get the current URL even if there's an error
                    addNewSubResult = `Fail: ${error.message}`;
                    await saveResults([currentURL, loginResult, urlRedirectionResult, titleInputResult, createOutlineResult, addNewSubResult]);
                }
            });
        }, 60000);

        test.afterAll(async () => {
            if (browser) {
                await browser.close();
                console.log("Browser closed after all tests");
            }
        });

    });
});
