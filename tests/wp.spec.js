import { test, chromium } from '@playwright/test';
import wpData from './wp.json';  // Import wp.json
import ExcelJS from 'exceljs';

// Initialize the Excel workbook and worksheet outside of the loop
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Test Results');

// Function to save results in the desired format
const saveResults = async (rowIndex, result, testCaseIndex) => {
    try {
        const row = worksheet.getRow(rowIndex + 1); // +1 because row index is 1-based in ExcelJS
        row.getCell(testCaseIndex + 2).value = result; // Start from column B (index + 2)
        row.commit();
    } catch (error) {
        console.error(`Error during saving results: ${error}`);
    }
};

// Function to finalize the Excel file writing process
const finalizeExcelFile = async () => {
    try {
        const path = 'TestResultswp1102024.xlsx';
        await workbook.xlsx.writeFile(path);
        console.log(`Results written to ${path}`);
    } catch (error) {
        console.error(`Error during finalizing Excel file: ${error}`);
    }
};

const apidatalogin = async (page) => {
    try {
        const response = await page.waitForResponse(
            (res) =>
                res.status() === 200 &&
                res.url() === 'https://staging-api.gravitywrite.com/api/auth/login'
        );
        const responseData = await response.json();
        return responseData;
    } catch (error) {
        console.error(`Error during fetching API data: ${error}`);
    }
};
const apidatagetuser = async (page) => {
    try {
        const response = await page.waitForResponse(
            (res) =>
                res.status() === 200 &&
                res.url() === 'https://staging-api.gravitywrite.com/api/auth/getUser'
        );
        const responseData = await response.json();
        return responseData;
    } catch (error) {
        console.error(`Error during fetching API data: ${error}`);
    }
};

// Prepopulate the worksheet with test names in the first column
const testNames = [
    'Test Data for Execution',
    'URL Launch',
    'Login Page Test',
    'User Subscription',
    'URL Checking Test',
    'Title Input Test',
    'Create Outline Button Test',
    'Image Source Validation',
    'Add New Subheading',
    'Generate Keyword with AI',
    'Outline Continue Button Test',
    'Generate All Button Test',
    'Image Validation Test',
    'Preview Button Test',
    'Download Button Test'
];

testNames.forEach((name, index) => {
    worksheet.getCell(`A${index + 1}`).value = name;
});


// Iterate over each item in wpData.data and perform tests
wpData.data.forEach((item, testCaseIndex) => {
    const { Username, Password, URL, Newsubheading, Blog_Title, Keyword } = item;

    test.describe.serial(`Tests for URL: ${URL}`, () => {
        let browser, context, page;

        test.beforeAll(async () => {
            browser = await chromium.launch();
            context = await browser.newContext();
            page = await context.newPage();
        });

        test(`Test Data for Execution ${testCaseIndex}`, async () => {
            try {
                await saveResults(0, JSON.stringify(wpData), testCaseIndex);
                console.log(wpData);
            } catch (error) {
                await handlePageError(error, page, saveResults, 0, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`URL Launch ${testCaseIndex}`, async () => {
            try {
                const response = await page.goto(URL);
                const responseCode = response?.status();
                console.log(`Response Code: ---${responseCode}`);
                await saveResults(1, `URL Response: ${responseCode}`, testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 1, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Login Page Test ${testCaseIndex}`, async () => {
            try {
                const emailLocator = page.locator("#text");
                await emailLocator.waitFor({ state: 'visible', timeout: 10000 });
                await emailLocator.fill(Username);

                const passwordLocator = page.locator("#password");
                await passwordLocator.waitFor({ state: 'visible', timeout: 10000 });
                await passwordLocator.fill(Password);

                const button = page.locator("//button[@type='submit']");
                await button.waitFor({ state: 'visible', timeout: 10000 });
                await button.click();

                const loginOtherPageInfo = page.locator("//div[contains(@class, 'w-11/12') and contains(@class, 'mx-auto') and contains(@class, 'bg-white') and contains(@class, 'rounded-lg') and contains(@class, 'md:w-auto')]").first();
                await loginOtherPageInfo.waitFor({ state: 'visible', timeout: 10000 });

                if (await loginOtherPageInfo) {
                    const continueButton = page.locator("//button[contains(text(),'Continue here')]");
                    await continueButton.waitFor({ state: 'visible', timeout: 10000 });
                    await continueButton.click();

                    const responseData = await apidatalogin(page);
                    console.log('User Subscription:', responseData.data.user.user_subscription.plan.title);
                    let usertype = responseData.data.user.user_subscription.plan.title;
                    console.log();

                    const offerButton = page.locator("//p[contains(text(),'No thanks!')]");
                    if (usertype !== "Pro") {
                        await offerButton.waitFor({ state: 'visible', timeout: 10000 });
                        await offerButton.click();
                    }

                    const skipSurveyButton = page.locator("//span[contains(text(),'Skip survey')]");
                    await skipSurveyButton.waitFor({ state: 'visible', timeout: 10000 });
                    await skipSurveyButton.click();

                    const tryNowButton = page.locator("(//button[contains(text(),'Try Now')])[2]");
                    await tryNowButton.waitFor({ state: 'visible', timeout: 10000 });
                    await tryNowButton.click();

                    await saveResults(2, "Username and password successfully entered", testCaseIndex);

                } else {
                    throw new Error("Login page did not load correctly");
                }
            } catch (error) {
                await handlePageError(error, page, saveResults, 2, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`User Subscription ${testCaseIndex}`, async () => {
            try {
                await page.reload();
                test.setTimeout(3000);
                const responseData = await page.waitForResponse(
                    (res) => res.status() === 200 &&
                        res.url() === 'https://staging-api.gravitywrite.com/api/auth/getUser'
                );
                if (responseData) {
                    const userData = await responseData.json();
                    console.log('User Subscription:', userData.data.user.user_subscription.plan.title);
                    await saveResults(3, userData.data.user.user_subscription.plan.title, testCaseIndex);
                }
            } catch (error) {
                await handlePageError(error, page, saveResults, 3, testCaseIndex, finalizeExcelFile);
            }
        });

        test(`URL Checking Test ${testCaseIndex}`, async () => {
            try {
                const currentURL = page.url();
                const expectedURL = "https://staging.gravitywrite.com/blog-writer";
                if (currentURL === expectedURL) {
                    await saveResults(4, "Page logged in successfully with blog writer URL", testCaseIndex);
                } else {
                    throw new Error("Unexpected URL after login");
                }
            } catch (error) {
                await handlePageError(error, page, saveResults, 4, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Title Input Test ${testCaseIndex}`, async () => {
            try {
                const titleInput = page.locator("#title");
                await titleInput.click();
                await titleInput.fill(Blog_Title);
                await saveResults(5, "Title input box value successfully entered", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 5, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Create Outline Button Test ${testCaseIndex}`, async () => {
            try {
                await page.locator("//button[contains(text(),'Create Outline')]").click();
                await saveResults(6, "Create outline button successfully clicked", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 6, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Image Source Validation ${testCaseIndex}`, async () => {
            try {
                const imageSourceElement = await page.locator('(//div[@class="w-full bg-white rounded-md simple-dropdown"])[3]');
                const source = await imageSourceElement.textContent();
                console.log(source);
                await saveResults(7, source, testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 7, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000);
        });

        test(`Add New Subheading ${testCaseIndex}`, async () => {
            try {
                await page.locator('(//div[@class="relative inline pr-[50px]"])[1]').hover();
                await page.locator('(//div[@aria-label="Add sub-heading"])[1]').click();
                const inputField = await page.locator('(//div[@class="relative inline pr-[50px]"])[2]');
                await inputField.scrollIntoViewIfNeeded();
                await inputField.waitFor({ state: 'visible' });
                await inputField.type(Newsubheading);
                await saveResults(8, "Subheading entered successfully", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 8, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Generate Keyword with AI ${testCaseIndex}`, async () => {
            try {
                const keyword = await page.locator('//textarea[@rows="3"]');
                await keyword.scrollIntoViewIfNeeded();
                await keyword.waitFor({ state: 'visible' });
                await keyword.fill(Keyword);
                const keydata = await page.locator("//button[contains(text(),'Generate keywords with ai')]");
                await keydata.click();
                await page.waitForTimeout(3000);
                console.log("done");
                await saveResults(9, "Generate Keywords with AI successfully", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 9, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Outline Continue Button Test ${testCaseIndex}`, async () => {
            try {
                await page.locator("(//button[contains(text(),'Continue')])[2]").click();
                await saveResults(10, "Continue button successfully clicked", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 10, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Generate All Button Test ${testCaseIndex}`, async () => {
            try {
                await page.locator("//button[contains(text(),'Generate All')]").click();
                await saveResults(11, "Generate All button successfully clicked", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 11, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Image Validation Test ${testCaseIndex}`, async () => {
            try {
                await page.waitForSelector('//img[@class="max-h-[120px] max-w-[170px] cursor-pointer"]', { state: 'visible' });
                const image = await page.locator('//img[@class="max-h-[120px] max-w-[170px] cursor-pointer"]');
                const imageSrc = await image.getAttribute('src');
                console.log(`Image src: ${imageSrc}`);
                const imageResponse = await page.request.get(imageSrc);
                const responseCode = imageResponse.status();
                console.log(`Image Response Code: ${responseCode}`);
                await page.waitForTimeout(5000);
                if (!page.isClosed()) {
                    const stopGenerationButton = page.locator("//span[contains(text(),'Stop Generation')]");
                    await stopGenerationButton.click();
                    await saveResults(12, `Image generated successfully, src:    ${imageSrc}, Response Code: ${responseCode}`, testCaseIndex);
                } else {
                    throw new Error("The page was closed before clicking the Stop Generation button.");
                }
            } catch (error) {
                if (!page.isClosed()) {
                    await handlePageError(error, page, saveResults, 12, testCaseIndex, finalizeExcelFile);
                } else {
                    console.error("Error during test execution: ", error);
                }
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        test(`Preview Button Test ${testCaseIndex}`, async () => {
            try {
                const preview = await page.locator("//button[contains(text(),'Preview and Export')]");
                await preview.waitFor({ state: 'visible' });
                if (page.isClosed()) {
                    throw new Error("The page was closed before clicking the preview button.");
                }
                await preview.click();
                console.log("Preview button successfully clicked");
                await saveResults(13, "Preview button successfully clicked", testCaseIndex);
            } catch (error) {
                await handlePageError(error, page, saveResults, 13, testCaseIndex, finalizeExcelFile);
            }
            await page.waitForTimeout(2000); // 2-second timeout
        });

        // test(`Download Button Test ${testCaseIndex}`, async () => {
        //     try {
        //         console.log("Waiting for the popup to become visible...");
        //         const popup = await page.locator("//div[contains(text(),'Export Blog Post')]");
        //         await popup.waitFor({ state: 'visible', timeout: 10000 });
        //         console.log("Popup is visible. Locating the download button...");
        //         const downloadButton = await page.locator("//button[contains(text(),'Download')]");
        //         console.log("Waiting for the download button to become visible...");
        //         await downloadButton.waitFor({ state: 'visible', timeout: 10000 });
        //         console.log("Download button is visible. Clicking the download button...");
        //         await downloadButton.click();
        //         console.log("Download button successfully clicked");
        //         await saveResults(14, "Download button successfully clicked", testCaseIndex);
        //     } catch (error) {
        //         console.error("Error encountered:", error);
        //         await handlePageError(error, page, saveResults, 14, testCaseIndex, finalizeExcelFile);
        //     }
        //     await page.waitForTimeout(2000); // 2-second timeout
        // });

        test.afterAll(async () => {
            if (browser) {
                await browser.close();
                console.log("Browser closed after all tests");
                await finalizeExcelFile();
            }
        });
    });
});
