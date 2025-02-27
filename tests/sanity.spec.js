import { test, expect } from '@playwright/test';
import Data from '../Test-Data/Data.json';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

const __dirname = path.resolve();

// Initialize the Excel workbook and worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Test Results');

// Define green tick and red cross symbols
const greenTick = '✅';
const redCross = '❌';

// Initialize the headers in the Excel sheet
worksheet.getCell('A1').value = 'Test Case Name';
worksheet.getCell('B1').value = 'Test Result';
worksheet.getCell('C1').value = 'RAG';

// Prepopulate the test case names in column A
const testNames = [
    'Login Test',
    'Single Prompt Test',
    'Multi Prompt Test',
    'AI Text Humanizer Test'
];

testNames.forEach((name, index) => {
    worksheet.getCell(`A${index + 2}`).value = name;
});

// Function to dynamically create a folder for the current month
const getMonthlyFolderPath = () => {
    const currentDate = new Date();
    const monthName = currentDate.toLocaleString('en-US', { month: 'long' }); // e.g., "April"
    const year = currentDate.getFullYear(); // e.g., "2025"
    const folderName = `${monthName}_${year}`; // e.g., "April_2025"
    const folderPath = path.join(__dirname, folderName);

    // Create the folder if it doesn't exist
    if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath, { recursive: true });
        console.log(`📁 Created folder: ${folderPath}`);
    }

    return folderPath;
};

// Function to finalize Excel file writing after all tests
const finalizeExcelFile = async () => {
    try {
        const currentDate = new Date();
        const formattedDate = `${String(currentDate.getDate()).padStart(2, '0')}${String(currentDate.getMonth() + 1).padStart(2, '0')}${currentDate.getFullYear()}`;
        const fileName = `Sanity_Test_Results_${formattedDate}.xlsx`;

        // Get month-wise folder path
        const folderPath = getMonthlyFolderPath();
        const filePath = path.join(folderPath, fileName);

        await workbook.xlsx.writeFile(filePath);
        console.log(`✅ Results written to ${filePath}`);
    } catch (error) {
        console.error(`Error finalizing Excel file: ${error.message}`);
    }
};

// Function to save test results in Excel
const saveResults = async (rowIndex, result, status) => {
    try {
        const row = worksheet.getRow(rowIndex + 2);
        row.getCell(2).value = result;
        row.getCell(3).value = status === 'pass' ? greenTick : redCross;
        row.commit();
    } catch (error) {
        console.error(`Error saving results: ${error.message}`);
    }
};

// Error handling function - Logs error but doesn't stop execution
const handlePageError = async (error, testCaseIndex) => {
    console.error(`Test Case Failed: ${testNames[testCaseIndex]}`);
    console.error(`Error: ${error.message}`);
    await saveResults(testCaseIndex, `❌ ${error.message}`, 'fail');
};

// **Main Test Suite**
test.describe('Testsuite', () => {
    let page;

    // Before all tests: Login once
    test.beforeAll(async ({ browser }) => {
        const context = await browser.newContext();
        page = await context.newPage();
        for (const [index, GW_Data] of Data.Test_Data.entries()) {
            try {
                console.log(`Navigating to URL: ${GW_Data.URL}`);
                await page.goto(GW_Data.URL, { waitUntil: 'domcontentloaded', timeout: 60000 });

                console.log('URL loaded successfully');

                const emailLocator = page.locator("#text");
                await emailLocator.waitFor({ state: 'visible', timeout: 10000 });
                await emailLocator.fill(GW_Data.Username);
                await page.waitForTimeout(2000);

                const passwordLocator = page.locator("#password");
                await passwordLocator.waitFor({ state: 'visible', timeout: 10000 });
                await passwordLocator.fill(GW_Data.Password);
                await page.waitForTimeout(2000);

                const button = page.locator("//button[@type='submit']");
                await button.waitFor({ state: 'visible', timeout: 10000 });
                await button.click();
                await page.waitForTimeout(2000);

                const continueButton = page.locator("//button[contains(text(),'Continue here')]");
                if (await continueButton.isVisible({ timeout: 5000 })) {
                    console.log("Continue button found, clicking...");
                    await continueButton.click();
                } else {
                    console.log("Continue button not visible, proceeding...");
                }
                await page.waitForTimeout(2000);

                const skipSurveyButton = page.locator("//span[contains(text(),'Skip survey')]");
                await skipSurveyButton.waitFor({ state: 'visible', timeout: 10000 });
                await skipSurveyButton.click();
                await page.waitForTimeout(2000);

                const popup1 = page.locator("//p[contains(text(),'Not Now')]");
                await popup1.waitFor({ state: 'visible', timeout: 10000 });
                await popup1.click();
                console.log("Not now button clicked...");
                console.log('Login successful for:', GW_Data.Username);
                await saveResults(0, "Login Passed", 'pass');

            } catch (error) {
                await handlePageError(error, 0);
            }
        }
    });
    // **Test Case 1: Single Prompt Test**
    test(`TC 1_Single_Prompt`, async () => {
        try {
            await page.click("//span[contains(text(),'Blog Title Generator')]");
            await page.fill('[placeholder="Please enter your focus keyword"]', Data.Test_Data[0].Single_Prompt_Q1);
            await page.fill('//input[@name="question_243"]', Data.Test_Data[0].Single_Prompt_Q2);
            await page.click("#button-submit");

            const response = await page.waitForResponse(res => res.url().includes('/api/singlePrompt/storeContent') && res.status() === 200);
            const responseData = await response.json();

            let singlePromptResult = responseData?.data?.content || '';
            singlePromptResult = typeof singlePromptResult === 'string' ? singlePromptResult.replace(/\*\*|-!"/g, '').trim() : '';

            if (singlePromptResult.length > 0) {
                console.log('\x1b[32m✔️ Single Prompt Test Passed!\x1b[0m');
                await saveResults(1, `Single prompt test passed:${singlePromptResult}`, 'pass');
            } else {
                throw new Error("Single prompt result is empty.");
            }
        } catch (error) {
            await handlePageError(error, 1);
        }
    });

    // **Test Case 2: Multi Prompt Test**
    test(`TC 2_Multi_Prompt`, async () => {
        try {
            await page.locator("(//span[contains(text(),'Home')])[1]").click();
            await page.locator('//div[@class="flex items-center gap-4"]').click();
            await page.locator("//div[contains(text(),'Kit Tools')]").click();
            await page.locator("//span[contains(text(),'LinkedIn Profile Builder')]").click();
            await page.locator('(//div[@class="flex-1 mt-2"])[2]').type(Data.Test_Data[0].Multi_Prompt_Q1);
            await page.locator("//button[contains(text(),'Save and Continue')]").click();
            await page.locator("(//span[contains(text(),'Regenerate')])[1]").click();

            const response = await page.waitForResponse(res =>
                res.url().includes('/api/multiPrompt/storeContent') && res.status() === 200
            );
            const responseData = await response.json();
            console.log('Full API Response:', JSON.stringify(responseData, null, 2));

            let multiPromptResult = responseData?.data?.content?.content || responseData?.data?.content || '';
            multiPromptResult = typeof multiPromptResult === 'string' ? multiPromptResult.replace(/\*\*|"/g, '').trim() : '';

            console.log(`Extracted Multi Prompt Result:`, multiPromptResult);
            console.log(`Character Count:`, multiPromptResult.length);

            if (multiPromptResult.length > 0) {
                console.log('\x1b[32m✔️ Multi Prompt Test Passed!\x1b[0m');
                await saveResults(2, `Multi prompt test passed with content: ${multiPromptResult}`, 'pass');
            } else {
                throw new Error("Multi prompt result is empty.");
            }
        } catch (error) {
            console.log('\x1b[31m❌ Multi Prompt Test Failed!\x1b[0m');
            await saveResults(2, `Multi prompt test failed: ${error.message}`, 'fail');
            await handlePageError(error, 2);
        }
    });

    test(`TC 3_AI_Text_Humanizer`, async () => {
        try {
            await page.click("//span[contains(text(),'AI Text Humanizer')]");
            await page.fill('//textarea[@placeholder="Please paste the AI-produced content that you\'d like to be revised to sound more human."]', Data.Test_Data[0].AI_Text_Humanizer);
            await page.click("//button[contains(text(),'Humanize')]");

            // Locate error message
            let errorLocator = page.locator('//p[@class="text-xs text-red-500 mt-1"]');

            // Check if the error message is displayed
            if (await errorLocator.isVisible()) {
                let errorMessage = await errorLocator.textContent();
                console.log(`Error Message Displayed: ${errorMessage}`);
                await saveResults(3, `AI Text Humanizer Test Failed: ${errorMessage}`, 'fail');
                throw new Error(`Validation error: ${errorMessage}`);
            } else {
                console.log("Hello World"); // Message when no error is found
                console.log("⏳ Waiting for API response: /api/singlePrompt/storeContent");

                const response = await page.waitForResponse(async (res) => {
                    const url = res.url();
                    if (url.includes('/api/singlePrompt/storeContent')) {
                        console.log(`✅ API Response Received: ${url}`);
                        return res.status() === 200;
                    }
                    return false;  // Keep waiting if the condition is not met
                });

                const responseData = await response.json();
                console.log('Full API Response:', JSON.stringify(responseData, null, 2));
                let singlePromptResult = responseData?.data?.content || '';
                singlePromptResult = typeof singlePromptResult === 'string' ? singlePromptResult.replace(/\*\*|-!"/g, '').trim() : '';

                if (singlePromptResult.length > 0) {
                    console.log('\x1b[32m✔️ Single Prompt Test Passed!\x1b[0m');
                    await saveResults(3, `AI Text Humanizer test passed:${singlePromptResult}`, 'pass');
                } else {
                    throw new Error("Single prompt result is empty.");
                }
            }
        } catch (error) {
            console.log('\x1b[31m❌ AI Text Humanizer Test Failed!\x1b[0m');
            await saveResults(3, `AI Text Humanizer Test Failed: ${error.message}`, 'fail');
            await handlePageError(error, 3);
        }
    });



    // After all tests: Finalize Excel report
    test.afterAll(async () => {
        await page.close();
        await finalizeExcelFile();
        console.log("\x1b[34mTest suite completed. Results saved.\x1b[0m");
    });
});

