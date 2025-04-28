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
    'Date',
    'Login Test',
    'Single Prompt Test',
    'Multi Prompt Test',
    'AI Text Humanizer Test',
    'AI_Chat',
    'AI_Image_Generate',
    'Storybook_Generate',
    'Storybook_CoverImage'
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
        const fileName = `Staging_Test_Results_${formattedDate}.xlsx`;

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

                // const popup1 = page.locator("//p[contains(text(),'Not Now')]");
                // await popup1.waitFor({ state: 'visible', timeout: 10000 });
                // await popup1.click();
                console.log("Not now button clicked...");
                console.log('Login successful for:', GW_Data.Username);
                await saveResults(1, "Login Passed", 'pass');

            } catch (error) {
                await handlePageError(error, 1);
            }
        }
    });

    test('Test Execution Date', async () => {
        let formattedDate;
        try {
            // Create a new Date object
            let currentDate = new Date();

            // Format the date as day/month/year
            let day = String(currentDate.getDate()).padStart(2, '0'); // Ensure two digits for day
            let month = String(currentDate.getMonth() + 1).padStart(2, '0'); // Months are zero-based, so add 1
            let year = currentDate.getFullYear();

            formattedDate = `${day}/${month}/${year}`; // Format as day/month/year

            // Print the current date
            console.log("Current Date (DD/MM/YYYY):", formattedDate);

            await saveResults(0, formattedDate, 'pass'); // Save the result with formatted date
        } catch (error) {
            // Handle error
            console.error('❌ Error during the test:', error.message);
            await saveResults(0, `❌ ${error.message}`, 'fail');
        }
        await page.waitForTimeout(2000); // 2-second timeout
    });


    // **Test Case 1: Single Prompt Test**
    test(`TC 1_Single_Prompt`, async () => {
        try {
            await page.click("//span[contains(text(),'Blog Title Generator')]");
            await page.fill('[placeholder="Please enter your focus keyword"]', Data.Test_Data[0].Single_Prompt_Q1);
            await page.fill('//input[@name="question_243"]', Data.Test_Data[0].Single_Prompt_Q2);
            await page.click("#button-submit");

            console.log("⏳ Waiting for API response: /api/singlePrompt/storeContent");

            // **Wait for API Response**
            const response = await page.waitForResponse(res => res.url().includes('/api/singlePrompt/storeContent'));

            // Get response status code
            const statusCode = response.status();

            // Extract full response
            const responseData = await response.json();
            const fullResponse = JSON.stringify(responseData, null, 2);

            // Print response details
            console.log(`🔹 API Status Code: ${statusCode}`);
            console.log('🔹 Full API Response:', fullResponse);

            if (statusCode === 200) {
                let singlePromptResult = responseData?.data?.content || null;

                if (typeof singlePromptResult === 'string') {
                    singlePromptResult = singlePromptResult.replace(/\*\*|-!"/g, '').trim(); // Clean unwanted characters
                }

                if (singlePromptResult && singlePromptResult.length > 0) {
                    console.log('\x1b[32m✔️ Single Prompt Test Passed!\x1b[0m');
                    await saveResults(2, `Single prompt test passed: ${singlePromptResult}`, 'pass');
                } else {
                    throw new Error(`Single prompt result is empty. Full Response:\n${fullResponse}`);
                }
            } else {
                throw new Error(`❌ API returned ${statusCode} - Error Message: ${fullResponse}`);
            }

        } catch (error) {
            console.log('\x1b[31m❌ Single Prompt Test Failed!\x1b[0m');
            console.error(`❌ Error: ${error.message}`);
            await handlePageError(error, 2);
        }
    });

    // **Test Case 2: Multi Prompt Test**
    test.skip(`TC 2_Multi_Prompt`, async () => {
        try {
            await page.locator("(//span[contains(text(),'Home')])[1]").click();
            await page.locator('//div[@class="flex items-center gap-4"]').click();
            await page.locator("//div[contains(text(),'Kit Tools')]").click();
            await page.locator("//span[contains(text(),'LinkedIn Profile Builder')]").click();
            console.log(Data.Test_Data[0].Multi_Prompt_Q1);
            await page.locator('(//div[@class="flex-1 mt-2"])[2]').click();
            await page.locator('(//div[@class="flex-1 mt-2"])[2]').type(Data.Test_Data[0].Multi_Prompt_Q1);

            await page.locator("//button[contains(text(),'Save and Continue')]").click();
            await page.locator("(//span[contains(text(),'Regenerate')])[1]").click();

            console.log("⏳ Waiting for API response: /api/multiPrompt/storeContent");

            // **Wait for API Response**
            const response = await page.waitForResponse(res => res.url().includes('/api/multiPrompt/storeContent'));

            // Get response status code
            const statusCode = response.status();

            // Extract full response
            const responseData = await response.json();
            const fullResponse = JSON.stringify(responseData, null, 2);

            // Print response details
            console.log(`🔹 API Status Code: ${statusCode}`);
            console.log('🔹 Full API Response:', fullResponse);

            if (statusCode === 200) {
                let multiPromptResult = responseData?.data?.content?.content || responseData?.data?.content || null;

                if (typeof multiPromptResult === 'string') {
                    multiPromptResult = multiPromptResult.replace(/\*\*|"/g, '').trim(); // Clean unwanted characters
                }

                if (multiPromptResult && multiPromptResult.length > 0) {
                    console.log('\x1b[32m✔️ Multi Prompt Test Passed!\x1b[0m');
                    await saveResults(3, `Multi prompt test passed with content: ${multiPromptResult}`, 'pass');
                } else {
                    throw new Error(`Multi prompt result is empty. Full Response:\n${fullResponse}`);
                }
            } else {
                throw new Error(`❌ API returned ${statusCode} - Error Message: ${fullResponse}`);
            }

        } catch (error) {
            console.log('\x1b[31m❌ Multi Prompt Test Failed!\x1b[0m');
            console.error(`❌ Error: ${error.message}`);
            await handlePageError(error, 3);
        }
    });



    test(`TC 3_AI_Text_Humanizer`, async () => {
        try {
            //test.setTimeout(120000);
            await page.click("//span[contains(text(),'AI Text Humanizer')]");
            await page.fill('//textarea[@placeholder="Please paste the AI-produced content that you\'d like to be revised to sound more human."]', Data.Test_Data[0].AI_Text_Humanizer);
            await page.click("//button[contains(text(),'Humanize')]");

            console.log("Hello World"); // Message when no error is found
            console.log("⏳ Waiting for API response: /api/singlePrompt/storeContent");

            // **Wait for API Response**
            const response = await page.waitForResponse(res => res.url().includes('/api/singlePrompt/storeContent'));

            // Get response status code
            const statusCode = response.status();

            // Extract full response
            const responseData = await response.json();
            const fullResponse = JSON.stringify(responseData, null, 2);

            // Print response details
            console.log(`🔹 API Status Code: ${statusCode}`);
            console.log('🔹 Full API Response:', fullResponse);

            if (statusCode === 200) {
                let singlePromptResult = responseData?.data?.content || null;

                if (typeof singlePromptResult === 'string') {
                    singlePromptResult = singlePromptResult.replace(/\*\*|-!"/g, '').trim(); // Clean unwanted characters
                }

                if (singlePromptResult && singlePromptResult.length > 0) {
                    console.log('\x1b[32m✔️ AI Text Humanizer Test Passed!\x1b[0m');
                    await saveResults(4, `AI Text Humanizer test passed: ${singlePromptResult}`, 'pass');
                } else {
                    throw new Error(`AI Text Humanizer result is empty. Full Response:\n${fullResponse}`);
                }
            } else {
                throw new Error(`❌ API returned ${statusCode} - Error Message: ${fullResponse}`);
            }

        } catch (error) {
            console.log('\x1b[31m❌ AI Text Humanizer Test Failed!\x1b[0m');
            console.error(`❌ Error: ${error.message}`);
            await handlePageError(error, 4);
        }
    });



    test(`TC 4_AI_Chat`, async () => {
        try {
            await page.locator("//span[contains(text(),'AI Chat')]").click();
            await page.locator('//textarea[@name="banner-search"]').type(Data.Test_Data[0].AI_Chat);
            await page.keyboard.press('Enter');

            console.log("⏳ AI Chat clicked, waiting for API response: /api/chat/storeMessage");

            // **Wait for API Response**
            const response = await page.waitForResponse(res => res.url().includes('/api/chat/storeMessage'));

            // Get response status code
            const statusCode = response.status();

            // Extract full response
            const responseData = await response.json();
            const fullResponse = JSON.stringify(responseData, null, 2);

            // Print response details
            console.log(`🔹 API Status Code: ${statusCode}`);
            console.log('🔹 Full API Response:', fullResponse);

            if (statusCode === 200) {
                // Extract `message_output` correctly
                let chatResponse = responseData?.data?.messages?.message_output || responseData?.data?.message_output || null;

                if (typeof chatResponse === 'string') {
                    chatResponse = chatResponse.replace(/\*\*|["\n\r]/g, '').trim(); // Remove unwanted characters
                }

                console.log(`📌 Extracted Chat Response:`, chatResponse);
                console.log(`🔢 Character Count:`, chatResponse ? chatResponse.length : 0);

                if (chatResponse && chatResponse.length > 0) {
                    console.log('\x1b[32m✔️ AI Chat Test Passed!\x1b[0m');
                    await saveResults(5, `AI Chat test passed with content: ${chatResponse}`, 'pass');
                } else {
                    throw new Error(`AI Chat result is empty. Full Response:\n${fullResponse}`);
                }
            } else {
                throw new Error(`❌ API returned ${statusCode} - Error Message: ${fullResponse}`);
            }

        } catch (error) {
            console.log('\x1b[31m❌ AI Chat Test Failed!\x1b[0m');
            console.error(`❌ Error: ${error.message}`);
            await handlePageError(error, 5);
        }
    });



    test(`TC 5_AI_Image`, async () => {
        try {
            await page.locator("//span[contains(text(),'AI Image Generator')]").click();
            await page.locator('(//button[@aria-label="Create image prompt with AI for better results"])[1]').click();

            let popupCreate = await page.locator('//div[@class="text-[#1E2022] p-6 relative flex flex-col overflow-hidden"]');
            if (await popupCreate) {
                console.log("✅ Popup is visible");

                await page.locator("(//div[contains(@class, 'field-container')]//input[@type='text'])[1]").fill(Data.Test_Data[0].AI_Image_Q1);
                // await page.locator("(//div[contains(@class, 'field-container')]//input[@type='text'])[2]").fill(Data.Test_Data[0].AI_Image_Q2);
                // await page.locator("(//div[contains(@class, 'field-container')]//input[@type='text'])[3]").fill(Data.Test_Data[0].AI_Image_Q3);
                // test.setTimeout(2000);


                // await page.locator("//div[contains(text(),'Advanced Options')]").click();
                // test.setTimeout(2000);

                // await page.locator("//p[contains(text(),'Model B')]").click();
                // test.setTimeout(2000);

                // await page.locator("//div[contains(text(),'Model A')]").click();
                // test.setTimeout(2000);

                // await page.locator("(//p[contains(text(),'Realistic')])[2]").click();
                // test.setTimeout(2000);

                // await page.locator("//p[contains(text(),'General')]").click();
                // test.setTimeout(2000);


                await page.locator("//span[contains(text(),'Generate Image')]").click();
                console.log("✅ Generate Image button clicked");

                // **Wait for API Response**
                const response = await page.waitForResponse(res => res.url().includes('/api/ImgGenerate'));

                // Get response status code
                const statusCode = response.status();

                // Extract full response
                const responseData = await response.json();
                const fullResponse = JSON.stringify(responseData, null, 2);

                // Print response details
                console.log(`🔹 API Status Code: ${statusCode}`);
                console.log('🔹 Full API Response:', fullResponse);

                if (statusCode === 200) {
                    const imageUrl = responseData?.data?.image_url || null;
                    // const compressedImageUrl = responseData?.data?.compressed_image_url || null;

                    if (imageUrl) {
                        console.log('✅ Image URL:', imageUrl);
                        //   console.log('✅ Compressed Image URL:', compressedImageUrl);
                        await saveResults(6, `AI Image test passed with content:\n${imageUrl}`, 'pass');
                    } else {
                        throw new Error(`❌ Missing Image URLs. Full Response:\n${fullResponse}`);
                    }
                } else {
                    throw new Error(`❌ API returned ${statusCode} - Error Message: ${fullResponse}`);
                }
            } else {
                throw new Error("❌ Popup is not displayed");
            }
        } catch (error) {
            console.log('\x1b[31m❌ AI Image Test Failed!\x1b[0m');
            console.error(`❌ Error: ${error.message}`);
            //await saveResults(5, `AI Image test failed: ${error.message}`, 'fail');
            await handlePageError(error, 6);
        }
    });



    test('StoryBook', async () => {
        try {
            // Wait for page reload to complete and the page to be loaded
            // await page.reload({ waitUntil: 'domcontentloaded' });
            test.setTimeout(3000);
            // Wait for the 'AI Image Generator' element to be visible and then click it
            await page.waitForSelector("//span[contains(text(),'AI Image Generator')]", { state: 'visible' });
            await page.locator("//span[contains(text(),'AI Image Generator')]").click();
            test.setTimeout(3000);
            // await page.waitForSelector("//span[contains(text(),'AI Image Generator')]", { state: 'visible' });
            // await page.locator("//span[contains(text(),'AI Image Generator')]").click();

            // Click on 'Storybook Image Creator's
            await page.waitForSelector("//span[contains(text(),'Storybook Image Creator')]", { state: 'visible' });
            await page.locator("//span[contains(text(),'Storybook Image Creator')]").click();
            await page.waitForSelector("//button[contains(text(),'+ Create New Storybook Images')]", { state: 'visible' });
            await page.locator("//button[contains(text(),'+ Create New Storybook Images')]").click();
            

            // Fill in project details
            let project = await page.locator('#projectName');
            await project.click();
            await project.fill(Data.Test_Data[0].Project_Name);

            // Fill in the story details
            let story = await page.locator('#story');
            await story.click();
            await story.fill(Data.Test_Data[0].Story);

            // Submit the form
            await page.locator('//button[@type="submit"]').click();

            // Wait for the popup to appear
            await page.locator('//div[@class="flex flex-col justify-center items-center text-[#1E2022] p-4 md:p-12"]').waitFor();

            // Wait for the "Generate All Images" button to become visible
            await page.locator("//button[contains(text(),'Generate All Images')]").waitFor({ state: 'visible' });

            // Now click the button
            await page.locator("(//button[contains(text(),'Generate Images')])[1]").click();

            // Get the current URL
            const currentUrl = page.url();
            const urlObj = new URL(currentUrl);
            const storybookId = urlObj.searchParams.get("storybook_id");

            // Now use it in your API wait
            const response = await page.waitForResponse(res =>
                res.url().includes(`/api/storyboard/${storybookId}/generate/images`)
            );

            // Get response status code
            const statusCode = response.status();

            // Extract full response
            const responseData = await response.json();
            const fullResponse = JSON.stringify(responseData, null, 2);

            // Print response details
            console.log(`🔹 API Status Code: ${statusCode}`);
            console.log('🔹 Full API Response:', fullResponse);

            if (statusCode === 200) {
                const imageUrl = responseData?.data?.story_images?.[0]?.image_url; // Corrected this line

                if (imageUrl) {
                    console.log('✅ Image URL:', imageUrl);
                    await saveResults(7, `AI Image test passed with content:\n${imageUrl}`, 'pass');
                } else {
                    throw new Error(`❌ Missing image_url in API response.\nFull Response:\n${fullResponse}`);
                }
            } else {
                throw new Error(`❌ API returned status ${statusCode}.\nError Message:\n${fullResponse}`);
            }

        } catch (error) {
            console.error("Error during the test:", error);
            await handlePageError(error, 7);
        }
    });

    test('CoverImage', async () => {
        let author = Data.Test_Data[0].CoverimgAuthor;
        try {

            await page.locator("//button[contains(text(),'Generate Cover Image')]").click();
            await page.locator("#authorName").fill(author);
            await page.click('//input[@value="create-prompt"]');
            await page.locator('//button[@type="submit"]').click();
            await page.locator("//button[contains(text(),'Generate Cover Images')]").click();

            const currentUrl = page.url();
            const urlObj = new URL(currentUrl);
            const storybookId = urlObj.searchParams.get("storybook_id");

            // Now use it in your API wait
            const response = await page.waitForResponse(res =>
                res.url().includes(`/api/storyboard/${storybookId}/cover-image`)
            );

            // Get response status code
            const statusCode = response.status();

            // Extract full response
            const responseData = await response.json();
            const fullResponse = JSON.stringify(responseData, null, 2);

            // Print response details
            console.log(`🔹 API Status Code: ${statusCode}`);
            console.log('🔹 Full API Response:', fullResponse);

            if (statusCode === 200) {
                const imageUrl = responseData?.data?.story_images?.[0]?.image_url; // Corrected this line

                if (imageUrl) {
                    console.log('✅ Image URL:', imageUrl);
                    await saveResults(8, `AI Image test passed with content:\n${imageUrl}`, 'pass');
                } else {
                    throw new Error(`❌ Missing image_url in API response.\nFull Response:\n${fullResponse}`);
                }
            } else {
                throw new Error(`❌ API returned status ${statusCode}.\nError Message:\n${fullResponse}`);
            }


        } catch (error) {
            console.error("Error during the test:", error);
            await handlePageError(error, 8);

        }



    })









    // After all tests: Finalize Excel report



    test.afterAll(async () => {
        console.log("\x1b[34mFinalizing results...\x1b[0m");
        await finalizeExcelFile(); // Write results at the end of all tests
        console.log("\x1b[34mTest suite completed. Results saved.\x1b[0m");
    });
});

