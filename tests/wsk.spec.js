const { test, expect, chromium } = require('@playwright/test'); // Import chromium
const ExcelJS = require('exceljs'); // Import ExcelJS
const data = require('../tests/data.json'); // Import test data

// Initialize ExcelJS workbook and worksheet
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
        const path = 'TestResults.xlsx';
        await workbook.xlsx.writeFile(path);
        console.log(`Results written to ${path}`);  
    } catch (error) {
        console.error(`Error during finalizing Excel file: ${error}`);
    }
};

// Prepopulate the worksheet with test names in the first columnzˀ
const testNames = ['Url validation', 'User_Type'];
testNames.forEach((name, index) => {
    worksheet.getCell(`A${index + 1}`).value = name;
});

// Error handling function
const handlePageError = async (error, rowIndex, testCaseIndex) => {
    try {
        console.error(`Error: ${error.message}`);
        await saveResults(rowIndex, `Error: ${error.message}`, testCaseIndex);
    } catch (errorHandlingError) {
        console.error(`Error during error handling: ${errorHandlingError.message}`);
        await saveResults(rowIndex, `Error during error handling: ${errorHandlingError.message}`, testCaseIndex);
    } finally {
        await finalizeExcelFile();
        process.exit(1); // Stop further execution
    }
};
data.data.forEach((item, testCaseIndex) => {
    const { URL, Domain } = item; 
// Test suite for running tests serially
test.describe.serial('URL and Domain Tests', () => {
    let browser, context, page;

    // Initialize browser, context, and page before all tests
    test.beforeAll(async () => {
        browser = await chromium.launch();
        context = await browser.newContext();
        page = await context.newPage();
    });

    // Close browser and context after all tests
    test.afterAll(async () => {
        await page.close();
        await context.close();
        await browser.close();
        await finalizeExcelFile(); // Finalize the Excel file after all tests
    });

    // Sequentially run tests for each data item
  // Destructure Domain from item

        // Define first test
        test(`URL Test ${testCaseIndex}`, async () => {
            try {
                // Navigate to the specified URL
                await page.goto(URL, { waitUntil: 'networkidle' });

                // Log "Done" to the console after the page has loaded
                console.log("Done");
                await saveResults(0, "Test1 Passed", testCaseIndex); // Save results for test case 1

                // Add assertions here if needed
                // Example: expect(await page.title()).toBe('Expected Title');
            } catch (error) {
                await handlePageError(error, 0, testCaseIndex);
            }
        });

        // Define second test
        test(`Domain Click ${testCaseIndex}`, async () => {
            try {
                // Interact with the page
                const domaininput = page.locator('#domainName');
                await domaininput.click();
                await domaininput.fill(Domain); // Fill the domain value
                await page.locator("//span[contains(text(),'Search')]").click();
                await saveResults(1, "Test2 Passed", testCaseIndex); // Save results for test case 2
            } catch (error) {
                await handlePageError(error, 1, testCaseIndex);
            }
        });
        test(`Exsiting User ${testCaseIndex}`, async () => {
            try {
                // Interact with the page
                const user= await page.locator("//span[contains(text(),'Existing Customer Login')]"); 
               user.click();
console.log("Working fine");
                await saveResults(2, "Test2 Passed", testCaseIndex); // Save results for test case 2
            } catch (error) {
                await handlePageError(error, 2, testCaseIndex);
            }
        });
    });
});
