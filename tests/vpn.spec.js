const { test, expect } = require("@playwright/test");
const xlsx = require("xlsx");
const fs = require("fs");

const excelFilePath = "wsk_results.xlsx";

// Function to append data to the Excel file
async function appendToExcel(data) {
  let workbook;
  let worksheet;

  if (fs.existsSync(excelFilePath)) {
    workbook = xlsx.readFile(excelFilePath);
    worksheet = workbook.Sheets["Sheet1"];
  } else {
    workbook = xlsx.utils.book_new();
    worksheet = xlsx.utils.aoa_to_sheet([
      ["Plan Type", "Plan Name", "Actual Price", "Expected Price", "Result", "Date"],
    ]);
    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  }

  const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  const newData = data.map((row) => [
    row.planType,
    row.plan,
    row.price,
    row.expectedPrice,
    row.result,
    row.date,
  ]);
  const updatedData = [...existingData, ...newData];

  const updatedWorksheet = xlsx.utils.aoa_to_sheet(updatedData);
  workbook.Sheets["Sheet1"] = updatedWorksheet;
  xlsx.writeFile(workbook, excelFilePath);
}

// Function to verify prices
async function verifyPrices(page, url, planType, expectedPrices, priceSelectors) {
  console.log(`Navigating to: ${url}`);
  await page.goto(url, { waitUntil: "load", timeout: 60000 });

  try {
    // Close popup if present
    const closePopupSelector = "//div[@tabindex='0' and contains(@class, 'brave_popup__close')]";
    if (await page.locator(closePopupSelector).isVisible()) {
      await page.locator(closePopupSelector).click();
      console.log("Popup closed successfully.");
    }
  } catch (error) {
    console.error("Popup closure failed:", error.message);
  }

  let scrollXPath = "//h3[contains(text(),'Starter')]";
  await page.locator(scrollXPath).scrollIntoViewIfNeeded();
  await page.waitForSelector(scrollXPath, { state: "visible", timeout: 30000 });

  const priceData = [];

  for (const [planName, selector] of Object.entries(priceSelectors)) {
    try {
      await page.waitForSelector(selector);
      let priceText = await page.locator(selector);
      let actualPrice = parseInt(priceText.replace(/₹|\/month|\s/g, ""));

      let result = actualPrice === expectedPrices[planName] ? "Pass" : "Fail";
      console.log(`Plan: ${planName}, Expected: ${expectedPrices[planName]}, Actual: ${actualPrice}, Result: ${result}`);

      priceData.push({
        planType,
        plan: planName,
        price: actualPrice,
        expectedPrice: expectedPrices[planName],
        result,
      });
    } catch (error) {
      console.error(`Error fetching price for ${planName}:`, error.message);
      priceData.push({
        planType,
        plan: planName,
        price: "N/A",
        expectedPrice: expectedPrices[planName],
        result: "Fail",
      });
    }
  }

  const formattedDate = new Date().toLocaleString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  }).replace(",", "");

  priceData.forEach((row) => (row.date = formattedDate));
  await appendToExcel(priceData);
}

test.describe("Hosting Plans Price Verification", () => {
  test("Verify WordPress Hosting Prices", async ({ page }) => {
    const url = "https://webspacekit.com/wordpress-hosting-plans/";
    const planType = "WordPress Hosting";

    const expectedPrices = {
      Starter: 149,
      Basic: 199,
      Premium: 269,
      Ecommerce: 599,
    };

    const priceSelectors = {
      Starter: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[1]",
      Basic: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[2]",
      Premium: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[3]",
      Ecommerce: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[4]",
    };

    await verifyPrices(page, url, planType, expectedPrices, priceSelectors);
  });

  test("Verify Web Hosting Prices", async ({ page }) => {
    const url = "https://webspacekit.com/web-hosting-plans/";
    const planType = "Web Hosting";

    const expectedPrices = {
      Starter: 149,
      Basic: 199,
      Premium: 269,
    };

    const priceSelectors = {
      Starter: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[1]",
      Basic: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[2]",
      Premium: "(//div[@class='dcc-content']//h4/span[@class='currency']/following-sibling::text())[3]",
    };

    await verifyPrices(page, url, planType, expectedPrices, priceSelectors);
  });
});
