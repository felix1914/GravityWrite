const { test, expect } = require("@playwright/test");
const xlsx = require("xlsx");
const fs = require("fs");

// Define the Excel file path
const excelFilePath = "wsk_results.xlsx";

// Function to append data to the Excel file
async function appendToExcel(data) {
  let workbook;
  let worksheet;

  // Check if the Excel file already exists
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

  // Get the current rows in the worksheet
  const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // Insert data with headers
  const newData = data.map((row) => [
    row.planType,
    row.plan,
    row.price,
    row.expectedPrice,
    row.result,
    row.date,
  ]);
  const updatedData = [...existingData, ...newData];

  // Update the worksheet with the combined data
  const updatedWorksheet = xlsx.utils.aoa_to_sheet(updatedData);
  workbook.Sheets["Sheet1"] = updatedWorksheet;

  // Write the updated workbook to the file
  xlsx.writeFile(workbook, excelFilePath);
}

// Function to verify prices for a given page
async function verifyPrices(page, url, planType, expectedPrices, priceSelectors) {
  await page.goto(url);

  // Data to store the results of the price checks
  const priceData = [];

  // Helper function to extract price text and compare with expected price
  const checkPrice = async (planName, expectedPrice, priceSelector) => {
    const priceText = await page.locator(priceSelector).innerText();
    const actualPrice = parseInt(priceText.replace(/₹|\/month|\s/g, "")); // Remove symbols and parse as integer
    const result = actualPrice === expectedPrice ? "Pass" : "Fail";
    priceData.push({
      planType,
      plan: planName,
      price: actualPrice,
      expectedPrice,
      result,
    });
  };

  // Check each plan's price
  for (const [planName, selector] of Object.entries(priceSelectors)) {
    await checkPrice(planName, expectedPrices[planName], selector);
  }

  // Get the current date and time in 12-hour format
  const currentDate = new Date();
  const dateOptions = {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  };
  const formattedDate = currentDate
    .toLocaleString("en-GB", dateOptions)
    .replace(",", "");

  // Append the formatted date to each row
  priceData.forEach((row) => {
    row.date = formattedDate;
  });

  // Append data to Excel
  await appendToExcel(priceData);
}

// Main tests
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
      Starter: "//h3[contains(text(), 'Starter')]/following-sibling::h4",
      Basic: "//h3[contains(text(), 'Basic')]/following-sibling::h4",
      Premium: "//h3[contains(text(), 'Premium')]/following-sibling::h4",
      Ecommerce: "//h3[contains(text(), 'E-commerce')]/following-sibling::h4",
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
      Starter: "//h3[contains(text(), 'Starter')]/following-sibling::h4",
      Basic: "//h3[contains(text(), 'Basic')]/following-sibling::h4",
      Premium: "//h3[contains(text(), 'Premium')]/following-sibling::h4",
    };

    await verifyPrices(page, url, planType, expectedPrices, priceSelectors);
  });

  test("Verify E-commerce Hosting Prices", async ({ page }) => {
    const url = "https://webspacekit.com/ecommerce-plans/";
    const planType = "E-commerce Hosting";
    const expectedPrices = {
      "E-commerce": 599,
      "E-commerce Plus": 799,
    };
    const priceSelectors = {
      "E-commerce":
        "//div[contains(@class, 'pricing-card') and .//h3[contains(text(), 'E-commerce') and not(contains(text(), 'Plus'))]]//h4",
      "E-commerce Plus":
        "//div[contains(@class, 'pricing-card') and .//h3[contains(text(), 'E-commerce Plus')]]//h4",
    };

    await verifyPrices(page, url, planType, expectedPrices, priceSelectors);
  });
});
