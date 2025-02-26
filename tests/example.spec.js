import { test } from '@playwright/test';
const ExcelJS = require('exceljs');

test('AgentLoginPage with Authentication', async ({ page }) => {
  try {
    const sitemapURL = 'https://uat.truhomefinance.in/sitemap.xml';

    // Authenticate using basic HTTP credentials
    await page.authenticate({
      username: 'hfl',
      password: 'Hfl@123',
    });

    await page.goto(sitemapURL);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sitemap');

    // Add headers to the Excel worksheet
    worksheet.addRow(['URL', 'Response Code']);

    // Evaluate an XPath expression to extract values of <loc> tags
    const locValues = await page.$$eval('loc', (elements) =>
      elements.map((element) => element.textContent)
    );

    for (const locValue of locValues) {
      const response = await page.goto(locValue);
      const responseCode = response?.status();
      console.log(`Response Code: ${locValue}---${responseCode}`);
      worksheet.addRow([locValue, responseCode]);
    }

    const path = 'Shriramlife05062024.xlsx';
    await workbook.xlsx.writeFile(path);

    console.log('Sitemap data has been written to the Excel file successfully.');
  } catch (error) {
    console.error('An error has occurred ', error);
  }
});
