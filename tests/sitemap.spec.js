import { test }  from '@playwright/test';
const ExcelJS = require('exceljs');
const fs = require('fs');
test('AgentLoginPage ',async ({page})=>{
    try {
 
    const sitemapURL = 'sitemap.xml'; 
    await page.goto(sitemapURL);
    // test.setTimeout(7000000);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sitemap');
 
    // Add headers to the Excel worksheet
    worksheet.addRow(['URL', 'Response Code']); 
    // Evaluate an XPath expression to extract values of <loc> tags
    const locValues = await page.$$eval('loc', (elements) => elements.map((element) => element.textContent));
 
    for (const locValue of locValues) {
   
   
      const response = await page.goto(locValue);
      const responseCode = response?.status();
      console.log(`Response Code: ${locValue}---${responseCode}`);
      worksheet.addRow([locValue, responseCode]);
      const path = 'Shriramlife05062024.xlsx';
      await workbook.xlsx.writeFile(path);

    }
}catch (error) {
    console.error('An error has occurred ', error);
  }

});

 