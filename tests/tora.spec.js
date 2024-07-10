// @ts-check
const { test, expect } = require('@playwright/test');
const { randomDelay, logResponses, saveResponsesToExcel, logFailure } = require('../pages/functions');
const xlsx = require('xlsx');
const fs = require('fs');

// Helper function to read URLs from the Excel file
function readUrlsFromExcel(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = xlsx.readFile(filePath);
  const sheetName = 'TORA';
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);
  return data; // returns an array of objects with keys 'country' and 'url'
}

// Read URLs from the Excel file
const urls = readUrlsFromExcel('urls.xlsx');

urls.forEach(({ country, url }) => {
  test(`has title and checks status code for TORA ${country} URL`, async ({ page }) => {
    // Create an array to store URLs and status codes
    let responses = [];
    const mainUrl = url;

    // Log in credentials
    const username = 'ejaz';
    const password = 'NsgHyXb1!';

    try {
      // Navigate to the login page
      await page.goto(mainUrl, { waitUntil: 'load' });
      responses.push({ url: mainUrl, status: 200 }); // Update the status after successful navigation

      // Log responses after hitting the URL
      await logResponses(page, responses, mainUrl);

      // Perform login
      await page.fill('input[name="username"]', username, { timeout: 10000 }); // Increase timeout if necessary
      await page.fill('input[name="password"]', password);
      await page.click('button[type="submit"]');

      // Wait for the page to fully load
      await page.waitForLoadState('networkidle');

      // Log responses after login
      await logResponses(page, responses, mainUrl);

      // Additional manual wait if necessary
      await page.waitForTimeout(randomDelay(3000, 5000));

      // Expect the title to contain at least one non-whitespace character
      await expect(page).toHaveTitle(/\S+/);

      // Save responses to Excel file in a sheet named after the country
      await saveResponsesToExcel(responses, country, 'TORA');
      
    } catch (error) {
      console.error(`Error testing ${country} - ${mainUrl}: ${error}`);
      const statusCode = error.message.includes('Timeout' || 'Test timeout') ? 404 : (responses.length > 0 ? responses[responses.length - 1].status : 404);
      await logFailure(country, mainUrl, error.message, statusCode, 'TORA');

    } finally {
      // If no entries were recorded, log an unknown status entry
      if (responses.length === 0) {
        await logFailure(country, mainUrl, 'No response recorded', 404, 'TORA');
      }
    }
  });
});
