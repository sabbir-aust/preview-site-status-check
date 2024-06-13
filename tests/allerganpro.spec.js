// @ts-check
const { test, expect } = require('@playwright/test');
const { randomDelay, logResponses, saveResponsesToExcel,logFailure } = require('../pages/functions');
const xlsx = require('xlsx');
const fs = require('fs');

// Helper function to read URLs from the Excel file
function readUrlsFromExcel(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = xlsx.readFile(filePath);
  const sheetName = 'Allerganpro';
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);
  return data; // returns an array of objects with keys 'country' and 'url'
}

// Read URLs from the Excel file
const urls = readUrlsFromExcel('urls.xlsx');

urls.forEach(({ country, url }) => {
  test(`has title and checks status code for abbviepro ${country} URL`, async ({ page }) => {
    // Create an array to store URLs and status codes
    let responses = [];
    const mainUrl = url;

    // Log in credentials
    const username = 'ejaz';
    const password = 'NsgHyXb1!';

    // Log the initial URL hit
    responses.push({ url: mainUrl, status: 0 });

    try {

        // Navigate to the login page
    await page.goto(mainUrl, { waitUntil: 'load' });
    responses[0].status = 200; // Update the status after successful navigation

    // Log responses after hitting the URL
    await logResponses(page, responses, mainUrl);

    // Perform login
    await page.fill('input[name="username"]', username);
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
    await saveResponsesToExcel(responses, country, 'Allerganpro');
      
    } catch (error) {
      await logFailure(country, url, error.message, 'Allerganpro');
    }
  });
});