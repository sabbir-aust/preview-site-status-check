const { test, expect } = require("@playwright/test");
const xlsx = require('xlsx');
const fs = require('fs');

// Utility function to add a random delay
function randomDelay(min = 1000, max = 3000) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

async function logResponses(page, responses, mainUrl) {
  page.on('response', async (response) => {
    const url = response.url();
    const status = response.status();
    console.log(`URL: ${url} - Status: ${status}`);

    let statusCode = status;
    try {
      const responseBody = await response.json();
      if (responseBody && responseBody.data && responseBody.data.code && responseBody.data.code !== 200) {
        statusCode = responseBody.data.code;
      }
    } catch (e) {
      // Not a JSON response or failed to parse, ignore the error
    }

    responses.push({ url, status: statusCode });

    if (url === mainUrl && statusCode !== 302 && statusCode !== 429 && statusCode !== 404) {
      expect(statusCode).toBe(200);
    }
  });
}

async function saveResponsesToExcel(responses, sheetName, fileName) {
  let filePath;
  switch (fileName) {
    case 'Abbvie':
      filePath = './report/abbviepro_responses.xlsx';
      break;
    case 'Support':
      filePath = './report/sitegen_responses.xlsx';
      break;
    case 'ADPA':
      filePath = './report/adpa_responses.xlsx';
      break;
    case 'TORA':
      filePath = './report/tora_responses.xlsx';
      break;
    case 'TOTP':
      filePath = './report/totp_responses.xlsx';
      break;  
    case 'Care':
      filePath = './report/abbviecare_responses.xlsx';
      break;
    case 'Medical':
      filePath = './report/abbviemedical_responses.xlsx';
      break;
    case 'Allerganpro':
      filePath = './report/allerganpro_responses.xlsx';
      break;
    case 'AMI':
      filePath = './report/ami_responses.xlsx';
      break;
    case 'LMBC':
      filePath = './report/lmbc_responses.xlsx';
      break;
    case 'Sitegen':
      filePath = './report/sitegen_responses.xlsx';
      break;
    case 'ACUITI':
      filePath = './report/acuiti_responses.xlsx';
      break;  
    default:
      throw new Error(`Unknown file name: ${fileName}`);
  }

  console.log(`Saving responses to ${filePath} in sheet "${sheetName}"`);

  let workbook;
  if (fs.existsSync(filePath)) {
    console.log('File exists. Reading the existing file.');
    workbook = xlsx.readFile(filePath);
  } else {
    console.log('File does not exist. Creating a new workbook.');
    workbook = xlsx.utils.book_new();
  }

  let existingData = [];
  if (workbook.SheetNames.includes(sheetName)) {
    console.log(`Sheet "${sheetName}" exists. Reading existing data.`);
    const worksheet = workbook.Sheets[sheetName];
    existingData = xlsx.utils.sheet_to_json(worksheet);
  } else {
    console.log(`Sheet "${sheetName}" does not exist. It will be created.`);
  }

  console.log(`Appending ${responses.length} new responses to the sheet.`);
  const newData = existingData.concat(responses);

  const ws = xlsx.utils.json_to_sheet(newData);
  workbook.Sheets[sheetName] = ws;
  if (!workbook.SheetNames.includes(sheetName)) {
    xlsx.utils.book_append_sheet(workbook, ws, sheetName);
  }

  console.log(`Writing data to file: ${filePath}`);
  xlsx.writeFile(workbook, filePath);
  console.log('Data successfully written to file.');
}

async function logFailure(country, url, errorMessage, statusCode, projectName) {
  const filePath = './error_log_report/failed_cases.xlsx';
  const sheetName = projectName;

  console.log(`Logging failure for ${country} - ${url}`);

  let workbook;
  if (fs.existsSync(filePath)) {
    console.log('Failed cases file exists. Reading the existing file.');
    workbook = xlsx.readFile(filePath);
  } else {
    console.log('Failed cases file does not exist. Creating a new workbook.');
    workbook = xlsx.utils.book_new();
  }

  let existingData = [];
  if (workbook.SheetNames.includes(sheetName)) {
    console.log(`Sheet "${sheetName}" exists. Reading existing data.`);
    const worksheet = workbook.Sheets[sheetName];
    existingData = xlsx.utils.sheet_to_json(worksheet);
  } else {
    console.log(`Sheet "${sheetName}" does not exist. It will be created.`);
  }

  const newFailure = { country, url, errorMessage, statusCode, timestamp: new Date().toISOString() };
  existingData.push(newFailure);

  const ws = xlsx.utils.json_to_sheet(existingData);
  workbook.Sheets[sheetName] = ws;
  if (!workbook.SheetNames.includes(sheetName)) {
    xlsx.utils.book_append_sheet(workbook, ws, sheetName);
  }

  console.log(`Writing failed case data to file: ${filePath}`);
  xlsx.writeFile(workbook, filePath);
  console.log('Failed case data successfully written to file.');
}

module.exports = { randomDelay, logResponses, saveResponsesToExcel, logFailure };
