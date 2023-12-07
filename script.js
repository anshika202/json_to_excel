const fs = require('fs');
const XLSX = require('xlsx');

// function to convert json to excel
function convertJson_dataToExcel(jsonData, workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName] || XLSX.utils.json_to_sheet([]);
  XLSX.utils.sheet_add_json(sheet, jsonData, { skipHeader: true, origin: -1 });
  workbook.Sheets[sheetName] = sheet;
}

// function to process nested objects and arrays
function processNestedObj(data, workbook, sheetName) {
  if (Array.isArray(data)) {
    // for arrays
    data.forEach((item, index) => {
      const newSheetName = `${sheetName}_Array${index + 1}`;
      convertJson_dataToExcel(item, workbook, newSheetName);
    });
  } else if (typeof data === 'object') {
    // for nested objects
    Object.entries(data).forEach(([key, value]) => {
      const newSheetName = `${sheetName}_${key}`;
      convertJson_dataToExcel(value, workbook, newSheetName);
    });
  }
}

// function to read json file 
function convertJsonFileToExcel(inputFilePath, outputFilePath) {
  try {
    // Read JSON file
    const jsonData = JSON.parse(fs.readFileSync(inputFilePath, 'utf8'));

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Convert nested JSON to Excel
    convertJson_dataToExcel(jsonData, workbook, 'Sheet1');

    // Saving workbook to Excel file
    XLSX.writeFile(workbook, outputFilePath);

    console.log(`Conversion successful. Excel file saved at: ${outputFilePath}`);
  } catch (error) {
    console.error(`Error: ${error.message}`);
  }
}


const inputFilePath = 'path/to/your/input.json';
const outputFilePath = 'path/to/your/output.xlsx';

convertJsonFileToExcel(inputFilePath, outputFilePath);


// to run the script:
// command => node script.js