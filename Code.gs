/**
 * Creates custom menu in Google Sheets navbar when opened
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏢 Commerciology")
    .addItem("📄 Upload Bank Statement", "showUploadDialog")
    .addToUi();
}

/**
 * Shows the HTML upload modal dialog
 */
function showUploadDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("UploadDialog")
    .setWidth(460)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Commerciology - Bank Statement Extractor");
}

/**
 * Ensures standard tabs (Dashboard, Aggregated, Consolidated) exist.
 * In 'Consolidated', ensures row 1 has key "Tabs" in A1.
 */
function ensureBaseTabsExist(ss) {
  const baseTabs = ["Dashboard", "Aggregated", "Consolidated"];
  baseTabs.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  const consolidatedSheet = ss.getSheetByName("Consolidated");
  // Set "Tabs" key in A1 if not already set
  const a1Value = consolidatedSheet.getRange(1, 1).getValue();
  if (!a1Value || a1Value.toString().trim() === "") {
    consolidatedSheet.getRange(1, 1).setValue("Tabs").setFontWeight("bold");
  }
}

/**
 * Appends the new tab name to row 1 (B1, C1, D1...) under the "Tabs" key
 * in the Consolidated sheet.
 */
function appendTabToConsolidated(ss, tabName) {
  const sheet = ss.getSheetByName("Consolidated");
  if (!sheet) return;

  // Search across Row 1 for the first empty column starting from column 2 (B1)
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const row1Values = sheet.getRange(1, 1, 1, Math.max(lastCol, 20)).getValues()[0];

  let targetCol = -1;
  for (let c = 1; c < row1Values.length; c++) { // index 1 is column B
    if (!row1Values[c] || row1Values[c].toString().trim() === "") {
      targetCol = c + 1;
      break;
    }
  }

  if (targetCol === -1) {
    targetCol = row1Values.length + 1;
  }

  sheet.getRange(1, targetCol).setValue(tabName);
}

/**
 * Generates a unique tab name.
 * If user gives "HDFC" and "HDFC" already exists, it becomes "HDFC 1", "HDFC 2", "HDFC 3", etc.
 */
function getUniqueTabName(ss, baseName) {
  let uniqueName = baseName;
  if (!ss.getSheetByName(uniqueName)) {
    return uniqueName;
  }

  let counter = 1;
  uniqueName = baseName + " " + counter;
  while (ss.getSheetByName(uniqueName)) {
    counter++;
    uniqueName = baseName + " " + counter;
  }
  return uniqueName;
}

/**
 * Server-side function called by the HTML dialog
 * @param {string} bankName Name entered in input bar for the tab
 * @param {string} base64Data Base64 encoded PDF file data
 * @param {string} fileName Original file name
 */
function processBankStatement(bankName, base64Data, fileName) {
  if (!bankName || !bankName.trim()) {
    throw new Error("Please enter a valid Bank Name.");
  }
  if (!base64Data) {
    throw new Error("No PDF file provided.");
  }

  const cleanBankName = bankName.trim();
  const RAILWAY_API_URL = "https://pdftosheet-production.up.railway.app/extract-tables";
  
  // Decode base64 to binary blob
  const decodedBytes = Utilities.base64Decode(base64Data);
  const pdfBlob = Utilities.newBlob(decodedBytes, "application/pdf", fileName || "statement.pdf");

  // Construct multipart/form-data payload
  const boundary = "------WebKitFormBoundary" + Utilities.getUuid();
  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelimiter = "\r\n--" + boundary + "--";

  const header = delimiter +
    'Content-Disposition: form-data; name="file"; filename="' + pdfBlob.getName() + '"\r\n' +
    'Content-Type: application/pdf\r\n\r\n';

  const payload = Utilities.newBlob(header)
    .getBytes()
    .concat(pdfBlob.getBytes())
    .concat(Utilities.newBlob(closeDelimiter).getBytes());

  const options = {
    method: "post",
    contentType: "multipart/form-data; boundary=" + boundary,
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(RAILWAY_API_URL, options);
  const statusCode = response.getResponseCode();

  if (statusCode !== 200) {
    throw new Error("API Error (" + statusCode + "): " + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  const rows = result.data;

  if (!rows || rows.length === 0) {
    throw new Error("No tabular data detected in the uploaded PDF.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Ensure Dashboard, Aggregated, Consolidated exist
  ensureBaseTabsExist(ss);

  // 2. Generate unique tab name (e.g. HDFC, HDFC 1, HDFC 2...)
  const finalTabName = getUniqueTabName(ss, cleanBankName);

  // 3. Create the bank statement tab
  const targetSheet = ss.insertSheet(finalTabName);

  // 4. Record the tab name in Consolidated (Row 1 under 'Tabs' key)
  appendTabToConsolidated(ss, finalTabName);

  // 5. Ensure rectangular 2D array and write data
  const numRows = rows.length;
  const numCols = Math.max(...rows.map(r => r.length));
  const formattedRows = rows.map(r => {
    while (r.length < numCols) {
      r.push("");
    }
    return r;
  });

  targetSheet.getRange(1, 1, numRows, numCols).setValues(formattedRows);
  targetSheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
  targetSheet.activate();

  return {
    sheetName: finalTabName,
    rowCount: numRows
  };
}
