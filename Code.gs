/**
 * Creates custom menu in Google Sheets navbar when opened
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Commerciology")
    .addItem("Upload Bank Statement", "showUploadDialog")
    .addToUi();
}

/**
 * Shows the HTML upload modal dialog
 */
function showUploadDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("UploadDialog")
    .setWidth(450)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Commerciology - Bank Statement Extractor");
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

  // Handle Sheet Tab creation
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(cleanBankName);

  // If a tab with this bank name already exists, generate a unique name (e.g. HDFC (1))
  if (targetSheet) {
    let counter = 1;
    let uniqueName = cleanBankName + " (" + counter + ")";
    while (ss.getSheetByName(uniqueName)) {
      counter++;
      uniqueName = cleanBankName + " (" + counter + ")";
    }
    targetSheet = ss.insertSheet(uniqueName);
  } else {
    targetSheet = ss.insertSheet(cleanBankName);
  }

  // Ensure rectangular 2D array
  const numRows = rows.length;
  const numCols = Math.max(...rows.map(r => r.length));
  const formattedRows = rows.map(r => {
    while (r.length < numCols) {
      r.push("");
    }
    return r;
  });

  // Write data to sheet
  targetSheet.getRange(1, 1, numRows, numCols).setValues(formattedRows);

  // Optional styling: Bold header row & auto-resize columns
  targetSheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
  targetSheet.activate();

  return {
    sheetName: targetSheet.getName(),
    rowCount: numRows
  };
}
