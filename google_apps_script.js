/**
 * Google Apps Script snippet to extract PDF tables from Google Drive
 * and append them directly into your Google Sheet.
 */
function extractPdfToSheet() {
  // 1. Your deployed Railway service endpoint
  const RAILWAY_API_URL = "https://pdf-to-sheet-api-production.up.railway.app/extract-tables";
  
  // 2. Specify the PDF file ID from Google Drive
  const PDF_FILE_ID = "YOUR_GOOGLE_DRIVE_PDF_FILE_ID";
  
  const file = DriveApp.getFileById(PDF_FILE_ID);
  const blob = file.getBlob();
  
  // Build multipart form payload
  const boundary = "------WebKitFormBoundary" + Utilities.getUuid();
  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelimiter = "\r\n--" + boundary + "--";
  
  const header = delimiter +
    'Content-Disposition: form-data; name="file"; filename="' + file.getName() + '"\r\n' +
    'Content-Type: application/pdf\r\n\r\n';
  
  const requestData = Utilities.newBlob(header)
    .getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(closeDelimiter).getBytes());
  
  const options = {
    method: "post",
    contentType: "multipart/form-data; boundary=" + boundary,
    payload: requestData,
    muteHttpExceptions: true
  };
  
  Logger.log("Sending PDF to Railway API...");
  const response = UrlFetchApp.fetch(RAILWAY_API_URL, options);
  const statusCode = response.getResponseCode();
  
  if (statusCode !== 200) {
    Logger.log("Error: " + response.getContentText());
    SpreadsheetApp.getUi().alert("API Error: " + response.getContentText());
    return;
  }
  
  const json = JSON.parse(response.getContentText());
  const rows = json.data;
  
  if (!rows || rows.length === 0) {
    SpreadsheetApp.getUi().alert("No tables found in the PDF.");
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Write rows to the sheet
  const startRow = sheet.getLastRow() + 1;
  const numRows = rows.length;
  const numCols = Math.max(...rows.map(r => r.length));
  
  // Pad rows so all have equal number of columns
  const formattedRows = rows.map(row => {
    while (row.length < numCols) {
      row.push("");
    }
    return row;
  });
  
  sheet.getRange(startRow, 1, numRows, numCols).setValues(formattedRows);
  SpreadsheetApp.getUi().alert(`Successfully extracted ${numRows} rows into the sheet!`);
}
