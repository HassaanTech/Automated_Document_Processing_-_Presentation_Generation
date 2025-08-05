function extractContentToSpreadsheet() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var fullText = body.getText(); // Entire document text
  var spreadsheetId = "1N2I2_fLbeKsQLAVBPxEGelr0Bdr59OuEMCvUAcEy4K4";
  var sheetName = "Sheet1";

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Error: Sheet not found. Check the sheet name.");
    return;
  }

  // Clear all rows from row 3 onward, leaving rows 1 and 2 intact.
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearContent();
  }

  // Updated mapping: only the keys you want.
  var sectionMap = {
    "SLIDE1_TITLE": "A3",
    "SLIDE2_HEADER": "C3",
    "SLIDE2_CONTENT": "D3",
    "SLIDE3_HEADER": "F3",
    "SLIDE3_CONTENT": "G3",
    "SLIDE4_TITLE": "I3",
    "SLIDE5_HEADER1": "K3",
    "SLIDE5_HEADER2": "L3",
    "SLIDE5_CONTENT": "M3",
    "SLIDE6_HEADER1": "O3",
    "SLIDE6_HEADER2": "P3",
    "SLIDE6_CONTENT1": "R3",
    "SLIDE6_CONTENT2": "T3",
    "SLIDE6_CONTENT3": "V3",
    "SLIDE7_HEADER1": "X3",
    "SLIDE7_HEADER2": "Y3",
    "SLIDE7_CONTENT1": "AA3",
    "SLIDE7_CONTENT2": "AC3"
  };

  // Prepare counters for each tag.
  var rowCounters = {};
  for (var tag in sectionMap) {
    var cellAddress = sectionMap[tag];
    var match = cellAddress.match(/([A-Z]+)(\d+)/);
    if (match) {
      // Save the starting row number for each tag.
      rowCounters[tag] = parseInt(match[2], 10);
    } else {
      rowCounters[tag] = 3; // Default to row 3 if parsing fails.
    }
  }

  // Modify regex to capture each occurrence.
  var regex = /\{(SLIDE[0-9A-Z_]+)\}\s*\{\{\s*([\s\S]*?)\s*\}\}/g;
  var match;
  // Instead of a simple map, store arrays of values per tag.
  var contentMap = {};
  while ((match = regex.exec(fullText)) !== null) {
    var tag = match[1].trim();
    var content = match[2].trim();
    if (sectionMap[tag]) {
      if (!contentMap[tag]) {
        contentMap[tag] = [];
      }
      contentMap[tag].push(content);
    }
  }
  
  Logger.log("Extracted content: " + JSON.stringify(contentMap));

  // Write the captured content into the designated cells,
  // appending to new rows if more than one value exists for a tag.
  for (var tag in contentMap) {
    var values = contentMap[tag];
    // Extract the column letter from the cell address.
    var cellAddress = sectionMap[tag];
    var colMatch = cellAddress.match(/([A-Z]+)/);
    if (!colMatch) continue;
    var colLetter = colMatch[1];
    
    values.forEach(function(value) {
      var rowNumber = rowCounters[tag];
      var targetCell = colLetter + rowNumber;
      Logger.log("Writing " + tag + " to " + targetCell + ": " + value);
      sheet.getRange(targetCell).setValue(value);
      // Increment the row counter for this tag.
      rowCounters[tag]++;
    });
  }

  Logger.log("Content successfully transferred to the spreadsheet.");
}
