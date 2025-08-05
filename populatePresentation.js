function populatePresentation() {
  // --- Configuration: update these IDs and sheet name ---
  var spreadsheetId = "1N2I2_fLbeKsQLAVBPxEGelr0Bdr59OuEMCvUAcEy4K4";  // Replace with your spreadsheet ID
  var sheetName = "Sheet1";                   // Replace with your sheet name if different
  var templatePresentationId = "1zI5aCcWBZtz_4ZZQo2zJQ3YsAiofA9eQr4uharu0uAw";
  var targetFolderId = "1rp7w_OaaAoCK1XRronbz_Hn__c1zeEg9";
  
  // --- Open the Spreadsheet and get the specific sheet ---
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  
  // --- Read all data rows starting from row 3 ---
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 2; // starting at row 3
  // We assume the highest used column is AC (i.e. 29 columns)
  var numCols = 29;
  var values = sheet.getRange(3, 1, numRows, numCols).getValues();
  var richValues = sheet.getRange(3, 1, numRows, numCols).getRichTextValues();
  
  // --- Use the value in A3 for naming the presentation ---
  var slide1TitleFirstRow = values[0][0]; // cell A3
  var presentationName = slide1TitleFirstRow 
      ? slide1TitleFirstRow.toString().replace(/\[BOLD\]/g, "").replace(/\[\/BOLD\]/g, "")
      : "Presentation";
  
  // --- Copy the template presentation to create a new presentation ---
  var templateFile = DriveApp.getFileById(templatePresentationId);
  var newPresentationFile = templateFile.makeCopy(presentationName);
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  targetFolder.addFile(newPresentationFile);
  DriveApp.getRootFolder().removeFile(newPresentationFile);
  
  var presentation = SlidesApp.openById(newPresentationFile.getId());
  // Get a static reference to the template slides.
  // We assume slides 1 to 7 (indexes 0 to 6) will be processed; slide 8 (index 7) is left unchanged.
  var originalSlides = presentation.getSlides();
  
  /* --- Mapping object for slides 1 to 7 ---
     For each slide (using its template index), we list objects defining:
       - placeholder: the text to be replaced in the slide
       - col: the column index (0-based) from the sheet's data array
       - isRich: whether to use the rich text value (true) or the plain text value (false)
       
     The column mapping is as follows:
     
     Slide 1 (template index 0):
       • A: slide title (col 0)
       
     Slide 2 (template index 1):
       • C: header (col 2)
       • D: content (rich text) (col 3)
       
     Slide 3 (template index 2):
       • F: header (col 5)
       • G: content (rich text) (col 6)
       
     Slide 4 (template index 3):
       • I: title (col 8)
       
     Slide 5 (template index 4):
       • K: header1 (col 10)
       • L: header2 (col 11)
       • M: content (rich text) (col 12)
       
     Slide 6 (template index 5):
       • O: header1 (col 14)
       • P: header2 (col 15)
       • R: content1 (rich text) (col 17)
       • T: content2 (rich text) (col 19)
       • V: content3 (rich text) (col 21)
       (Removed header3, header4, header5)
       
     Slide 7 (template index 6):
       • X: header1 (col 23)
       • Y: header2 (col 24)
       • AA: content1 (rich text) (col 26)
       • AC: content2 (rich text) (col 28)
       (Removed header3 and header4)
  */
  var mappings = {
    0: [ // Slide 1
      { placeholder: "{{SLIDE1_TITLE}}", col: 0, isRich: false }
    ],
    1: [ // Slide 2
      { placeholder: "{{SLIDE2_HEADER}}", col: 2, isRich: false },
      { placeholder: "{{SLIDE2_CONTENT}}", col: 3, isRich: true }
    ],
    2: [ // Slide 3
      { placeholder: "{{SLIDE3_HEADER}}", col: 5, isRich: false },
      { placeholder: "{{SLIDE3_CONTENT}}", col: 6, isRich: true }
    ],
    3: [ // Slide 4
      { placeholder: "{{SLIDE4_TITLE}}", col: 8, isRich: false }
    ],
    4: [ // Slide 5
      { placeholder: "{{SLIDE5_HEADER1}}", col: 10, isRich: false },
      { placeholder: "{{SLIDE5_HEADER2}}", col: 11, isRich: false },
      { placeholder: "{{SLIDE5_CONTENT}}", col: 12, isRich: true }
    ],
    5: [ // Slide 6 (removed extra headers)
      { placeholder: "{{SLIDE6_HEADER1}}", col: 14, isRich: false },
      { placeholder: "{{SLIDE6_HEADER2}}", col: 15, isRich: false },
      { placeholder: "{{SLIDE6_CONTENT1}}", col: 17, isRich: true },
      { placeholder: "{{SLIDE6_CONTENT2}}", col: 19, isRich: true },
      { placeholder: "{{SLIDE6_CONTENT3}}", col: 21, isRich: true }
    ],
    6: [ // Slide 7 (removed extra headers)
      { placeholder: "{{SLIDE7_HEADER1}}", col: 23, isRich: false },
      { placeholder: "{{SLIDE7_HEADER2}}", col: 24, isRich: false },
      { placeholder: "{{SLIDE7_CONTENT1}}", col: 26, isRich: true },
      { placeholder: "{{SLIDE7_CONTENT2}}", col: 28, isRich: true }
    ]
    // Slide 8 (template index 7) remains unchanged.
  };
  
  // --- For each slide type (template indexes 0 to 6) process all data rows.
  // Loop through rows in reverse order so that the final slide order matches the sheet order.
  for (var slideIdx = 0; slideIdx <= 6; slideIdx++) {
    var templateSlide = originalSlides[slideIdx];
    var mapping = mappings[slideIdx];
    
    for (var r = numRows - 1; r >= 0; r--) {
      // Check if at least one field for this slide is non-empty in this row.
      var hasData = false;
      for (var m = 0; m < mapping.length; m++) {
        var col = mapping[m].col;
        var cellValue = values[r][col];
        var cellRich = richValues[r][col];
        if (!isEmptyValue(cellValue) || !isEmptyValue(cellRich)) {
          hasData = true;
          break;
        }
      }
      if (hasData) {
        // Duplicate the template slide first, then process the duplicate.
        var currentSlide = templateSlide.duplicate();
        // Replace each placeholder with data from this row.
        for (var m = 0; m < mapping.length; m++) {
          var placeholder = mapping[m].placeholder;
          var col = mapping[m].col;
          if (mapping[m].isRich) {
            var richValue = richValues[r][col];
            if (!isEmptyValue(richValue)) {
              replacePlaceholderWithRichText(currentSlide, placeholder, richValue);
            } else {
              currentSlide.replaceAllText(placeholder, "");
            }
          } else {
            var cellValue = values[r][col];
            if (!isEmptyValue(cellValue)) {
              currentSlide.replaceAllText(placeholder, processHeader(cellValue.toString()));
            } else {
              currentSlide.replaceAllText(placeholder, "");
            }
          }
        }
      }
    }
    // Remove the original template slide after processing all rows.
    templateSlide.remove();
  }
  
  // Slide 8 (template index 7) remains unchanged.
  presentation.saveAndClose();
}

/**
 * Checks whether a plain text value or a RichTextValue is empty.
 */
function isEmptyValue(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === "string") {
    return value.trim() === "";
  }
  if (value.getText) {
    return value.getText().trim() === "";
  }
  return false;
}

/**
 * Processes header text by removing [BOLD][/BOLD] and [BULLET] markers.
 * The result is used in a simple replaceAllText call so that the slide's native styling is preserved.
 */
function processHeader(text) {
  return text.replace(/\[BOLD\]/g, "").replace(/\[\/BOLD\]/g, "").replace(/\[BULLET\]/g, "");
}

/**
 * Searches for a shape in the slide that contains the placeholder text,
 * then replaces its content with the provided rich text value (preserving formatting)
 * after processing the [BULLET] and [BOLD] markers.
 */
function replacePlaceholderWithRichText(slide, placeholder, richTextValue) {
  var shapes = slide.getShapes();
  for (var i = 0; i < shapes.length; i++) {
    if (shapes[i].getText) {
      var textRange = shapes[i].getText();
      if (textRange.asString().indexOf(placeholder) !== -1) {
        setRichTextInShape(shapes[i], richTextValue);
        break; // Replace only the first occurrence of the placeholder
      }
    }
  }
}

/**
 * Clears the text in a shape and inserts rich text content by processing
 * the [BULLET] and [BOLD][/BOLD] markers:
 * - Replaces all [BULLET] markers with a newline and bullet ("\n• ").
 * - Sets text between [BOLD] and [/BOLD] to bold and removes the markers.
 * Accepts a RichTextValue from the spreadsheet.
 */
function setRichTextInShape(shape, richTextValue) {
  var originalText = richTextValue.getText();
  // Replace bullet markers with newline+bullet and remove a leading newline if present.
  var processedText = originalText.replace(/\[BULLET\]/g, "\n• ");
  if (processedText.indexOf("\n") === 0) {
    processedText = processedText.substring(1);
  }
  
  // Parse processedText for bold markers.
  var segments = [];
  var boldRegex = /\[BOLD\](.*?)\[\/BOLD\]/g;
  var lastIndex = 0;
  var match;
  while ((match = boldRegex.exec(processedText)) !== null) {
    if (match.index > lastIndex) {
      segments.push({
        text: processedText.substring(lastIndex, match.index),
        bold: false
      });
    }
    segments.push({
      text: match[1],
      bold: true
    });
    lastIndex = boldRegex.lastIndex;
  }
  if (lastIndex < processedText.length) {
    segments.push({
      text: processedText.substring(lastIndex),
      bold: false
    });
  }
  
  var textRange = shape.getText();
  textRange.setText(""); // Clear existing text
  
  // Append each segment with the appropriate bold styling.
  for (var i = 0; i < segments.length; i++) {
    var appendedRange = textRange.appendText(segments[i].text);
    appendedRange.getTextStyle().setBold(segments[i].bold);
  }
}
