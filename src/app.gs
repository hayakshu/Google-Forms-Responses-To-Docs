/**
 * Get's all of the rows that are visible on the active sheet.
 *
 * @returns {visibleData} rows that are visible on active sheet
 */
function getVisibleData() {
  try {
    // get the active sheet
    var sheet = SpreadsheetApp.getActiveSheet();
    Logger.log("Active sheet name: " + sheet.getName());

    // data in the active sheet
    var data = sheet.getDataRange().getValues();
    Logger.log("Data in Active Sheet: " + data);

    // filter to get data that is visible on UI
    var visibleData = data.filter(function(_, i) {
      return (
        !sheet.isRowHiddenByFilter(i + 1) && !sheet.isRowHiddenByUser(i + 1)
      );
    });
    Logger.log("Visible data in Active Sheet: " + visibleData);

    return visibleData;
  } catch (err) {
    throw new Exception(err);
  }
}

/**
 * Determines if column is hidden by user
 *
 * @param {colNum}
 * @returns {result}
 */
function isColumnHidden(colNum) {
  var sheet = SpreadsheetApp.getActiveSheet();
  return sheet.isColumnHiddenByUser(colNum);
}

/**
 * Creates Google Doc that contains converts Google Forms responses
 * that are visible on the current active sheet.
 *
 * @param {fileName} name of saved file
 */
function createDoc(fileName) {
  try {
    var data = getVisibleData();

    // Check that data contains more than column definition
    if (data.length > 1) {
      // create document on users drive
      var doc = DocumentApp.create(fileName);

      for (var i = 1; i < data.length; i++) {
        for (var j = 0; j < data[0].length; j++) {
          // check from index 1 due to google reqs
          if (isColumnHidden(j + 1)) {
            continue;
          }

          // check data exists
          if (data[i][j] !== "") {
            doc.appendParagraph(data[0][j])
               .setHeading(DocumentApp.ParagraphHeading.HEADING5);
            doc.appendParagraph(data[i][j]);
          }
        }
        if (i < data.length - 1) {
          doc.appendPageBreak();
        }
      }
    }
  } catch (err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

/**
 * Main Function that prompts user for FileName.
 */
function main() {
  var fileName = Browser.inputBox("Save responses as (ex. myDocFile):");

  if (fileName.length !== 0) {
    createDoc(fileName);
    return;
  }

  Browser.msgBox("Error: Invalid file name.");
}

/**
 * Optional:
 *
 * Creates a custom menu for Google Spread Sheets as Add-On.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Forms-Tools")
    .addItem("Response to Docs", "main")
    .addToUi();
}
