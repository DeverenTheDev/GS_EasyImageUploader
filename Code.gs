function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Image Tools')
    .addItem('Upload images from computer', 'showUploadDialog')
    .addToUi();
}

function showUploadDialog() {
  // Remember where the user's selection is when the dialog opens.
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveRange();
  const target = {
    sheetId: sheet.getSheetId(),
    row: cell.getRow(),       // top-left cell of the selection
    col: cell.getColumn(),
    next: cell.getRow()       // next row to write into
  };
  PropertiesService.getDocumentProperties()
    .setProperty('uploadTarget', JSON.stringify(target));

  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(440)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload images');
}

// Called from the dialog with a small batch of images at a time.
// batch = [{ name: 'file.png', dataUri: 'data:image/png;base64,...' }, ...]
function addImages(batch) {
  const props = PropertiesService.getDocumentProperties();
  const target = JSON.parse(props.getProperty('uploadTarget'));

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter(s => s.getSheetId() === target.sheetId)[0];

  let row = target.next;
  const skipped = [];
  let inserted = 0;

  batch.forEach(item => {
    try {
      // Add rows if we run past the bottom of the sheet
      if (row > sheet.getMaxRows()) {
        sheet.insertRowsAfter(sheet.getMaxRows(), 50);
      }
      const cellImage = SpreadsheetApp.newCellImage()
        .setSourceUrl(item.dataUri)
        .setAltTextTitle(item.name)
        .build();
      sheet.getRange(row, target.col).setValue(cellImage);
      row++;
      inserted++;
    } catch (e) {
      skipped.push(item.name);
    }
  });

  target.next = row;
  props.setProperty('uploadTarget', JSON.stringify(target));

  SpreadsheetApp.flush();
  return { inserted: inserted, skipped: skipped };
}
