function updateNewUrls() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const START_ROW = 2;
  const CURRENT_URL_COL = 1; // Column A
  const NEW_URL_COL = 2;     // Column B
  const FIND_COL = 3;        // Column C
  const REPLACE_COL = 4;     // Column D

  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) {
    return;
  }

  const currentUrls = sheet
    .getRange(START_ROW, CURRENT_URL_COL, lastRow - START_ROW + 1, 1)
    .getDisplayValues();

  const replacementRows = sheet
    .getRange(START_ROW, FIND_COL, lastRow - START_ROW + 1, 2)
    .getDisplayValues();

  const replacements = replacementRows
    .filter(row => row[0] !== '' && row[1] !== '')
    .map(row => ({
      find: row[0],
      replaceWith: row[1]
    }));

  const newUrls = currentUrls.map(row => {
    let url = row[0];

    if (url === '') {
      return [''];
    }

    replacements.forEach(rule => {
      url = url.split(rule.find).join(rule.replaceWith);
    });

    return [url];
  });

  sheet
    .getRange(START_ROW, NEW_URL_COL, newUrls.length, 1)
    .setValues(newUrls);
}
