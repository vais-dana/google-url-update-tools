/**
 * Google Slides script.
 *
 * Before using:
 * 1. Paste this into Extensions > Apps Script inside your Google Slides deck.
 * 2. Replace PASTE_YOUR_SHEET_ID_HERE with the ID of your URL Mapping Google Sheet.
 * 3. Save the script.
 * 4. Reload the Slides deck.
 * 5. Use URL Tools > Update slide links.
 */
const URL_MAPPING_SPREADSHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';
const URL_MAPPING_SHEET_NAME = 'Sheet1';

function onOpen() {
  SlidesApp.getUi()
    .createMenu('URL Tools')
    .addItem('Update slide links', 'updateSlideLinks')
    .addToUi();
}

function updateSlideLinks() {
  const mapping = getUrlMappingFromSheet_();
  const presentation = SlidesApp.getActivePresentation();

  let replacedCount = 0;
  let checkedCount = 0;

  presentation.getSlides().forEach(slide => {
    slide.getPageElements().forEach(element => {
      const result = updateLinksInPageElement_(element, mapping);
      replacedCount += result.replaced;
      checkedCount += result.checked;
    });
  });

  SlidesApp.getUi().alert(
    `Done.\nChecked ${checkedCount} link(s).\nUpdated ${replacedCount} link(s).`
  );
}

function getUrlMappingFromSheet_() {
  const spreadsheet = SpreadsheetApp.openById(URL_MAPPING_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(URL_MAPPING_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet not found: ${URL_MAPPING_SHEET_NAME}`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  const mapping = {};

  values.forEach(row => {
    const currentUrl = String(row[0]).trim();
    const newUrl = String(row[1]).trim();

    if (currentUrl !== '' && newUrl !== '') {
      mapping[currentUrl] = newUrl;
    }
  });

  return mapping;
}

function updateLinksInPageElement_(element, mapping) {
  let replaced = 0;
  let checked = 0;
  const type = element.getPageElementType();

  if (type === SlidesApp.PageElementType.SHAPE) {
    const shape = element.asShape();

    const objectResult = updateObjectLink_(shape, mapping);
    replaced += objectResult.replaced;
    checked += objectResult.checked;

    try {
      const textResult = updateLinksInTextRange_(shape.getText(), mapping);
      replaced += textResult.replaced;
      checked += textResult.checked;
    } catch (err) {}
  }

  if (type === SlidesApp.PageElementType.IMAGE) {
    const result = updateObjectLink_(element.asImage(), mapping);
    replaced += result.replaced;
    checked += result.checked;
  }

  if (type === SlidesApp.PageElementType.LINE) {
    const result = updateObjectLink_(element.asLine(), mapping);
    replaced += result.replaced;
    checked += result.checked;
  }

  if (type === SlidesApp.PageElementType.VIDEO) {
    const result = updateObjectLink_(element.asVideo(), mapping);
    replaced += result.replaced;
    checked += result.checked;
  }

  if (type === SlidesApp.PageElementType.TABLE) {
    const table = element.asTable();

    for (let row = 0; row < table.getNumRows(); row++) {
      for (let col = 0; col < table.getNumColumns(); col++) {
        try {
          const textResult = updateLinksInTextRange_(table.getCell(row, col).getText(), mapping);
          replaced += textResult.replaced;
          checked += textResult.checked;
        } catch (err) {}
      }
    }
  }

  if (type === SlidesApp.PageElementType.GROUP) {
    element.asGroup().getChildren().forEach(child => {
      const childResult = updateLinksInPageElement_(child, mapping);
      replaced += childResult.replaced;
      checked += childResult.checked;
    });
  }

  return { replaced, checked };
}

function updateLinksInTextRange_(textRange, mapping) {
  let replaced = 0;
  let checked = 0;

  textRange.getRuns().forEach(run => {
    const link = run.getTextStyle().getLink();
    if (!link) return;

    const currentUrl = link.getUrl();
    if (!currentUrl) return;

    checked++;

    if (mapping[currentUrl]) {
      run.getTextStyle().setLinkUrl(mapping[currentUrl]);
      replaced++;
    }
  });

  return { replaced, checked };
}

function updateObjectLink_(object, mapping) {
  if (typeof object.getLink !== 'function') {
    return { replaced: 0, checked: 0 };
  }

  if (typeof object.setLinkUrl !== 'function') {
    return { replaced: 0, checked: 0 };
  }

  const link = object.getLink();
  if (!link) {
    return { replaced: 0, checked: 0 };
  }

  const currentUrl = link.getUrl();
  if (!currentUrl) {
    return { replaced: 0, checked: 0 };
  }

  if (mapping[currentUrl]) {
    object.setLinkUrl(mapping[currentUrl]);
    return { replaced: 1, checked: 1 };
  }

  return { replaced: 0, checked: 1 };
}

