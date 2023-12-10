/**
 * Adds a custom menu to the Google Sheets UI
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Translation Tools')
    .addItem('Translate Text', 'translateText')
    .addToUi();
}

/**
 * Main function to start the translation process
 */
function translateText() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const config = {
    sourceColumn: 'E',
    targetColumn: 'F',
    sourceLanguage: 'en',
    targetLanguage: 'nl',
    chunkSize: 500,
    maxRow: 36774
  };

  processTranslations(sheet, config);
}

/**
 * Processes translations in chunks
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to process.
 * @param {Object} config - Configuration object for translation.
 */
function processTranslations(sheet, config) {
  const lastRow = Math.min(sheet.getLastRow(), config.maxRow);
  let startRow = getStartRow(sheet, config.targetColumn);

  while (startRow <= lastRow) {
    const endRow = Math.min(lastRow, startRow + config.chunkSize - 1);
    applyTranslations(sheet, startRow, endRow, config);
    startRow += config.chunkSize;
  }
}

/**
 * Applies translations to a specified range and replaces formulas with values.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to apply translations.
 * @param {number} startRow - The starting row for the range.
 * @param {number} endRow - The ending row for the range.
 * @param {Object} config - Configuration object for translation.
 */
function applyTranslations(sheet, startRow, endRow, config) {
  const sourceRange = sheet.getRange(`${config.sourceColumn}${startRow}:${config.sourceColumn}${endRow}`);
  const targetRange = sheet.getRange(`${config.targetColumn}${startRow}:${config.targetColumn}${endRow}`);
  const sourceValues = sourceRange.getValues();

  for (let i = 0; i < sourceValues.length; i++) {
    const sourceValue = sourceValues[i][0];
    if (sourceValue) {
      const formulaText = prepareTextForFormula(sourceValue);
      targetRange.getCell(i + 1, 1).setFormula(`=GOOGLETRANSLATE(${formulaText}, "${config.sourceLanguage}", "${config.targetLanguage}")`);
    }
  }

  SpreadsheetApp.flush();
  const values = targetRange.getValues();
  targetRange.setValues(values);
}

/**
 * Finds the next row to start processing based on the target column.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check.
 * @param {string} targetColumn - The column to check for the next row.
 * @return {number} The row number to start processing.
 */
function getStartRow(sheet, targetColumn) {
  const range = sheet.getRange(`${targetColumn}2:${targetColumn}`);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      return i + 2;
    }
  }

  return values.length + 1;
}

/**
 * Prepares text to be safely inserted into a formula.
 * 
 * @param {string} text - The text to prepare.
 * @return {string} The prepared text.
 */
function prepareTextForFormula(text) {
  const escapedText = text.replace(/"/g, '""');
  return `"${escapedText}"`;
}
