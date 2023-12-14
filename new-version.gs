// increased error handlings but need better documentation and more clean code to make it more generic

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Translation')
    .addItem('Start Translating', 'startTranslating')
    .addToUi();
}

function startTranslating() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const chunkSize = 100; // Process 100 rows at a time
  const statusColumn = 'G'; // Status column

  for (let startRow = getStartRow(sheet); startRow <= lastRow; startRow += chunkSize) {
    let endRow = Math.min(startRow + chunkSize - 1, lastRow);
    let sourceRange = sheet.getRange('E' + startRow + ':E' + endRow);
    let formulaRange = sheet.getRange('F' + startRow + ':F' + endRow);
    let statusRange = sheet.getRange(statusColumn + startRow + ':' + statusColumn + endRow);

    let sourceValues = sourceRange.getValues();
    let statusValues = sourceValues.map(() => ['In Progress']);
    statusRange.setValues(statusValues);

    applyTranslationFormulas(formulaRange, sourceValues);

    let retries = 5;
    while (retries > 0) {
      Utilities.sleep(3000); // Wait for 3 seconds
      let translationStatus = checkTranslations(formulaRange);

      if (translationStatus.allTranslated) {
        let translatedValues = formulaRange.getValues();
        formulaRange.setValues(translatedValues); // Replace formulas with values
        statusValues = statusValues.map(() => ['Completed']);
        break;
      } else if (translationStatus.retryRows.length > 0) {
        applyTranslationFormulas(formulaRange, sourceValues, translationStatus.retryRows);
        statusValues = updateStatusValues(statusValues, translationStatus.retryRows, 'Re-calculating');
      } else {
        statusValues = statusValues.map(() => ['Failed']);
        break;
      }

      retries--;
    }

    statusRange.setValues(statusValues);
    if (endRow === lastRow) {
      break; // Stop if we've reached the last row
    }
  }

  Logger.log('Translation process completed.');
}

function applyTranslationFormulas(formulaRange, sourceValues, retryRows = []) {
  let formulasToSet = sourceValues.map((value, index) => {
    if (retryRows.length === 0 || retryRows.includes(index)) {
      return value[0] ? ['=GOOGLETRANSLATE("' + value[0].replace(/"/g, '""') + '", "en", "ja")'] : [""];
    }
    return [formulaRange.getValues()[index][0]];
  });
  formulaRange.setValues(formulasToSet);
  SpreadsheetApp.flush();
}

function updateStatusValues(statusValues, retryRows, status) {
  return statusValues.map((value, index) => retryRows.includes(index) ? [status] : value);
}

function checkTranslations(range) {
  let values = range.getValues();
  let allTranslated = true;
  let retryRows = [];
  for (let i = 0; i < values.length; i++) {
    if (isFormula(values[i][0]) || isTranslationError(values[i][0])) {
      allTranslated = false;
      retryRows.push(i);
    }
  }
  return { allTranslated, retryRows };
}

function isFormula(value) {
  return typeof value === 'string' && value.startsWith('=');
}

function isTranslationError(value) {
  return !value || value === '#VALUE!' || value === '#ERROR!';
}

function getStartRow(sheet) {
  const range = sheet.getRange('F2:F');
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      return i + 2;
    }
  }
  return values.length + 2;
}
