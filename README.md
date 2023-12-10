# GoogleSheets-Translator

Automate translations in Google Sheets using the Google Translate formula. This script is designed to efficiently handle large datasets, processing translations in batches and updating the sheet with minimal delay.

## Features

- **Batch Translation**: Processes translations in manageable chunks to avoid spreadsheet performance issues.
- **Customizable**: Easily configurable for different source/target languages and columns.
- **User-Friendly**: Provides a simple menu in Google Sheets for easy access to the translation functionality.
- **Resumable**: Picks up where it left off, making it suitable for very large datasets.

## Setup

1. **Open Your Google Sheet**: The sheet where you want to perform translations.
2. **Access Apps Script**: Go to `Extensions` > `Apps Script` in the Google Sheets menu.
3. **Create a New Script**: Replace any existing code with the contents of `Translator.gs`.
4. **Save and Close**: After pasting the code, save the project and close the script editor.
5. **Reload the Sheet**: Refresh your Google Sheets tab to see the new 'Translation Tools' menu.

## Usage

- Click on the `Translation Tools` menu in your Google Sheet.
- Select `Translate Text` to start the translation process.
- The script will process translations in batches (default 500 rows at a time).
- You can rerun the script from the menu to process additional batches.

## Configuration

Edit the `translateText` function in `Translator.gs` to change the configuration:

```javascript
const config = {
  sourceColumn: 'E',        // Column containing the original text
  targetColumn: 'F',        // Column where the translated text will be placed
  sourceLanguage: 'en',     // Source language (e.g., 'en' for English)
  targetLanguage: 'pt',     // Target language (e.g., 'pt' for Portuguese)
  chunkSize: 500,           // Number of rows processed in each batch
  maxRow: 36774             // Maximum row to process
};
```
