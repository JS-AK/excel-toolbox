# Merging Sheets into Base File

This document explains how to merge sheets from multiple Excel files into a base file using the Excel Toolbox library.

## Overview

The merge functionality allows you to:

- Combine rows from multiple sheets into a single base sheet
- Maintain cell formatting and merged cells
- Insert gaps between merged sections
- Remove unwanted sheets from the output file

## Basic Usage

### Asynchronous Merge

```typescript
import { mergeSheetsToBaseFile } from '@js-ak/excel-toolbox';

const result = await mergeSheetsToBaseFile({
  baseFile: baseFileBuffer,
  baseSheetIndex: 1, // Optional, defaults to 1
  additions: [
    {
      file: sourceFileBuffer,
      sheetIndexes: [1, 2], // Sheets to merge from source file
      isBaseFile: false // Optional, set to true if source is the base file
    }
  ],
  gap: 1, // Optional, number of empty rows between sections
  sheetNamesToRemove: [], // Optional, names of sheets to remove
  sheetsToRemove: [] // Optional, indices of sheets to remove
});
```

### Synchronous Merge

```typescript
import { mergeSheetsToBaseFileSync } from '@js-ak/excel-toolbox';

const result = mergeSheetsToBaseFileSync({
  baseFile: baseFileBuffer,
  baseSheetIndex: 1,
  additions: [
    {
      file: sourceFileBuffer,
      sheetIndexes: [1, 2]
    }
  ],
  gap: 1
});
```

## Parameters

### Base File Configuration

- `baseFile`: Buffer containing the base Excel file
- `baseSheetIndex`: 1-based index of the sheet to merge into (default: 1)

### Additions Configuration

- `additions`: Array of objects specifying files and sheets to merge
  - `file`: Buffer containing the source Excel file
  - `sheetIndexes`: Array of 1-based sheet indices to merge
  - `isBaseFile`: Optional boolean indicating if source is the base file

### Optional Parameters

- `gap`: Number of empty rows to insert between merged sections (default: 0)
- `sheetNamesToRemove`: Array of sheet names to remove from output
- `sheetsToRemove`: Array of 1-based sheet indices to remove from output

## Features

### Merged Cells Support

The merge process automatically handles merged cells by:

- Preserving existing merged cells in the base sheet
- Adjusting merged cell references when rows are shifted
- Maintaining merge cell count and references

### Row Shifting

- Rows from source sheets are automatically shifted to maintain proper row numbers
- Cell references are updated to reflect new row positions
- Gaps between sections are properly maintained

### Sheet Management

- Can remove specific sheets by name or index
- Updates workbook.xml, workbook.xml.rels, and [Content_Types].xml accordingly
- Maintains proper sheet relationships and references

## Error Handling

The merge process will throw errors for:

- Missing base file or sheet
- Invalid sheet indices
- Missing source files or sheets
- Invalid XML structure
- Duplicate merged cells
- Overlapping merge ranges

## Best Practices

1. Always specify a valid `baseSheetIndex` if your base file has multiple sheets
2. Use `gap` parameter to improve readability of merged data
3. Remove unnecessary sheets to reduce file size
4. Validate input files before merging
5. Handle errors appropriately in your application

## Example

```typescript
// Merge two sheets from source file into base file
const result = await mergeSheetsToBaseFile({
  baseFile: baseFileBuffer,
  additions: [
    {
      file: sourceFileBuffer,
      sheetIndexes: [1, 2]
    }
  ],
  gap: 1,
  sheetsToRemove: [3, 4] // Remove sheets 3 and 4 from output
});

// Save the result
await fs.writeFile('merged.xlsx', result);
```
