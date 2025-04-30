# mergeSheetsToBaseFile

Merges rows from multiple sheets into a single sheet of a base Excel file while preserving formatting and structure.

> ⚠️ **Experimental API**
> Interface may change in future releases.

---

## 📦 Functions

### `mergeSheetsToBaseFile`

Asynchronously merges rows from specified sheets of one or more Excel files into a sheet of a base Excel file.

```ts
mergeSheetsToBaseFile(options: MergeSheetsOptions): Promise<Buffer>
```

- Input: `options: MergeSheetsOptions` — configuration for merging
- Output: `Promise<Buffer>` — the resulting `.xlsx` file as a buffer
- Throws if:
  - Required files or sheets are missing
  - Sheet indices are invalid
  - XML structure is malformed
  - Merged cells conflict

---

### `mergeSheetsToBaseFileSync`

Synchronous version of `mergeSheetsToBaseFile`.

```ts
mergeSheetsToBaseFileSync(options: MergeSheetsOptions): Buffer
```

- Input: `options: MergeSheetsOptions`
- Output: `Buffer` — resulting Excel file
- Same error conditions as async version

---

## ⚙️ Parameters

### `MergeSheetsOptions`

```ts
interface MergeSheetsOptions {
  baseFile: Buffer;
  baseSheetIndex?: number; // default: 1
  additions: Array<{
    file: Buffer;
    sheetIndexes: number[];
    isBaseFile?: boolean;
  }>;
  gap?: number; // default: 0
  sheetNamesToRemove?: string[];
  sheetsToRemove?: number[];
}
```

- `baseFile` — buffer of the base Excel file
- `baseSheetIndex` — 1-based index of the sheet to merge into
- `additions` — list of files and sheets to merge
- `gap` — number of empty rows to insert between sections
- `sheetNamesToRemove` — optional sheet names to remove
- `sheetsToRemove` — optional sheet indices (1-based) to remove

---

## 🧩 Features

### ✅ Merged Cell Support

- Preserves existing merged cells
- Adjusts merge ranges on row shifts
- Handles overlapping cells gracefully

### 📐 Row Shifting

- Automatically shifts incoming rows
- Updates row numbers and references
- Supports configurable row gaps

### 🗂️ Sheet Management

- Removes sheets by name or index
- Cleans up `workbook.xml`, `workbook.xml.rels`, `[Content_Types].xml`
- Preserves valid sheet references

---

## ❌ Errors

Throws when:

- `baseFile` or required `sheetIndexes` are missing
- Sheets in `additions` don’t exist
- Sheet merges produce invalid or overlapping ranges
- Sheet names or indices for removal are invalid
- Input files are corrupt or malformed

---

## 💡 Best Practices

1. Specify `baseSheetIndex` when using multi-sheet base files
2. Use `gap` for readability in merged output
3. Clean unused sheets to minimize output size
4. Validate all inputs prior to merge
5. Catch and log merge errors during integration

---

## 🧪 Usage Examples

### Async Merge

```ts
import { mergeSheetsToBaseFile } from '@js-ak/excel-toolbox';

const result = await mergeSheetsToBaseFile({
  baseFile: baseFileBuffer,
  baseSheetIndex: 1,
  additions: [
    {
      file: sourceFileBuffer,
      sheetIndexes: [1, 2],
    }
  ],
  gap: 1,
  sheetsToRemove: [3, 4]
});

await fs.writeFile('merged.xlsx', result);
```

---

### Sync Merge

```ts
import { mergeSheetsToBaseFileSync } from '@js-ak/excel-toolbox';

const result = mergeSheetsToBaseFileSync({
  baseFile: baseFileBuffer,
  baseSheetIndex: 1,
  additions: [
    {
      file: sourceFileBuffer,
      sheetIndexes: [1, 2],
    }
  ],
  gap: 1,
});

fs.writeFileSync('merged.xlsx', result);
```

---

## 🧼 Cleanup & Validation

- Sheet removals affect XML metadata and relationships
- Output file is fully valid `.xlsx`
- Sheet relationships are adjusted during merge

---
