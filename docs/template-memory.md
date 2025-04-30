# TemplateMemory

The `TemplateMemory` class is designed for working with Excel (`.xlsx`) templates entirely in memory. It enables modifying templates, inserting rows, substituting placeholders, and saving as a Buffer or to a writable stream — all without touching the filesystem.

> ⚠️ **Experimental API**
> Interface is subject to change in future versions.

---

## 🔧 Constructor

```ts
new TemplateMemory(files: Record<string, Buffer>)
```

- Input: `files` — a map of file paths to their contents as `Buffer`s, representing the `.xlsx` file structure.
- Output: `TemplateMemory` instance
- Preconditions: None
- Postconditions: Instance is ready for use with provided files

> Prefer using the static method `TemplateMemory.from()` to create instances.

---

## 📄 Properties

- `files: Record<string, Buffer>` — in-memory map of file contents.
- `destroyed: boolean` — indicates whether the instance has been destroyed (read-only).

---

## 📚 Methods

### `copySheet`

Creates a copy of an existing worksheet with a new name.

- Input:
  - `sourceName: string` - name of existing sheet
  - `newName: string` - name for new sheet
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - `sourceName` exists
  - `newName` does not exist
- Postconditions:
  - New sheet created with content from source
  - Sheet relationships updated
- Throws if:
  - `sourceName` does not exist.
  - `newName` already exists.

---

### `substitute`

Replaces placeholders of the form `${key}` with values from the `replacements` object. For arrays, use placeholders with key `${table:key}`.

- Input:
  - `sheetName: string` - name of worksheet
  - `replacements: Record<string, unknown>` - key-value map for substitution
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - Sheet exists
- Postconditions:
  - Placeholders replaced with values
  - Shared strings updated if needed

---

### `insertRows`

Inserts rows into a specified worksheet.

- Input:
  - `sheetName: string` - name of worksheet
  - `startRowNumber?: number` - starting row index (default: append to the end).
  - `rows: unknown[][]` - array of arrays, each representing a row of values.
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - Sheet exists
  - Row number valid if specified
  - Cells within bounds
- Postconditions:
  - Rows inserted at specified position
  - Sheet data updated
- Throws if:
  - The sheet does not exist.
  - The row number is invalid.
  - A cell is out of bounds.

---

### `insertRowsStream`

Streams and inserts rows into a worksheet, useful for handling large datasets.

- Input:
  - `sheetName: string` - name of worksheet
  - `startRowNumber?: number` - starting row index (default: append to the end).
  - `rows: AsyncIterable<unknown[]>` - an async iterable where each item is an array of cell values.
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - Sheet exists
  - Row number valid if specified
  - Cells within bounds
- Postconditions:
  - Rows streamed and inserted
  - Sheet data updated
- Same error conditions as `insertRows`.

---

### `save`

Generates a new Excel file and returns it as a `Buffer`.

- Input: None
- Output: `Promise<Buffer>` — the full `.xlsx` file contents in memory.
- Preconditions:
  - Instance not destroyed
- Postconditions:
  - Instance marked as destroyed
  - All buffers cleared
  - ZIP archive created
- Throws if:
  - The instance has been destroyed.
  - There was a failure while rebuilding the ZIP archive.

---

### `set`

Replaces the content of a specific file in the template.

- Input:
  - `key: string`  — the Excel path of the file (e.g., `xl/worksheets/sheet1.xml`).
  - `content: Buffer` - new file content as a Buffer.
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - File exists
- Postconditions:
  - File content updated
- Throws if:
  - The instance has been destroyed.
  - The file does not exist in the template.

---

### `mergeSheets`

Merges multiple worksheets into a single base worksheet.

- Input:
- `additions` — defines the sheets to merge:
  - `additions.sheetIndexes?: number[]` — array of 1-based sheet indexes to merge.
  - `additions.sheetNames?: string[]` — array of sheet names to merge.
  - `baseSheetIndex?: number` — 1-based index of the base sheet to merge into (optional, default is 1).
  - `baseSheetName?: string` — name of the base sheet to merge into (optional).
  - `gap?: number` - number of empty rows to insert between merged sections (default: `0`).
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - Valid sheet names/indexes
  - Either baseSheetIndex or baseSheetName defined
- Postconditions:
  - Sheets merged into base sheet
  - Row numbers adjusted
  - Merge cells updated
- Throws if:
  - The instance is destroyed.
  - Invalid sheet names or indexes are provided.
  - Both `baseSheetIndex` and `baseSheetName` are undefined.

---

### `removeSheets`

Removes worksheets from the workbook.

- Input:
  - `sheetNames?: string[]` - names of sheets to remove
  - `sheetIndexes?: number[]` - 1-based indexes of sheets to remove
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - Sheets exist
  - Either sheetNames or sheetIndexes provided
- Postconditions:
  - Sheets removed
  - Workbook relationships updated
  - Content types updated
- Throws if:
  - The instance is destroyed.
  - Sheet names or indexes do not exist.
  - Neither `sheetNames` nor `sheetIndexes` are provided.

---

### `from`

Creates a new `TemplateMemory` instance from a source Excel file.

- Input:
  - `options: object` — configuration object
    - `source: string | Buffer` — path to source Excel file or its Buffer content
- Output: `Promise<TemplateMemory>`
- Preconditions:
  - Valid source file/Buffer
- Postconditions:
  - Files loaded into memory
  - Instance ready for use
- Throws if:
  - Source file invalid or missing
  - File parsing fails

---

## 💡 Usage Examples

### Create from File and Modify

```ts
import { TemplateMemory } from '@js-ak/excel-toolbox';

const template = await TemplateMemory.from({
  source: fs.readFileSync("template.xlsx")
});

await template.copySheet("Sheet1", "Sheet2");
await template.substitute("Sheet1", { name: "John Doe" });

const modifiedExcel = await template.save();
fs.writeFileSync("output.xlsx", modifiedExcel);
```

### Insert Rows from a Stream

```ts
import { TemplateMemory } from '@js-ak/excel-toolbox';

async function* generateData() {
  for (let i = 0; i < 1000; i++) {
    yield [i, `Name ${i}`, new Date()];
  }
}

const template = await TemplateMemory.from({
  source: fs.readFileSync("template.xlsx")
});

await template.insertRowsStream({
  sheetName: "Data",
  rows: generateData()
});

fs.writeFileSync("large-output.xlsx", await template.save());
```

### Replace Internal File Content

```ts
import { TemplateMemory } from '@js-ak/excel-toolbox';

const template = await TemplateMemory.from({
  source: fs.readFileSync("template.xlsx")
});

// Manually modify a specific XML file in the template
await template.set(
  "xl/sharedStrings.xml",
  Buffer.from("<sst><si><t>New Text</t></si></sst>")
);
```

---

## 🛑 Internal Checks

Methods perform validation:

- Ensures the instance hasn't been destroyed.
- Prevents concurrent modifications.
