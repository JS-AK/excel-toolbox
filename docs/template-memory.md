# TemplateMemory

The `TemplateMemory` class is designed for working with Excel (`.xlsx`) templates entirely in memory. It enables modifying templates, inserting rows, substituting placeholders, and saving as a Buffer or to a writable stream — all without touching the filesystem.

> ⚠️ **Experimental API**
> Interface is subject to change in future versions.

---

## 🔧 Constructor

```ts
new TemplateMemory(files: Record<string, Buffer>)
```

- `files` — a map of file paths to their contents as `Buffer`s, representing the `.xlsx` file structure.

> Prefer using the static method `TemplateMemory.from()` to create instances.

---

### 📄 Properties

- `files: Record<string, Buffer>` — in-memory map of file contents.
- `destroyed: boolean` — indicates whether the instance has been destroyed (read-only).

---

### 📚 Methods

#### `copySheet(sourceName: string, newName: string): Promise<void>`

Creates a copy of an existing worksheet with a new name.

- `sourceName` — the name of the existing sheet.
- `newName` — the name for the new sheet.
- Throws if:
  - `sourceName` does not exist.
  - `newName` already exists.

---

#### `substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void>`

Replaces placeholders of the form `${key}` with values from the `replacements` object. For arrays, use placeholders with key `${table:key}`.

- `sheetName` — the name of the worksheet.
- `replacements` — key-value map for substitution.

---

#### `insertRows(data: { sheetName: string; startRowNumber?: number; rows: unknown[][] }): Promise<void>`

Inserts rows into a specified worksheet.

- `sheetName` — name of the worksheet.
- `startRowNumber` — starting row index (default: append to the end).
- `rows` — array of arrays, each representing a row of values.
- Throws if:
  - The sheet does not exist.
  - The row number is invalid.
  - A cell is out of bounds.

---

#### `insertRowsStream(data: { sheetName: string; startRowNumber?: number; rows: AsyncIterable<unknown[]> }): Promise<void>`

Streams and inserts rows into a worksheet, useful for handling large datasets.

- `sheetName` — name of the worksheet.
- `startRowNumber` — starting row index (default: append to the end).
- `rows` — an async iterable where each item is an array of cell values.
- Same error conditions as `insertRows`.

---

#### `save(): Promise<Buffer>`

Generates a new Excel file and returns it as a `Buffer`.

- Returns: `Promise<Buffer>` — the full `.xlsx` file contents in memory.
- Throws if:
  - The instance has been destroyed.
  - There was a failure while rebuilding the ZIP archive.

---

#### `set(key: string, content: Buffer): Promise<void>`

Replaces the content of a specific file in the template.

- `key` — the Excel path of the file (e.g., `xl/worksheets/sheet1.xml`).
- `content` — new file content as a Buffer.
- Throws if:
  - The instance has been destroyed.
  - The file does not exist in the template.

---

### 💡 Usage Examples

#### Create from File and Modify

```ts
const template = await TemplateMemory.from({
  source: fs.readFileSync("template.xlsx")
});

await template.copySheet("Sheet1", "Sheet2");
await template.substitute("Sheet1", { name: "John Doe" });

const modifiedExcel = await template.save();
fs.writeFileSync("output.xlsx", modifiedExcel);
```

#### Insert Rows from a Stream

```ts
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

#### Replace Internal File Content

```ts
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

### 🛑 Internal Checks

Methods perform validation:

- Ensures the instance hasn't been destroyed.
- Prevents concurrent modifications.
