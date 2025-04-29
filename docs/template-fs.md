# TemplateFs

The `TemplateFs` class is designed for working with Excel (`.xlsx`) templates extracted to the filesystem. It enables modifying templates, inserting rows, substituting placeholders, and saving either as a `Buffer` or directly into a writable stream.

> ⚠️ **Experimental API**
> Interface is subject to change in future versions.

---

## 🔧 Constructor

```ts
new TemplateFs(fileKeys: Set<string>, destination: string)
```

- Input:
  - `fileKeys` — a set of relative file paths representing the `.xlsx` file structure.
  - `destination` — path to a directory where the template is extracted and modified.
- Output: `TemplateFs` instance
- Preconditions: None
- Postconditions: Instance is ready for use with provided files.

> Prefer using the static method `TemplateFs.from()` to create instances.

---

## 📄 Properties

- `fileKeys: Set<string>` — set of template file paths used for rebuilding the `.xlsx`.
- `destination: string` — working directory for extracted and modified files.
- `destroyed: boolean` — indicates whether the instance has been destroyed (read-only).

---

## 📚 Methods

### `copySheet`

Creates a copy of an existing worksheet with a new name.

- Input:
  - `sourceName: string` — name of the existing sheet.
  - `newName: string` — name for the new sheet.
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
  - `sheetName: string` — name of worksheet
  - `replacements: Record<string, unknown>` — key-value map for substitution
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - Sheet exists
- Postconditions:
  - Placeholders replaced with values
  - Shared strings updated if needed

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
  - Temporary files removed if necessary
  - ZIP archive created
- Throws if:
  - The instance has been destroyed.
  - Failure occurs while rebuilding the ZIP archive.

---

### `saveStream`

Writes the resulting Excel file directly to a writable stream.

- Input:
  - `output: Writable` — target writable stream (e.g., file, HTTP response).
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
- Postconditions:
  - Excel file streamed to output
- Throws if:
  - The instance has been destroyed.
  - Streaming or rebuilding fails.

---

### `validate`

Validates the internal state by checking if all required files exist.

- Input: None
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
- Postconditions:
  - Missing files detected (if any)
- Throws if:
  - The instance has been destroyed.
  - Any required file is missing.

---

### `set`

Replaces the content of a specific file in the template.

- Input:
  - `key: string` — relative Excel path (e.g., `xl/worksheets/sheet1.xml`)
  - `content: Buffer | string` — new file content
- Output: `Promise<void>`
- Preconditions:
  - Instance not destroyed
  - File exists
- Postconditions:
  - File content updated
- Throws if:
  - The instance has been destroyed.
  - The file does not exist.

---

### `from`

Creates a new `TemplateFs` instance from a source Excel file.

- Input:
  - `options: object` — configuration object
    - `source: string | Buffer` — path to source Excel file or its Buffer content
    - `destination: string` — directory path for extracted files
- Output: `Promise<TemplateFs>`
- Preconditions:
  - Valid source file/Buffer
  - Writable destination path
- Postconditions:
  - Files extracted to destination
  - Instance ready for use
- Throws if:
  - Source file invalid or missing
  - Destination path not writable
  - Extraction fails

---

## 💡 Usage Examples

### Insert Rows from a Stream and Save to File Stream

```ts
import { TemplateFs } from '@js-ak/excel-toolbox';

async function* asyncRowsGenerator(count: number): AsyncIterable<unknown[]> {
 for (let i = 0; i < count; i++) {
  await new Promise((resolve) => setTimeout(resolve, 0));
  yield Array(1000).fill(["Name"]);
 }
}

const template = await TemplateFs.from({
  destination: path.resolve(process.cwd(), "temp", crypto.randomUUID()),
  source: path.resolve(process.cwd(), "assets", `./input-01.xlsx`),
});

await template.insertRowsStream({
  rows: asyncRowGenerator(10),
  sheetName: "Sheet1",
  startRowNumber: 10,
});

await template.validate();

const outputStream = fs.createWriteStream(
  path.resolve(process.cwd(), "assets", `./output-01.xlsx`)
);

await template.saveStream(outputStream);
```

### Insert Rows from a Stream and Save to Buffer

```ts
import { TemplateFs } from '@js-ak/excel-toolbox';

const template = await TemplateFs.from({
  destination: path.resolve(process.cwd(), "temp", crypto.randomUUID()),
  source: path.resolve(process.cwd(), "assets", `./input-01.xlsx`),
});

await template.insertRowsStream({
  rows: dataStream(),
  sheetName: "Sheet1",
});

await template.validate();

const output = await template.save();

fs.writeFileSync(
  path.resolve(process.cwd(), "assets", `./output-01.xlsx`),
  output
);
```

### Copy Sheet and Apply Substitutions

```ts
import { TemplateFs } from '@js-ak/excel-toolbox';

const template = await TemplateFs.from({
  destination: path.resolve(process.cwd(), "temp", crypto.randomUUID()),
  source: path.resolve(process.cwd(), "assets", `./input-01.xlsx`),
});

await template.copySheet("Sheet1", "Sheet2");

await template.substitute("sheet1", {
  user: {
    age: { value: 1050 },
    name: "John Doe",
  },
  users: [
    { name: "John Doe" },
    { name: "Xyz Doe" },
  ],
});

await template.validate();

const outputStream = fs.createWriteStream(
  path.resolve(process.cwd(), "assets", `./output-01.xlsx`)
);

await template.saveStream(outputStream);
```

---

## 🛑 Internal Checks

Methods perform validation:

- Ensure the instance hasn't been destroyed.
- Prevent concurrent modifications.
- Ensure required files are present before saving.
