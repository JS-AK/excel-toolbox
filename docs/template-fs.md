# TemplateFs

The `TemplateFs` class is designed for working with Excel (`.xlsx`) templates. It supports extracting, modifying, and rebuilding Excel files. Typical use cases include placeholder substitution, sheet duplication, and row insertion.

> ⚠️ **Experimental API**
> Interface is subject to change in future versions.

---

## 🔧 Constructor

```ts
new TemplateFs(fileKeys: Set<string>, destination: string)
```

- `fileKeys` — a set of relative file paths that make up the Excel template.
- `destination` — a path to a directory where the template is extracted and edited.

> Prefer using the static method `TemplateFs.from()` to create instances.

---

### 📄 Properties

- `fileKeys: Set<string>` — the set of template file paths involved in final assembly.
- `destination: string` — the working directory where files are extracted and edited.
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

Replaces placeholders of the form `${key}` with values from the `replacements` object. For arrays use placeholders with key `${table:key}`

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

#### `validate(): Promise<void>`

Validates the template by checking all required files exist.

- Returns: `Promise<void>`
- Throws:
  - If the template instance has been destroyed.
  - If any required files are missing.

---

#### `save(): Promise<Buffer>`

Generates a new Excel file and returns it as a `Buffer`.

- Returns: `Promise<Buffer>` — the full `.xlsx` file contents in memory.
- Throws if:
  - The instance has been destroyed.
  - There was a failure while rebuilding the ZIP archive.

---

#### `saveStream(output: Writable): Promise<void>`

Writes the resulting Excel file to a writable stream.

- `output` — any writable stream, e.g. a file or HTTP response.
- Throws if:
  - The instance has been destroyed.
  - There was a failure during streaming or rebuilding the ZIP archive.

---

### 💡 Usage Examples

#### Insert Rows from a Stream and Save to File Stream

```ts
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

#### Insert Rows from a Stream and Save to Buffer

```ts
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

#### Copy Sheet and Apply Substitutions

```ts
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

### 🛑 Internal Checks

Methods perform validation:

- Ensures the instance hasn't been destroyed.
- Prevents concurrent modifications.
