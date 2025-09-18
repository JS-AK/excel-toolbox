# excel-toolbox

![ci-cd](https://github.com/JS-AK/excel-toolbox/actions/workflows/ci-cd-master.yml/badge.svg)

📘 **Docs:** [js-ak.github.io/excel-toolbox](https://js-ak.github.io/excel-toolbox/)

A lightweight toolkit for working with `.xlsx` Excel files — modify templates, merge sheets, and handle massive datasets without dependencies.

## Installation

```bash
npm install @js-ak/excel-toolbox
```

## Features

- ✨ Work with templates using `TemplateFs` (filesystem) or `TemplateMemory` (in-memory)
- 📥 Insert and stream rows into Excel files
- 🧩 Merge sheets from multiple `.xlsx` files
- 🧼 Remove sheets by name or index
- 💎 Preserve styles, merges, and shared strings
- 🏗️ Build `.xlsx` from scratch with an experimental `WorkbookBuilder`

## Template API

### `TemplateFs` and `TemplateMemory`

Both classes provide the same API for modifying Excel templates.

#### Common Features

- `substitute()` — replace placeholders like `${name}` or `${table:name}`
- `insertRows()` / `insertRowsStream()` — insert rows statically or via stream
- `copySheet()` — duplicate existing sheets
- `validate()` and `save()` / `saveStream()` — output the result

```ts
import { TemplateFs } from "@js-ak/excel-toolbox";

const template = await TemplateFs.from({
  destination: "/tmp/template",
  source: fs.readFileSync("template.xlsx"),
});

await template.substitute("Sheet1", { name: "Alice" });
await template.insertRows({ sheetName: "Sheet1", rows: [["Data"]] });
const buffer = await template.save();
fs.writeFileSync("output.xlsx", buffer);
```

## Sheet Merging API

### `mergeSheetsToBaseFileSync(options): Buffer`

Synchronously merges sheets into a base file.

### `mergeSheetsToBaseFile(options): Promise<Buffer>`

Async version of the above.

#### Example

```ts
import fs from "node:fs";
import { mergeSheetsToBaseFileSync } from "@js-ak/excel-toolbox";

const baseFile = fs.readFileSync("base.xlsx");
const dataFile = fs.readFileSync("data.xlsx");

const result = mergeSheetsToBaseFileSync({
  baseFile,
  additions: [{ file: dataFile, sheetIndexes: [1] }],
  gap: 2,
});

fs.writeFileSync("output.xlsx", result);
```

## Workbook Builder (experimental)

Programmatically create `.xlsx` files from scratch with streaming-friendly writes, shared strings, styles, and merges.

> Note: This API is experimental and may change. Import paths can differ depending on your setup.

### Quick start

```ts
// ESM (experimental deep import; subject to change)
import { WorkbookBuilder } from "@js-ak/excel-toolbox";

// Create a workbook and work with sheets
const wb = new WorkbookBuilder(); // default sheet: "Sheet1"

// Get the default sheet and set a few cells
const sheet1 = wb.getSheet("Sheet1");
sheet1?.setCell(1, 1, { value: "Hello" });      // row 1, col 1 (A1)
sheet1?.setCell(2, "B", { value: 123.45 });     // row 2, col B, number

// Add another sheet and use strings/styles/merges
const sheet2 = wb.addSheet("Report");
sheet2.setCell(1, 1, { type: "s", value: "Title" }); // shared string
sheet2.setCell(2, 1, {
  style: { alignment: { horizontal: "center" }, font: { bold: true } },
  value: "Centered Bold",
});
sheet2.addMerge({ startRow: 2, startCol: 1, endRow: 2, endCol: 3 }); // merge A2:C2

// Save to file
await wb.saveToFile("output.xlsx");
```

### Saving to a stream

```ts
import fs from "node:fs";

await wb.saveToStream(fs.createWriteStream("output-stream.xlsx"), {
  destination: ".cache/wb-temp", // optional: where temp files are created
  cleanup: true,                  // defaults to true
});
```

### Cells and types

- `setCell(rowIndex, column, cell)` where:
  - **rowIndex**: number (1-based, ≤ 1,048,576)
  - **column**: number (1-based) or Excel letter like `"A"`, `"AB"`
  - **cell**: `{ value, type?, style?, isFormula? }`
- If `type` is omitted, it is inferred:
  - number → `n`
  - boolean → `b`
  - string → `inlineStr`
- To use shared strings, set `type: "s"` and provide a string `value`.
- To write formulas, pass `{ isFormula: true, value: "SUM(A1:B1)" }` (type is ignored for formulas).

### Styling (basic)

```ts
sheet2.setCell(3, 1, {
  value: 42,
  style: {
    numberFormat: "0.00",
    alignment: { horizontal: "right" },
    font: { bold: true, color: "#0030FF" },
    fill: { fgColor: "FFFFEEEE" },
    border: {
      top: { style: "thin", color: "#0030FF" },
      bottom: { style: "thin", color: "#0030FF" },
      left: { style: "thin", color: "#0030FF" },
      right: { style: "thin", color: "#0030FF" },
    },
  },
});
```

### Performance tip

When generating millions of rows, periodically yield to the event loop (e.g., every 10k rows) to keep the process responsive.

## License

MIT — see [LICENSE](./LICENSE)
