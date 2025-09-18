# Workbook Builder

Programmatically build `.xlsx` files from scratch with streaming-friendly writes, shared strings, styles, and merges.

> ⚠️ Experimental API — surface may change in future releases.

---

## 🚀 Quick start

```ts
// ESM deep import (subject to change)
import { WorkbookBuilder } from "@js-ak/excel-toolbox";

const wb = new WorkbookBuilder(); // default sheet: "Sheet1"

// Get default sheet and set a few cells
const sheet1 = wb.getSheet("Sheet1");
sheet1?.setCell(1, 1, { value: "Hello" });      // A1
sheet1?.setCell(2, "B", { value: 123.45 });     // B2 number

// Add another sheet and use strings/styles/merges
const sheet2 = wb.addSheet("Report");

sheet2.setCell(1, 1, { type: "s", value: "Title" }); // shared string
sheet2.setCell(2, 1, {
  style: { alignment: { horizontal: "center" }, font: { bold: true } },
  value: "Centered Bold",
});

sheet2.addMerge({ startRow: 2, startCol: 1, endRow: 2, endCol: 3 }); // merge A2:C2

await wb.saveToFile("output.xlsx");
```

---

## 💾 Saving to a stream

```ts
import fs from "node:fs";

await wb.saveToStream(fs.createWriteStream("output-stream.xlsx"), {
  destination: ".cache/wb-temp", // optional directory for temp files
  cleanup: true,                  // default: true
});
```

---

## 🔢 Cells and types

- `setCell(rowIndex, column, cell)`
  - rowIndex: number (1-based, ≤ 1,048,576)
  - column: number (1-based) or Excel letter string (e.g., "A", "AB")
  - cell: `{ value, type?, style?, isFormula? }`

Type inference when `type` is omitted:

- number → `n`
- boolean → `b`
- string → `inlineStr`

Shared strings: set `type: "s"` and `value: string`. The builder will deduplicate and store them in `xl/sharedStrings.xml`.

Formulas: pass `{ isFormula: true, value: "SUM(A1:B1)" }` (type is ignored for formulas).

---

## 🎨 Styling (basic)

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

---

## 🔗 Merges

```ts
sheet2.addMerge({ startRow: 5, startCol: 2, endRow: 5, endCol: 6 }); // B5:F5
```

---

## 🧩 API surface

`new WorkbookBuilder(options?)`

- `options.defaultSheetName?: string` — name for the initial sheet (default: "Sheet1").

### Methods

- `addSheet(name: string): SheetData`
- `getSheet(name: string): SheetData | undefined`
- `removeSheet(name: string): true`
- `getInfo(): { sheetsNames, sharedStrings, styles, mergeCells, ... }` — frozen snapshot for inspection/tests
- `saveToFile(path: string): Promise<void>`
- `saveToStream(output: Writable, options?: { destination?: string; cleanup?: boolean }): Promise<void>`

`SheetData` helpers:

- `setCell(rowIndex: number, column: string | number, cell: CellData): void`
- `getCell(rowIndex: number, column: string | number): CellData | undefined`
- `removeCell(rowIndex: number, column: string | number): boolean`
- `addMerge(range: { startRow; startCol; endRow; endCol }): { ... }`
- `removeMerge(range): boolean`

---

## ⚡ Performance tips

When generating very large worksheets, periodically yield to the event loop (e.g., every 10k rows) to keep the process responsive.
