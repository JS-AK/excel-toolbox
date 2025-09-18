# excel-toolbox

![ci-cd](https://github.com/JS-AK/excel-toolbox/actions/workflows/ci-cd-master.yml/badge.svg)

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
- `set()` — manually modify or inject internal files in the template

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

## License

MIT — see [LICENSE](./LICENSE)
