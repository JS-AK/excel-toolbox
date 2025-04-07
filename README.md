# excel-toolbox

![ci-cd](https://github.com/JS-AK/excel-toolbox/actions/workflows/ci-cd-master.yml/badge.svg)

A lightweight toolkit for merging sheets from multiple `.xlsx` Excel files without dependencies.

## Installation

You can install the package via npm:

```bash
npm install @js-ak/excel-toolbox
```

## Getting Started

To merge rows from multiple Excel files into one:

```ts
import fs from "node:fs";
import { mergeSheetsToBaseFileSync } from "@js-ak/excel-toolbox";

const baseFile = fs.readFileSync("base.xlsx");
const otherFile = fs.readFileSync("data.xlsx");

const resultBuffer = mergeSheetsToBaseFileSync({
  baseFile,
  additions: [
    { file: otherFile, sheetIndexes: [1] }
  ],
  gap: 2,
});

fs.writeFileSync("output.xlsx", resultBuffer);
```

## Features

- 🧩 **Merge sheets** from multiple Excel files
- 🧼 **Clean sheet removal** — by name or index
- 📎 **Keeps styles and merged cells**
- 🪶 **Lightweight ZIP and XML handling**

## API

### `mergeSheetsToBaseFileSync(options)`

#### Parameters

| Name                  | Type                                                               | Description                                    |
|-----------------------|--------------------------------------------------------------------|------------------------------------------------|
| `baseFile`            | `Buffer`                                                           | The base Excel file.                           |
| `additions`           | `{ file: Buffer; sheetIndexes: number[]; isBaseFile?: boolean }[]` | Files and sheet indices to merge.              |
| `baseSheetIndex`      | `number` (default: `1`)                                            | The sheet index in the base file to append to. |
| `gap`                 | `number` (default: `0`)                                            | Empty rows inserted between merged blocks.     |
| `sheetNamesToRemove`  | `string[]` (default: `[]`)                                         | Sheets to remove by name.                      |
| `sheetsToRemove`      | `number[]` (default: `[]`)                                         | Sheets to remove by index (1-based).           |

#### Returns

`Buffer` — the merged Excel file.

### `mergeSheetsToBaseFile(options): Promise<Buffer>`

Asynchronous version of `mergeSheetsToBaseFileSync`.

#### Parameters

Same as [`mergeSheetsToBaseFileSync`](#mergesheetstobasefilesyncoptions).

#### Returns

`Promise<Buffer>` — resolves with the merged Excel file.

#### Example

```ts
import fs from "node:fs/promises";
import { mergeSheetsToBaseFile } from "@js-ak/excel-toolbox";

const baseFile = await fs.readFile("base.xlsx");
const otherFile = await fs.readFile("data.xlsx");

const output = await mergeSheetsToBaseFile({
  baseFile,
  additions: [
    { file: otherFile, sheetIndexes: [1] }
  ],
  gap: 1,
});

await fs.writeFile("output.xlsx", output);
```

## Contributing

Contributions are welcome! Feel free to open an issue or submit a pull request if you have ideas or encounter bugs.

## License

MIT — see [LICENSE](./LICENSE) for details.
