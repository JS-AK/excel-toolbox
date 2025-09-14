import fs from "node:fs";
import fsPromises from "node:fs/promises";
import path from "node:path";
import readline from "node:readline";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";

import type { MergeCell, RowData, SheetData } from "../types/index.js";

import { columnIndexToLetter } from "../../template/utils/column-index-to-letter.js";
import { columnLetterToIndex } from "../../template/utils/column-letter-to-index.js";

/**
 * Writes a worksheet XML file to the given destination using a streaming approach.
 *
 * - Streams rows and cells to minimize memory usage
 * - Tracks the worksheet dimension (min/max rows/cols) while streaming
 * - Updates the <dimension> ref at the end by rewriting the file in a buffered pass
 *
 * @param destination - Absolute or relative file path for the worksheet XML
 * @param rows - Map of rowIndex -> RowData with cells
 * @param merges - Merge ranges to append after <sheetData>
 * @returns Promise that resolves when the file is fully written and dimension updated
 */
export async function writeWorksheetXml(
	destination: string,
	rows: SheetData["rows"] = new Map<number, RowData>(),
	merges: MergeCell[] = [],
): Promise<void> {
	// ensure folder exists
	await fsPromises.mkdir(path.dirname(destination), { recursive: true });

	const stream = fs.createWriteStream(destination, { encoding: "utf-8" });

	// Track dimension incrementally while streaming
	let minRow = Number.POSITIVE_INFINITY;
	let maxRow = 0;
	let minCol = Number.POSITIVE_INFINITY;
	let maxCol = 0;

	try {
		// header
		stream.write(XML_DECLARATION + "\n");
		stream.write(`<worksheet xmlns="${XML_NAMESPACES.SPREADSHEET_ML}" xmlns:r="${XML_NAMESPACES.OFFICE_DOCUMENT}">\n`);

		// dimension
		stream.write("  <dimension ref=\"A1:A1\"/>\n");

		// sheetViews
		stream.write("  <sheetViews>\n");
		stream.write("    <sheetView workbookViewId=\"0\"/>\n");
		stream.write("  </sheetViews>\n");

		// sheetFormatPr
		stream.write("  <sheetFormatPr defaultRowHeight=\"15\"/>\n");

		// sheetData start
		stream.write("  <sheetData>\n");

		// rows (stream, without accumulation)
		const typeSetSkipping = new Set<string>(["n"]);
		let processedRows = 0;

		for (const [rowNumber, row] of rows) {
			if (row.cells.size > 0) {
				if (rowNumber < minRow) minRow = rowNumber;
				if (rowNumber > maxRow) maxRow = rowNumber;
			}

			// Build the row string
			const rowStart = `    <row r="${rowNumber}">\n`;
			const rowEnd = "    </row>\n";

			// write to stream with backpressure control
			if (!stream.write(rowStart)) {
				await new Promise<void>(resolve => stream.once("drain", () => resolve()));
			}

			// if you need to stream cells, you can do so as well:
			for (const [colNumber, cellData] of row.cells) {
				const colIndex = columnLetterToIndex(colNumber) - 1;

				if (colIndex < minCol) minCol = colIndex;
				if (colIndex > maxCol) maxCol = colIndex;

				const attrT = cellData.type && typeSetSkipping.has(cellData.type)
					? ""
					: ` t="${cellData.type}"`;
				const attrS = cellData.style?.index
					? ` s="${cellData.style?.index}"`
					: "";

				let cellXml = "";

				if (cellData.isFormula) {
					cellXml = `      <c r="${colNumber}${rowNumber}"${attrS}${attrT}><f>${cellData.value}</f></c>\n`;
				} else {
					switch (cellData.type) {
						case "b": {
							cellXml = `      <c r="${colNumber}${rowNumber}"${attrS}${attrT}><v>${cellData.value ? "1" : "0"}</v></c>\n`;

							break;
						}
						case "inlineStr": {
							cellXml = `      <c r="${colNumber}${rowNumber}"${attrS}${attrT}><is><t>${cellData.value}</t></is></c>\n`;

							break;
						}
						default: {
							cellXml = `      <c r="${colNumber}${rowNumber}"${attrS}${attrT}><v>${cellData.value}</v></c>\n`;

							break;
						}
					}
				}

				stream.write(cellXml);

				// if (!stream.write(cellXml)) {
				// 	await new Promise<void>(resolve => stream.once("drain", () => resolve()));
				// }
			}

			if (!stream.write(rowEnd)) {
				await new Promise<void>(resolve => stream.once("drain", () => resolve()));
			}

			// Unload the event loop every 100 rows
			processedRows++;
			if (processedRows % 100 === 0) {
				await new Promise<void>(resolve => setImmediate(resolve));
			}

			// Unload the event loop every 100 rows
			// rows.delete(rowNumber);
		}

		// sheetData end
		stream.write("  </sheetData>\n");

		// mergeCells
		if (merges.length > 0) {
			stream.write(`  <mergeCells count="${merges.length}">\n`);
			for (const merge of merges) {
				if (merge.startCol < minCol) minCol = merge.startCol;
				if (merge.endCol > maxCol) maxCol = merge.endCol;
				if (merge.startRow < minRow) minRow = merge.startRow;
				if (merge.endRow > maxRow) maxRow = merge.endRow;

				const ref = `${columnIndexToLetter(merge.startCol)}${merge.startRow}:${columnIndexToLetter(merge.endCol)}${merge.endRow}`;

				stream.write(`    <mergeCell ref="${ref}"/>\n`);
			}

			stream.write("  </mergeCells>\n");
		}

		// close worksheet
		stream.write("</worksheet>\n");
	} finally {
		stream.end();
	}

	let finalDimensionRef = "A1:A1";
	if (Number.isFinite(minCol) && Number.isFinite(minRow) && maxCol >= 0 && maxRow > 0) {
		finalDimensionRef = `${columnIndexToLetter(minCol)}${minRow}:${columnIndexToLetter(maxCol)}${maxRow}`;
	}

	await new Promise<void>((resolve, reject) => {
		stream.on("error", reject);
		stream.on("finish", resolve);
	});

	await updateWorksheetDimensionInFile(destination, finalDimensionRef);

	return;
}

// async function updateWorksheetDimensionInFile(destination: string, dimensionRef: string): Promise<void> {
// 	const tempPath = `${destination}.tmp`;

// 	const readStream = fs.createReadStream(destination, { encoding: "utf-8" });
// 	const rl = readline.createInterface({ crlfDelay: Infinity, input: readStream });
// 	const writeStream = fs.createWriteStream(tempPath, { encoding: "utf-8" });

// 	let updated = false;
// 	const rlClosed = new Promise<void>(resolve => rl.once("close", resolve));

// 	for await (const line of rl) {
// 		let outLine = line;

// 		if (!updated && line.includes("<dimension")) {
// 			outLine = line.replace(/(<dimension[^>]*ref=")[^"]*("[^>]*\/>)/, `$1${dimensionRef}$2`);
// 			updated = true;
// 		}

// 		if (!writeStream.write(outLine + "\n")) {
// 			await new Promise<void>(resolve => writeStream.once("drain", resolve));
// 		}
// 	}

// 	await rlClosed;

// 	writeStream.end();
// 	await new Promise<void>(resolve => writeStream.once("finish", resolve));

// 	await fsPromises.rename(tempPath, destination);
// }

/**
 * Updates the <dimension> ref attribute in an existing worksheet XML file.
 * Performs a buffered line-by-line rewrite to a temporary file, then renames it in place.
 *
 * @param destination - Path to the worksheet XML file to update
 * @param dimensionRef - Final dimension reference, e.g., "A1:C25"
 */
async function updateWorksheetDimensionInFile(
	destination: string,
	dimensionRef: string,
): Promise<void> {
	const tempPath = `${destination}.tmp`;

	const readStream = fs.createReadStream(destination, { encoding: "utf-8" });
	const rl = readline.createInterface({ crlfDelay: Infinity, input: readStream });
	const writeStream = fs.createWriteStream(tempPath, { encoding: "utf-8" });

	let updated = false;
	let buffer = "";

	const BUFFER_SIZE = 64 * 1024; // 64 KB

	for await (const line of rl) {
		let outLine = line;

		if (!updated && line.includes("<dimension")) {
			outLine = line.replace(
				/(<dimension[^>]*ref=")[^"]*("[^>]*\/>)/,
				`$1${dimensionRef}$2`,
			);
			updated = true;
		}

		buffer += outLine + "\n";

		if (buffer.length >= BUFFER_SIZE) {
			if (!writeStream.write(buffer)) {
				await new Promise<void>(resolve => writeStream.once("drain", resolve));
			}
			buffer = "";
		}
	}

	// Write remaining buffer
	if (buffer) writeStream.write(buffer);

	// Close streams
	await new Promise<void>(resolve => writeStream.end(resolve));

	// Replace original file
	await fsPromises.rename(tempPath, destination);
}
