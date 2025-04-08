import * as Utils from "./utils/index.js";
import * as Xml from "./xml/index.js";

/**
 * Merges rows from other Excel files into a base Excel file.
 *
 * This function is a process-friendly version of mergeSheetsToBaseFile.
 * It takes a single object with the following properties:
 * - additions: An array of objects with two properties:
 *   - files: A dictionary of file paths to their corresponding XML content
 *   - sheetIndexes: The 1-based indexes of the sheet to extract rows from
 * - baseFiles: A dictionary of file paths to their corresponding XML content
 * - baseSheetIndex: The 1-based index of the sheet in the base file to add rows to
 * - gap: The number of empty rows to insert between each added section
 * - sheetNamesToRemove: The names of sheets to remove from the output file
 * - sheetsToRemove: The 1-based indices of sheets to remove from the output file
 *
 * The function returns a dictionary of file paths to their corresponding XML content.
 */
export function mergeSheetsToBaseFileProcess(data: {
	additions: { files: Record<string, Buffer>; sheetIndexes: number[] }[];
	baseFiles: Record<string, Buffer>;
	baseSheetIndex: number;
	gap: number;
	sheetNamesToRemove: string[];
	sheetsToRemove: number[];
}): void {
	const {
		additions,
		baseFiles,
		baseSheetIndex,
		gap,
		sheetNamesToRemove,
		sheetsToRemove,
	} = data;

	const basePath = `xl/worksheets/sheet${baseSheetIndex}.xml`;

	if (!baseFiles[basePath]) {
		throw new Error(`Base file does not contain ${basePath}`);
	}

	const {
		lastRowNumber,
		mergeCells: baseMergeCells,
		rows: baseRows,
		xml,
	} = Xml.extractRowsFromSheet(baseFiles[basePath]);

	const allRows = [...baseRows];
	const allMergeCells = [...baseMergeCells];
	let currentRowOffset = lastRowNumber + gap;

	for (const { files, sheetIndexes } of additions) {
		for (const sheetIndex of sheetIndexes) {
			const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;

			if (!files[sheetPath]) {
				throw new Error(`File does not contain ${sheetPath}`);
			}

			const { mergeCells, rows } = Xml.extractRowsFromSheet(files[sheetPath]);

			const shiftedRows = Xml.shiftRowIndices(rows, currentRowOffset);

			const shiftedMergeCells = mergeCells.map(cell => {
				const [start, end] = cell.ref.split(":");

				if (!start || !end) {
					return cell;
				}

				const shiftedStart = Utils.shiftCellRef(start, currentRowOffset);
				const shiftedEnd = Utils.shiftCellRef(end, currentRowOffset);

				return { ...cell, ref: `${shiftedStart}:${shiftedEnd}` };
			});

			allRows.push(...shiftedRows);
			allMergeCells.push(...shiftedMergeCells);
			currentRowOffset += Utils.getMaxRowNumber(rows) + gap;
		}
	}

	const mergedXml = Xml.buildMergedSheet(
		xml,
		allRows,
		allMergeCells,
	);

	baseFiles[basePath] = mergedXml;

	for (const sheetIndex of sheetsToRemove) {
		const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;
		delete baseFiles[sheetPath];

		if (baseFiles["xl/workbook.xml"]) {
			baseFiles["xl/workbook.xml"] = Buffer.from(Utils.removeSheetFromWorkbook(
				baseFiles["xl/workbook.xml"].toString(),
				sheetIndex,
			));
		}

		if (baseFiles["xl/_rels/workbook.xml.rels"]) {
			baseFiles["xl/_rels/workbook.xml.rels"] = Buffer.from(Utils.removeSheetFromRels(
				baseFiles["xl/_rels/workbook.xml.rels"].toString(),
				sheetIndex,
			));
		}

		if (baseFiles["[Content_Types].xml"]) {
			baseFiles["[Content_Types].xml"] = Buffer.from(Utils.removeSheetFromContentTypes(
				baseFiles["[Content_Types].xml"].toString(),
				sheetIndex,
			));
		}
	}

	for (const sheetName of sheetNamesToRemove) {
		Utils.removeSheetByName(baseFiles, sheetName);
	}
}
