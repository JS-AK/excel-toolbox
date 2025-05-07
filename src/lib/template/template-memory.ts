import * as fs from "node:fs/promises";

import * as Xml from "../xml/index.js";
import * as Zip from "../zip/index.js";

import * as Utils from "./utils/index.js";
import { MemoryWriteStream } from "./memory-write-stream.js";

/**
 * A class for manipulating Excel templates by extracting, modifying, and repacking Excel files.
 *
 * @experimental This API is experimental and might change in future versions.
 */
export class TemplateMemory {
	files: Record<string, Buffer>;

	/**
	 * Flag indicating whether this template instance has been destroyed.
	 * @type {boolean}
	 */
	destroyed: boolean = false;

	/**
	 * Flag indicating whether this template instance is currently being processed.
	 * @type {boolean}
	 */
	#isProcessing: boolean = false;

	/**
 * The keys for the Excel files in the template.
 */
	#excelKeys = {
		contentTypes: "[Content_Types].xml",
		sharedStrings: "xl/sharedStrings.xml",
		styles: "xl/styles.xml",
		workbook: "xl/workbook.xml",
		workbookRels: "xl/_rels/workbook.xml.rels",
	} as const;

	/**
	 * Creates a Template instance from a map of file paths to buffers.
	 *
	 * @param {Object<string, Buffer>} files - The files to create the template from.
	 * @throws {Error} If reading or writing files fails.
	 * @experimental This API is experimental and might change in future versions.
	 */
	constructor(files: Record<string, Buffer>) {
		this.files = files;
	}

	/** Private methods */

	/**
	 * Ensures that this Template instance has not been destroyed.
	 * @private
	 * @throws {Error} If the template instance has already been destroyed.
	 * @experimental This API is experimental and might change in future versions.
	 */
	#ensureNotDestroyed(): void {
		if (this.destroyed) {
			throw new Error("This Template instance has already been saved and destroyed.");
		}
	}

	/**
	 * Ensures that this Template instance is not currently being processed.
	 * @throws {Error} If the template instance is currently being processed.
	 * @experimental This API is experimental and might change in future versions.
	 */
	#ensureNotProcessing(): void {
		if (this.#isProcessing) {
			throw new Error("This Template instance is currently being processed.");
		}
	}

	/**
	 * Expand table rows in the given sheet and shared strings XML.
	 *
	 * @param {string} sheetXml - The XML content of the sheet.
	 * @param {string} sharedStringsXml - The XML content of the shared strings.
	 * @param {Record<string, unknown>} replacements - An object containing replacement values.
	 *
	 * @returns {Object} An object with two properties:
	 *   - sheet: The expanded sheet XML.
	 *   - shared: The expanded shared strings XML.
	 * @experimental This API is experimental and might change in future versions.
	 */
	#expandTableRows(
		sheetXml: string,
		sharedStringsXml: string,
		replacements: Record<string, unknown>,
	): { sheet: string; shared: string } | null {
		const {
			initialMergeCells,
			mergeCellMatches,
			modifiedXml,
		} = Utils.processMergeCells(sheetXml);

		const {
			sharedIndexMap,
			sharedStrings,
			sharedStringsHeader,
		} = Utils.processSharedStrings(sharedStringsXml);

		const {
			isTableReplacementsFound,
			lastIndex,
			resultRows,
			sheetMergeCells,
		} = Utils.processRows({
			mergeCellMatches,
			replacements,
			sharedIndexMap,
			sharedStrings,
			sheetXml: modifiedXml,
		});

		if (!isTableReplacementsFound) {
			return null;
		}

		const { shared, sheet } = Utils.processMergeFinalize({
			initialMergeCells,
			lastIndex,
			resultRows,
			sharedStrings,
			sharedStringsHeader,
			sheetMergeCells,
			sheetXml: modifiedXml,
		});

		return { shared, sheet };
	}

	/**
	 * Extracts the XML content from an Excel sheet file.
	 *
	 * @param {string} fileKey - The file key of the sheet to extract.
	 * @returns {string} The XML content of the sheet.
	 * @throws {Error} If the file key is not found.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #extractXmlFromSheet(fileKey: string): Promise<string> {
		if (!this.files[fileKey]) {
			throw new Error(`${fileKey} not found`);
		}

		return Xml.extractXmlFromSheet(this.files[fileKey]);
	}

	/**
	 * Extracts row data from an Excel sheet file.
	 *
	 * @param {string} fileKey - The file key of the sheet to extract.
	 * @returns {Object} An object containing:
	 *   - rows: Array of raw XML strings for each <row> element
	 *   - lastRowNumber: Highest row number found in the sheet (1-based)
	 *   - mergeCells: Array of merged cell ranges (e.g., [{ref: "A1:B2"}])
	 *   - xml: The XML content of the sheet
	 * @throws {Error} If the file key is not found
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #extractRowsFromSheet(fileKey: string): Promise<{
		rows: string[];
		lastRowNumber: number;
		mergeCells: { ref: string }[];
		xml: string;
	}> {
		if (!this.files[fileKey]) {
			throw new Error(`${fileKey} not found`);
		}

		return Xml.extractRowsFromSheet(this.files[fileKey]);
	}

	/**
	 * Returns the Excel path of the sheet with the given name.
	 *
	 * @param sheetName - The name of the sheet to find.
	 * @returns The Excel path of the sheet.
	 * @throws {Error} If the sheet with the given name does not exist.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #getSheetPathByName(sheetName: string): Promise<string> {
		// Find the sheet
		const workbookXml = await this.#extractXmlFromSheet(this.#excelKeys.workbook);
		const sheetMatch = workbookXml.match(Utils.sheetMatch(sheetName));

		if (!sheetMatch || !sheetMatch[1]) {
			throw new Error(`Sheet "${sheetName}" not found`);
		}

		const rId = sheetMatch[1];
		const relsXml = await this.#extractXmlFromSheet(this.#excelKeys.workbookRels);
		const relMatch = relsXml.match(Utils.relationshipMatch(rId));

		if (!relMatch || !relMatch[1]) {
			throw new Error(`Relationship "${rId}" not found`);
		}

		return "xl/" + relMatch[1].replace(/^\/?xl\//, "");
	}

	/**
	 * Returns the Excel path of the sheet with the given ID.
	 *
	 * @param {number} id - The 1-based index of the sheet to find.
	 * @returns {string} The Excel path of the sheet.
	 * @throws {Error} If the sheet index is less than 1.
	 * @experimental This API is experimental and might change in future versions.
	 */
	#getSheetPathById(id: number): string {
		if (id < 1) {
			throw new Error("Sheet index must be greater than 0");
		}

		return `xl/worksheets/sheet${id}.xml`;
	}

	/**
	 * Replaces the contents of a file in the template.
	 *
	 * @param {string} key - The Excel path of the file to replace.
	 * @param {Buffer|string} content - The new content.
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If the file does not exist in the template.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #set(key: string, content: Buffer): Promise<void> {
		this.files[key] = content;
	}

	/**
	 * Replaces placeholders in the given sheet with values from the replacements map.
	 *
	 * The function searches for placeholders in the format `${key}` within the sheet
	 * content, where `key` corresponds to a path in the replacements object.
	 * If a value is found for the key, it replaces the placeholder with the value.
	 * If no value is found, the original placeholder remains unchanged.
	 *
	 * @param sheetName - The name of the sheet to be replaced.
	 * @param replacements - An object where keys represent placeholder paths and values are the replacements.
	 * @returns A promise that resolves when the substitution is complete.
	 * @throws {Error} If the template instance has been destroyed.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void> {
		const sharedStringsPath = this.#excelKeys.sharedStrings;

		let sharedStringsContent = "";
		let sheetContent = "";

		if (this.files[sharedStringsPath]) {
			sharedStringsContent = await this.#extractXmlFromSheet(sharedStringsPath);
		}

		const sheetPath = await this.#getSheetPathByName(sheetName);

		if (this.files[sheetPath]) {
			sheetContent = await this.#extractXmlFromSheet(sheetPath);

			const TABLE_REGEX = /\$\{table:([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/g;

			const hasTablePlaceholders = TABLE_REGEX.test(sharedStringsContent) || TABLE_REGEX.test(sheetContent);
			const hasArraysInReplacements = Utils.foundArraysInReplacements(replacements);

			if (hasTablePlaceholders && hasArraysInReplacements) {
				const result = this.#expandTableRows(sheetContent, sharedStringsContent, replacements);

				if (result) {
					sheetContent = result.sheet;
					sharedStringsContent = result.shared;
				}
			}
		}

		if (this.files[sharedStringsPath]) {
			sharedStringsContent = Utils.applyReplacements(sharedStringsContent, replacements);

			await this.#set(sharedStringsPath, Buffer.from(sharedStringsContent));
		}

		if (this.files[sheetPath]) {
			sheetContent = Utils.applyReplacements(sheetContent, replacements);

			await this.#set(sheetPath, Buffer.from(sheetContent));
		}
	}

	/**
	 * Merges rows from other sheets into a base sheet.
	 *
	 * @param {Object} data
	 * @param {Object} data.additions
	 * @param {number[]} [data.additions.sheetIndexes] - The 1-based indexes of the sheets to extract rows from.
	 * @param {string[]} [data.additions.sheetNames] - The names of the sheets to extract rows from.
	 * @param {number} [data.baseSheetIndex=1] - The 1-based index of the sheet in the workbook to add rows to.
	 * @param {string} [data.baseSheetName] - The name of the sheet in the workbook to add rows to.
	 * @param {number} [data.gap=0] - The number of empty rows to insert between each added section.
	 * @throws {Error} If the base sheet index is less than 1.
	 * @throws {Error} If the base sheet name is not found.
	 * @throws {Error} If the sheet index is less than 1.
	 * @throws {Error} If the sheet name is not found.
	 * @throws {Error} If no sheets are found to merge.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #mergeSheets(data: {
		additions: { sheetIndexes?: number[]; sheetNames?: string[] };
		baseSheetIndex?: number;
		baseSheetName?: string;
		gap?: number;
	}): Promise<void> {
		const {
			additions,
			baseSheetIndex = 1,
			baseSheetName,
			gap = 0,
		} = data;

		let fileKey: string = "";

		if (baseSheetName) {
			fileKey = await this.#getSheetPathByName(baseSheetName);
		}

		if (baseSheetIndex && !fileKey) {
			if (baseSheetIndex < 1) {
				throw new Error("Base sheet index must be greater than 0");
			}

			fileKey = `xl/worksheets/sheet${baseSheetIndex}.xml`;
		}

		if (!fileKey) {
			throw new Error("Base sheet not found");
		}

		const {
			lastRowNumber,
			mergeCells: baseMergeCells,
			rows: baseRows,
			xml,
		} = await this.#extractRowsFromSheet(fileKey);

		const allRows = [...baseRows];
		const allMergeCells = [...baseMergeCells];
		let currentRowOffset = lastRowNumber + gap;

		const sheetPaths: string[] = [];

		if (additions.sheetIndexes) {
			sheetPaths.push(...(await Promise.all(additions.sheetIndexes.map(e => this.#getSheetPathById(e)))));
		}

		if (additions.sheetNames) {
			sheetPaths.push(...(await Promise.all(additions.sheetNames.map(e => this.#getSheetPathByName(e)))));
		}

		if (sheetPaths.length === 0) {
			throw new Error("No sheets found to merge");
		}

		for (const sheetPath of sheetPaths) {
			if (!this.files[sheetPath]) {
				throw new Error(`Sheet "${sheetPath}" not found`);
			}

			const { mergeCells, rows } = await Xml.extractRowsFromSheet(this.files[sheetPath]);

			const shiftedRows = Xml.shiftRowIndices(rows, currentRowOffset);

			const shiftedMergeCells = mergeCells.map(cell => {
				const [start, end] = cell.ref.split(":");

				if (!start || !end) {
					return cell;
				}

				const shiftedStart = Utils.Common.shiftCellRef(start, currentRowOffset);
				const shiftedEnd = Utils.Common.shiftCellRef(end, currentRowOffset);

				return { ...cell, ref: `${shiftedStart}:${shiftedEnd}` };
			});

			allRows.push(...shiftedRows);
			allMergeCells.push(...shiftedMergeCells);
			currentRowOffset += Utils.Common.getMaxRowNumber(rows) + gap;
		}

		const mergedXml = Xml.buildMergedSheet(
			xml,
			allRows,
			allMergeCells,
		);

		this.#set(fileKey, mergedXml);
	}

	/**
	 * Removes sheets from the workbook.
	 *
	 * @param {Object} data - The data for sheet removal.
	 * @param {number[]} [data.sheetIndexes] - The 1-based indexes of the sheets to remove.
	 * @param {string[]} [data.sheetNames] - The names of the sheets to remove.
	 * @returns {void}
	 *
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If the sheet does not exist.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #removeSheets(data: {
		sheetNames?: string[];
		sheetIndexes?: number[];
	}): Promise<void> {
		const { sheetIndexes = [], sheetNames = [] } = data;

		// first get index of sheets to remove
		const sheetIndexesToRemove: Set<number> = new Set(sheetIndexes);

		for (const sheetName of sheetNames) {
			const sheetPath = await this.#getSheetPathByName(sheetName);

			const sheetIndexMatch = sheetPath.match(/sheet(\d+)\.xml$/);

			if (!sheetIndexMatch || !sheetIndexMatch[1]) {
				throw new Error(`Sheet "${sheetName}" not found`);
			}

			const sheetIndex = parseInt(sheetIndexMatch[1], 10);

			sheetIndexesToRemove.add(sheetIndex);
		}

		// Remove sheets by index
		for (const sheetIndex of sheetIndexesToRemove.values()) {
			const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;

			if (!this.files[sheetPath]) {
				continue;
			}

			// remove sheet file
			delete this.files[sheetPath];

			// remove sheet from workbook
			const workbook = this.files[this.#excelKeys.workbook];
			if (workbook) {
				this.files[this.#excelKeys.workbook] = Buffer.from(Utils.Common.removeSheetFromWorkbook(
					workbook.toString(),
					sheetIndex,
				));
			}

			// remove sheet from workbook relations
			const workbookRels = this.files[this.#excelKeys.workbookRels];
			if (workbookRels) {
				this.files[this.#excelKeys.workbookRels] = Buffer.from(Utils.Common.removeSheetFromRels(
					workbookRels.toString(),
					sheetIndex,
				));
			}

			// remove sheet from content types
			const contentTypes = this.files[this.#excelKeys.contentTypes];
			if (contentTypes) {
				this.files[this.#excelKeys.contentTypes] = Buffer.from(Utils.Common.removeSheetFromContentTypes(
					contentTypes.toString(),
					sheetIndex,
				));
			}
		}
	}

	/** Public methods */

	/**
	 * Copies a sheet from the template to a new name.
	 *
	 * @param {string} sourceName - The name of the sheet to copy.
	 * @param {string} newName - The new name for the sheet.
	 * @returns {Promise<void>}
	 * @throws {Error} If the sheet with the source name does not exist.
	 * @throws {Error} If a sheet with the new name already exists.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async copySheet(sourceName: string, newName: string): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			if (sourceName === newName) {
				throw new Error("Source and new sheet names cannot be the same");
			}

			// Read workbook.xml and find the source sheet
			const workbookXmlPath = this.#excelKeys.workbook;
			const workbookXml = await this.#extractXmlFromSheet(this.#excelKeys.workbook);

			// Find the source sheet
			const sheetMatch = workbookXml.match(Utils.sheetMatch(sourceName));

			if (!sheetMatch || !sheetMatch[1]) {
				throw new Error(`Sheet "${sourceName}" not found`);
			}

			// Check if a sheet with the new name already exists
			if (new RegExp(`<sheet[^>]+name="${newName}"`).test(workbookXml)) {
				throw new Error(`Sheet "${newName}" already exists`);
			}

			// Read workbook.rels
			// Find the source sheet path by rId
			const rId = sheetMatch[1];
			const relsXmlPath = this.#excelKeys.workbookRels;
			const relsXml = await this.#extractXmlFromSheet(this.#excelKeys.workbookRels);
			const relMatch = relsXml.match(Utils.relationshipMatch(rId));

			if (!relMatch || !relMatch[1]) {
				throw new Error(`Relationship "${rId}" not found`);
			}

			const sourceTarget = relMatch[1]; // sheetN.xml
			const sourceSheetPath = "xl/" + sourceTarget.replace(/^\/?.*xl\//, "");

			// Get the index of the new sheet
			const sheetNumbers = Array.from(Object.keys(this.files))
				.map((key) => key.match(/^xl\/worksheets\/sheet(\d+)\.xml$/))
				.filter(Boolean)
				.map((match) => parseInt(match![1]!, 10));
			const nextSheetIndex = sheetNumbers.length > 0 ? Math.max(...sheetNumbers) + 1 : 1;

			const newSheetFilename = `sheet${nextSheetIndex}.xml`;
			const newSheetPath = `xl/worksheets/${newSheetFilename}`;
			const newTarget = `worksheets/${newSheetFilename}`;

			// Generate a unique rId
			const usedRIds = [...relsXml.matchAll(/Id="(rId\d+)"/g)].map(m => m[1]);
			let nextRIdNum = 1;
			while (usedRIds.includes(`rId${nextRIdNum}`)) nextRIdNum++;
			const newRId = `rId${nextRIdNum}`;

			// Copy the source sheet file
			const sheetContent = this.files[sourceSheetPath];

			if (!sheetContent) {
				throw new Error(`Sheet "${sourceSheetPath}" not found`);
			}

			function copyBuffer(source: Buffer): Buffer {
				const target: Buffer = Buffer.alloc(source.length);
				source.copy(target);
				return target;
			}

			await this.#set(newSheetPath, copyBuffer(sheetContent));

			// Update workbook.xml
			const updatedWorkbookXml = workbookXml.replace(
				"</sheets>",
				`<sheet name="${newName}" sheetId="${nextSheetIndex}" r:id="${newRId}"/></sheets>`,
			);

			await this.#set(workbookXmlPath, Buffer.from(updatedWorkbookXml));

			// Update workbook.xml.rels
			const updatedRelsXml = relsXml.replace(
				"</Relationships>",
				`<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${newTarget}"/></Relationships>`,
			);

			await this.#set(relsXmlPath, Buffer.from(updatedRelsXml));

			// Read [Content_Types].xml
			// Update [Content_Types].xml
			const contentTypesPath = "[Content_Types].xml";
			const contentTypesXml = await this.#extractXmlFromSheet(contentTypesPath);
			const overrideTag = `<Override PartName="/xl/worksheets/${newSheetFilename}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
			const updatedContentTypesXml = contentTypesXml.replace(
				"</Types>",
				overrideTag + "</Types>",
			);

			await this.#set(contentTypesPath, Buffer.from(updatedContentTypesXml));
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Replaces placeholders in the given sheet with values from the replacements map.
	 *
	 * The function searches for placeholders in the format `${key}` within the sheet
	 * content, where `key` corresponds to a path in the replacements object.
	 * If a value is found for the key, it replaces the placeholder with the value.
	 * If no value is found, the original placeholder remains unchanged.
	 *
	 * @param sheetName - The name of the sheet to be replaced.
	 * @param replacements - An object where keys represent placeholder paths and values are the replacements.
	 * @returns A promise that resolves when the substitution is complete.
	 */
	async substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			await this.#substitute(sheetName, replacements);
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Inserts rows into a specific sheet in the template.
	 *
	 * @param {Object} data - The data for row insertion.
	 * @param {string} data.sheetName - The name of the sheet to insert rows into.
	 * @param {number} [data.startRowNumber] - The row number to start inserting from.
	 * @param {unknown[][]} data.rows - The rows to insert.
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If the sheet does not exist.
	 * @throws {Error} If the row number is out of range.
	 * @throws {Error} If a column is out of range.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async insertRows(data: {
		sheetName: string;
		startRowNumber?: number;
		rows: unknown[][];
	}): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			const { rows, sheetName, startRowNumber } = data;

			const preparedRows = rows.map(row => Utils.toExcelColumnObject(row));

			// Validation
			Utils.checkStartRow(startRowNumber);
			Utils.checkRows(preparedRows);

			// Find the sheet
			const sheetPath = await this.#getSheetPathByName(sheetName);
			const sheetXml = await this.#extractXmlFromSheet(sheetPath);

			let nextRow = 0;

			if (!startRowNumber) {
				// Find the last row
				let lastRowNumber = 0;

				const rowMatches = [...sheetXml.matchAll(/<row[^>]+r="(\d+)"[^>]*>/g)];

				if (rowMatches.length > 0) {
					lastRowNumber = Math.max(...rowMatches.map((m) => parseInt(m[1] as string, 10)));
				}

				nextRow = lastRowNumber + 1;
			} else {
				nextRow = startRowNumber;
			}

			// Generate XML for all rows
			const rowsXml = preparedRows.map((cells, i) => {
				const rowNumber = nextRow + i;

				const cellTags = Object.entries(cells).map(([col, value]) => {
					const colUpper = col.toUpperCase();
					const ref = `${colUpper}${rowNumber}`;

					return `<c r="${ref}" t="inlineStr"><is><t>${Utils.escapeXml(value)}</t></is></c>`;
				}).join("");

				return `<row r="${rowNumber}">${cellTags}</row>`;
			}).join("");

			let updatedXml: string;

			if (/<sheetData\s*\/>/.test(sheetXml)) {
				updatedXml = sheetXml.replace(/<sheetData\s*\/>/, `<sheetData>${rowsXml}</sheetData>`);
			} else if (/<sheetData>([\s\S]*?)<\/sheetData>/.test(sheetXml)) {
				updatedXml = sheetXml.replace(/<\/sheetData>/, `${rowsXml}</sheetData>`);
			} else {
				updatedXml = sheetXml.replace(/<worksheet[^>]*>/, (match) => `${match}<sheetData>${rowsXml}</sheetData>`);
			}

			await this.#set(sheetPath, Buffer.from(Utils.updateDimension(updatedXml)));
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Inserts rows into a specific sheet in the template using an async stream.
	 *
	 * @param {Object} data - The data for row insertion.
	 * @param {string} data.sheetName - The name of the sheet to insert rows into.
	 * @param {number} [data.startRowNumber] - The row number to start inserting from.
	 * @param {AsyncIterable<unknown[]>} data.rows - Async iterable of rows to insert.
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If the sheet does not exist.
	 * @throws {Error} If the row number is out of range.
	 * @throws {Error} If a column is out of range.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async insertRowsStream(data: {
		sheetName: string;
		startRowNumber?: number;
		rows: AsyncIterable<unknown[]>;
	}): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			const { rows, sheetName, startRowNumber } = data;

			if (!sheetName) throw new Error("Sheet name is required");

			// Read XML workbook to find sheet name and path
			const sheetPath = await this.#getSheetPathByName(sheetName);
			const sheetXml = await this.#extractXmlFromSheet(sheetPath);

			const output = new MemoryWriteStream();

			let inserted = false;

			const initialDimension = sheetXml.match(/<dimension\s+ref="[^"]*"/)?.[0] || "";

			const dimension = {
				maxColumn: "A",
				maxRow: 1,
				minColumn: "A",
				minRow: 1,
			};

			if (initialDimension) {
				const dimensionMatch = initialDimension.match(/<dimension\s+ref="([^"]*)"/);
				if (dimensionMatch) {
					const dimensionRef = dimensionMatch[1];

					if (dimensionRef) {
						const [min, max] = dimensionRef.split(":");

						dimension.minColumn = min!.slice(0, 1);
						dimension.minRow = parseInt(min!.slice(1));
						dimension.maxColumn = max!.slice(0, 1);
						dimension.maxRow = parseInt(max!.slice(1));
					}
				}
			}

			// --- Case 1: <sheetData>...</sheetData> on one line ---
			const singleLineMatch = sheetXml.match(/(<sheetData[^>]*>)(.*)(<\/sheetData>)/);
			if (!inserted && singleLineMatch) {
				const maxRowNumber = startRowNumber ?? Utils.getMaxRowNumber(sheetXml);

				const openTag = singleLineMatch[1];
				const innerRows = singleLineMatch[2]!.trim();
				const closeTag = singleLineMatch[3];

				const innerRowsMap = Utils.parseRows(innerRows);

				output.write(sheetXml.slice(0, singleLineMatch.index!));
				output.write(openTag!);

				if (innerRows) {
					if (startRowNumber) {
						const filtered = Utils.getRowsBelow(innerRowsMap, startRowNumber);
						if (filtered) output.write(filtered);
					} else {
						output.write(innerRows);
					}
				}

				const { dimension: newDimension, rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

				if (Utils.compareColumns(newDimension.maxColumn, dimension.maxColumn) > 0) {
					dimension.maxColumn = newDimension.maxColumn;
				}

				if (newDimension.maxRow > dimension.maxRow) {
					dimension.maxRow = newDimension.maxRow;
				}

				if (innerRows) {
					const filtered = Utils.getRowsAbove(innerRowsMap, actualRowNumber);
					if (filtered) output.write(filtered);
				}

				output.write(closeTag!);
				output.write(sheetXml.slice(singleLineMatch.index! + singleLineMatch[0].length));
				inserted = true;
			}

			// --- Case 2: <sheetData/> ---
			if (!inserted && /<sheetData\s*\/>/.test(sheetXml)) {
				const maxRowNumber = startRowNumber ?? Utils.getMaxRowNumber(sheetXml);
				const match = sheetXml.match(/<sheetData\s*\/>/)!;
				const matchIndex = match.index!;

				output.write(sheetXml.slice(0, matchIndex));
				output.write("<sheetData>");

				const { dimension: newDimension } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

				if (Utils.compareColumns(newDimension.maxColumn, dimension.maxColumn) > 0) {
					dimension.maxColumn = newDimension.maxColumn;
				}

				if (newDimension.maxRow > dimension.maxRow) {
					dimension.maxRow = newDimension.maxRow;
				}

				output.write("</sheetData>");
				output.write(sheetXml.slice(matchIndex + match[0].length));
				inserted = true;
			}

			// --- Case 3: Multiline <sheetData> ---
			if (!inserted && sheetXml.includes("<sheetData")) {
				const openTagMatch = sheetXml.match(/<sheetData[^>]*>/);
				const closeTag = "</sheetData>";
				if (!openTagMatch) throw new Error("Invalid sheetData structure");

				const openTag = openTagMatch[0];
				const openIdx = sheetXml.indexOf(openTag);
				const closeIdx = sheetXml.lastIndexOf(closeTag);
				if (closeIdx === -1) throw new Error("Missing </sheetData>");

				const beforeRows = sheetXml.slice(0, openIdx + openTag.length);
				const innerRows = sheetXml.slice(openIdx + openTag.length, closeIdx).trim();
				const afterRows = sheetXml.slice(closeIdx + closeTag.length);

				const innerRowsMap = Utils.parseRows(innerRows);

				output.write(beforeRows);

				if (innerRows) {
					if (startRowNumber) {
						const filtered = Utils.getRowsBelow(innerRowsMap, startRowNumber);
						if (filtered) output.write(filtered);
					} else {
						output.write(innerRows);
					}
				}

				const { dimension: newDimension, rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, Utils.getMaxRowNumber(innerRows));

				if (Utils.compareColumns(newDimension.maxColumn, dimension.maxColumn) > 0) {
					dimension.maxColumn = newDimension.maxColumn;
				}

				if (newDimension.maxRow > dimension.maxRow) {
					dimension.maxRow = newDimension.maxRow;
				}

				if (innerRows) {
					const filtered = Utils.getRowsAbove(innerRowsMap, actualRowNumber);
					if (filtered) output.write(filtered);
				}

				output.write(closeTag);
				output.write(afterRows);
				inserted = true;
			}

			if (!inserted) throw new Error("Failed to locate <sheetData> for insertion");

			let result = output.toBuffer();

			// update dimension
			{
				const target = initialDimension;
				const refRange = `${dimension.minColumn}${dimension.minRow}:${dimension.maxColumn}${dimension.maxRow}`;
				const replacement = `<dimension ref="${refRange}"`;

				if (target) {
					result = Buffer.from(result.toString().replace(target, replacement));
				}
			}

			// Save the buffer to the sheet
			this.files[sheetPath] = result;
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Saves the modified Excel template to a buffer.
	 *
	 * @returns {Promise<Buffer>} The modified Excel template as a buffer.
	 * @throws {Error} If the template instance has been destroyed.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async save(): Promise<Buffer> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			const zipBuffer = await Zip.create(this.files);

			this.destroyed = true;

			// Clear all buffers
			for (const key in this.files) {
				if (this.files.hasOwnProperty(key)) {
					this.files[key] = Buffer.alloc(0); // Clear the buffer
				}
			}

			return zipBuffer;
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Replaces the contents of a file in the template.
	 *
	 * @param {string} key - The Excel path of the file to replace.
	 * @param {Buffer|string} content - The new content.
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If the file does not exist in the template.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async set(key: string, content: Buffer): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			await this.#set(key, content);
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Merges sheets into a base sheet.
	 *
	 * @param {Object} data
	 * @param {{ sheetIndexes?: number[]; sheetNames?: string[] }} data.additions - The sheets to merge.
	 * @param {number} [data.baseSheetIndex=1] - The 1-based index of the sheet to merge into.
	 * @param {string} [data.baseSheetName] - The name of the sheet to merge into.
	 * @param {number} [data.gap=0] - The number of empty rows to insert between each added section.
	 * @returns {void}
	 */
	async mergeSheets(data: {
		additions: { sheetIndexes?: number[]; sheetNames?: string[] };
		baseSheetIndex?: number;
		baseSheetName?: string;
		gap?: number;
	}): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			await this.#mergeSheets(data);
		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Removes sheets from the workbook.
	 *
	 * @param {Object} data
	 * @param {number[]} [data.sheetIndexes] - The 1-based indexes of the sheets to remove.
	 * @param {string[]} [data.sheetNames] - The names of the sheets to remove.
	 * @returns {void}
	 */
	async removeSheets(data: {
		sheetNames?: string[];
		sheetIndexes?: number[];
	}): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			await this.#removeSheets(data);
		} finally {
			this.#isProcessing = false;
		}
	}

	/** Static methods */

	/**
	 * Creates a Template instance from an Excel file source.
	 *
	 * @param {Object} data - The data to create the template from.
	 * @param {string | Buffer} data.source - The path or buffer of the Excel file.
	 * @returns {Promise<TemplateMemory>} A new Template instance.
	 * @throws {Error} If reading the file fails.
	 * @experimental This API is experimental and might change in future versions.
	 */
	static async from(data: {
		source: string | Buffer;
	}): Promise<TemplateMemory> {
		const { source } = data;

		const buffer = typeof source === "string"
			? await fs.readFile(source)
			: source;

		const files = await Zip.read(buffer);

		return new TemplateMemory(files);
	}
}
