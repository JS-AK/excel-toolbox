import * as fs from "node:fs/promises";
import * as fsSync from "node:fs";
import * as path from "node:path";
import * as readline from "node:readline";
import { Writable } from "node:stream";
import crypto from "node:crypto";

import * as Xml from "../xml/index.js";
import * as Zip from "../zip/index.js";

import * as Utils from "./utils/index.js";

/**
 * A class for manipulating Excel templates by extracting, modifying, and repacking Excel files.
 *
 * @experimental This API is experimental and might change in future versions.
 */
export class TemplateFs {
	/**
	 * Set of file paths (relative to the template) that will be used to create a new workbook.
	 * @type {Set<string>}
	 */
	fileKeys: Set<string>;

	/**
	 * The path where the template will be expanded.
	 * @type {string}
	 */
	destination: string;

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
	 * Creates a Template instance.
	 *
	 * @param {Set<string>} fileKeys - Set of file paths (relative to the template) that will be used to create a new workbook.
	 * @param {string} destination - The path where the template will be expanded.
	 * @experimental This API is experimental and might change in future versions.
	 */
	constructor(fileKeys: Set<string>, destination: string) {
		this.fileKeys = fileKeys;
		this.destination = destination;
	}

	/** Private methods */

	/**
	 * Removes the temporary directory created by this Template instance.
	 * @private
	 * @returns {Promise<void>}
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #cleanup(): Promise<void> {
		await fs.rm(this.destination, { force: true, recursive: true });
	}

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

		return Utils.processMergeFinalize({
			initialMergeCells,
			lastIndex,
			resultRows,
			sharedStrings,
			sharedStringsHeader,
			sheetMergeCells,
			sheetXml: modifiedXml,
		});
	}

	/**
	 * Get the path of the sheet with the given name inside the workbook.
	 * @param sheetName The name of the sheet to find.
	 * @returns The path of the sheet inside the workbook.
	 * @throws {Error} If the sheet is not found.
	 */
	async #getSheetPathByName(sheetName: string): Promise<string> {
		// Read XML workbook to find sheet name and path
		const workbookXml = await Xml.extractXmlFromSheet(await this.#readFile(this.#excelKeys.workbook));
		const sheetMatch = workbookXml.match(Utils.sheetMatch(sheetName));

		if (!sheetMatch || !sheetMatch[1]) {
			throw new Error(`Sheet "${sheetName}" not found`);
		}

		const rId = sheetMatch[1];
		const relsXml = await Xml.extractXmlFromSheet(await this.#readFile(this.#excelKeys.workbookRels));
		const relMatch = relsXml.match(Utils.relationshipMatch(rId));

		if (!relMatch || !relMatch[1]) {
			throw new Error(`Relationship "${rId}" not found`);
		}

		return "xl/" + relMatch[1].replace(/^\/?xl\//, "");
	}

	/**
	 * Reads all files from the destination directory.
	 * @private
	 * @returns {Promise<Record<string, Buffer>>} An object with file keys and their contents as Buffers.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #readAllFromDestination(): Promise<Record<string, Buffer>> {
		const result: Record<string, Buffer> = {};

		for (const key of this.fileKeys) {
			const fullPath = path.join(this.destination, ...key.split("/"));

			result[key] = await fs.readFile(fullPath);
		}

		return result;
	}

	/**
	 * Reads a single file from the destination directory.
	 * @private
	 * @param {string} pathKey - The Excel path of the file to read.
	 * @returns {Promise<Buffer>} The contents of the file as a Buffer.
	 * @throws {Error} If the file does not exist in the template.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #readFile(pathKey: string): Promise<Buffer> {
		if (!this.fileKeys.has(pathKey)) {
			throw new Error(`File "${pathKey}" not found in template.`);
		}

		const fullPath = path.join(this.destination, ...pathKey.split("/"));

		return await fs.readFile(fullPath);
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
	async #set(key: string, content: Buffer | string): Promise<void> {
		if (!this.fileKeys.has(key)) {
			throw new Error(`File "${key}" is not part of the original template.`);
		}

		const fullPath = path.join(this.destination, ...key.split("/"));

		await fs.writeFile(fullPath, Buffer.isBuffer(content) ? content : Buffer.from(content));
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
		const sheetPath = await this.#getSheetPathByName(sheetName);

		let sharedStringsContent = "";
		let sheetContent = "";

		if (this.fileKeys.has(sharedStringsPath)) {
			sharedStringsContent = await Xml.extractXmlFromSheet(await this.#readFile(sharedStringsPath));
		}

		if (this.fileKeys.has(sheetPath)) {
			sheetContent = await Xml.extractXmlFromSheet(await this.#readFile(sheetPath));

			const TABLE_REGEX = /\$\{table:([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/g;

			const hasTablePlaceholders = TABLE_REGEX.test(sharedStringsContent) || TABLE_REGEX.test(sheetContent);

			if (hasTablePlaceholders) {
				const result = this.#expandTableRows(sheetContent, sharedStringsContent, replacements);

				if (result) {
					sheetContent = result.sheet;
					sharedStringsContent = result.shared;
				}
			}
		}

		if (this.fileKeys.has(sharedStringsPath)) {
			sharedStringsContent = Utils.applyReplacements(sharedStringsContent, replacements);
			await this.#set(sharedStringsPath, sharedStringsContent);
		}

		if (this.fileKeys.has(sheetPath)) {
			sheetContent = Utils.applyReplacements(sheetContent, replacements);
			await this.#set(sheetPath, sheetContent);
		}
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

			if (!this.fileKeys.has(sheetPath)) {
				continue;
			}

			// remove sheet file
			await fs.unlink(path.join(this.destination, ...sheetPath.split("/")));
			this.fileKeys.delete(sheetPath);

			// remove sheet from workbook
			if (this.fileKeys.has(this.#excelKeys.workbook)) {
				this.#set(this.#excelKeys.workbook, Buffer.from(Utils.Common.removeSheetFromWorkbook(
					this.#readFile(this.#excelKeys.workbook).toString(),
					sheetIndex,
				)));
			}

			// remove sheet from workbook relations
			if (this.fileKeys.has(this.#excelKeys.workbookRels)) {
				this.#set(this.#excelKeys.workbookRels, Buffer.from(Utils.Common.removeSheetFromRels(
					this.#readFile(this.#excelKeys.workbookRels).toString(),
					sheetIndex,
				)));
			}

			// remove sheet from content types
			if (this.fileKeys.has(this.#excelKeys.contentTypes)) {
				this.#set(this.#excelKeys.contentTypes, Buffer.from(Utils.Common.removeSheetFromContentTypes(
					this.#readFile(this.#excelKeys.contentTypes).toString(),
					sheetIndex,
				)));
			}
		}
	}

	/**
	 * Validates the template by checking all required files exist.
	 *
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If any required files are missing.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async #validate(): Promise<void> {
		for (const key of this.fileKeys) {
			const fullPath = path.join(this.destination, ...key.split("/"));
			try {
				await fs.access(fullPath);
			} catch {
				throw new Error(`Missing file in template directory: ${key}`);
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
			const workbookXml = await Xml.extractXmlFromSheet(await this.#readFile(workbookXmlPath));

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
			const relsXml = await Xml.extractXmlFromSheet(await this.#readFile(relsXmlPath));
			const relMatch = relsXml.match(Utils.relationshipMatch(rId));

			if (!relMatch || !relMatch[1]) {
				throw new Error(`Relationship "${rId}" not found`);
			}

			const sourceTarget = relMatch[1]; // sheetN.xml
			const sourceSheetPath = "xl/" + sourceTarget.replace(/^\/?.*xl\//, "");

			// Get the index of the new sheet
			const sheetNumbers = [...this.fileKeys]
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
			const sheetContent = await this.#readFile(sourceSheetPath);
			await fs.writeFile(path.join(this.destination, ...newSheetPath.split("/")), sheetContent);
			this.fileKeys.add(newSheetPath);

			// Update workbook.xml
			const updatedWorkbookXml = workbookXml.replace(
				"</sheets>",
				`<sheet name="${newName}" sheetId="${nextSheetIndex}" r:id="${newRId}"/></sheets>`,
			);
			await this.#set(workbookXmlPath, updatedWorkbookXml);

			// Update workbook.xml.rels
			const updatedRelsXml = relsXml.replace(
				"</Relationships>",
				`<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${newTarget}"/></Relationships>`,
			);
			await this.#set(relsXmlPath, updatedRelsXml);

			// Read [Content_Types].xml
			// Update [Content_Types].xml
			const contentTypesPath = this.#excelKeys.contentTypes;
			const contentTypesXml = await Xml.extractXmlFromSheet(await this.#readFile(contentTypesPath));
			const overrideTag = `<Override PartName="/xl/worksheets/${newSheetFilename}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
			const updatedContentTypesXml = contentTypesXml.replace(
				"</Types>",
				overrideTag + "</Types>",
			);

			await this.#set(contentTypesPath, updatedContentTypesXml);
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

			const sheetXmlRaw = await this.#readFile(sheetPath);
			const sheetXml = await Xml.extractXmlFromSheet(sheetXmlRaw);

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

			await this.#set(sheetPath, Utils.updateDimension(updatedXml));
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

			// Get the path to the sheet
			const sheetPath = await this.#getSheetPathByName(sheetName);

			// The temporary file for writing
			const fullPath = path.join(this.destination, ...sheetPath.split("/"));
			const tempPath = fullPath + ".tmp";

			// Streams for reading and writing
			const input = fsSync.createReadStream(fullPath, { encoding: "utf-8" });
			const output = fsSync.createWriteStream(tempPath, { encoding: "utf-8" });

			// Inserted rows flag
			let inserted = false;

			let initialDimension = "";

			const dimension = {
				maxColumn: "A",
				maxRow: 1,
				minColumn: "A",
				minRow: 1,
			};

			const rl = readline.createInterface({
				// Process all line breaks
				crlfDelay: Infinity,
				input,
			});

			let isCollecting = false;
			let collection = "";

			for await (const line of rl) {
				// Process <dimension>
				if (!initialDimension && /<dimension\s+ref="[^"]*"/.test(line)) {
					const dimensionMatch = line.match(/<dimension\s+ref="([^"]*)"/);

					if (dimensionMatch) {
						const dimensionRef = dimensionMatch[1];

						if (dimensionRef) {
							const [min, max] = dimensionRef.split(":");

							dimension.minColumn = min!.slice(0, 1);
							dimension.minRow = parseInt(min!.slice(1));
							dimension.maxColumn = max!.slice(0, 1);
							dimension.maxRow = parseInt(max!.slice(1));
						}

						initialDimension = line.match(/<dimension\s+ref="[^"]*"/)?.[0] || "";
					}
				}

				// Collect lines between <sheetData> and </sheetData>
				if (!inserted && isCollecting) {
					collection += line;

					if (line.includes("</sheetData>")) {
						const maxRowNumber = startRowNumber ?? Utils.getMaxRowNumber(line);

						isCollecting = false;
						inserted = true;

						const openTag = collection.match(/<sheetData[^>]*>/)?.[0] ?? "<sheetData>";
						const closeTag = "</sheetData>";

						const openIdx = collection.indexOf(openTag);
						const closeIdx = collection.lastIndexOf(closeTag);

						const beforeRows = collection.slice(0, openIdx + openTag.length);
						const innerRows = collection.slice(openIdx + openTag.length, closeIdx).trim();
						const afterRows = collection.slice(closeIdx);

						output.write(beforeRows);

						const innerRowsMap = Utils.parseRows(innerRows);

						if (innerRows) {
							if (startRowNumber) {
								const filteredRows = Utils.getRowsBelow(innerRowsMap, startRowNumber);
								if (filteredRows) output.write(filteredRows);
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
							const filteredRows = Utils.getRowsAbove(innerRowsMap, actualRowNumber);
							if (filteredRows) output.write(filteredRows);
						}

						output.write(afterRows);
					}

					continue;
				}

				// Case 1: <sheetData> and </sheetData> on one line
				const singleLineMatch = line.match(/(<sheetData[^>]*>)(.*)(<\/sheetData>)/);

				if (!inserted && singleLineMatch) {
					const maxRowNumber = startRowNumber ?? Utils.getMaxRowNumber(line);

					const fullMatch = singleLineMatch[0];

					const before = line.slice(0, singleLineMatch.index);
					const after = line.slice(singleLineMatch.index! + fullMatch.length);

					const openTag = "<sheetData>";
					const closeTag = "</sheetData>";

					const openIdx = fullMatch.indexOf(openTag);
					const closeIdx = fullMatch.indexOf(closeTag);

					const beforeRows = fullMatch.slice(0, openIdx + openTag.length);
					const innerRows = fullMatch.slice(openIdx + openTag.length, closeIdx).trim();
					const afterRows = fullMatch.slice(closeIdx);

					if (before) {
						output.write(before);
					}

					output.write(beforeRows);

					const innerRowsMap = Utils.parseRows(innerRows);

					if (innerRows) {
						if (startRowNumber) {
							const filteredRows = Utils.getRowsBelow(innerRowsMap, startRowNumber);

							if (filteredRows) {
								output.write(filteredRows);
							}
						} else {
							output.write(innerRows);
						}
					}

					// new <row>
					const { dimension: newDimension, rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

					if (Utils.compareColumns(newDimension.maxColumn, dimension.maxColumn) > 0) {
						dimension.maxColumn = newDimension.maxColumn;
					}

					if (newDimension.maxRow > dimension.maxRow) {
						dimension.maxRow = newDimension.maxRow;
					}

					if (innerRows) {
						const filteredRows = Utils.getRowsAbove(innerRowsMap, actualRowNumber);

						if (filteredRows) {
							output.write(filteredRows);
						}
					}

					output.write(afterRows);

					if (after) {
						output.write(after);
					}

					inserted = true;

					continue;
				}

				// Case 2: <sheetData/>
				if (!inserted && /<sheetData\s*\/>/.test(line)) {
					const maxRowNumber = startRowNumber ?? Utils.getMaxRowNumber(line);

					const fullMatch = line.match(/<sheetData\s*\/>/)?.[0] || "";
					const matchIndex = line.indexOf(fullMatch);

					const before = line.slice(0, matchIndex);
					const after = line.slice(matchIndex + fullMatch.length);

					if (before) {
						output.write(before);
					}

					// Insert opening tag
					output.write("<sheetData>");

					// Prepare the rows
					const { dimension: newDimension } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

					if (Utils.compareColumns(newDimension.maxColumn, dimension.maxColumn) > 0) {
						dimension.maxColumn = newDimension.maxColumn;
					}

					if (newDimension.maxRow > dimension.maxRow) {
						dimension.maxRow = newDimension.maxRow;
					}

					// Insert closing tag
					output.write("</sheetData>");

					if (after) {
						output.write(after);
					}

					inserted = true;

					continue;
				}

				// Case 3: <sheetData>
				if (!inserted && /<sheetData[^>]*>/.test(line)) {
					isCollecting = true;

					collection += line;

					continue;
				}

				// After inserting rows, just copy the remaining lines
				output.write(line);
			}

			// Close the streams
			rl.close();
			output.end();

			// update dimension
			{
				const target = initialDimension;
				const refRange = `${dimension.minColumn}${dimension.minRow}:${dimension.maxColumn}${dimension.maxRow}`;
				const replacement = `<dimension ref="${refRange}"`;

				let buffer = "";
				let replaced = false;

				const input = fsSync.createReadStream(tempPath, { encoding: "utf8" });
				const output = fsSync.createWriteStream(fullPath);

				await new Promise((resolve, reject) => {
					input.on("data", chunk => {
						buffer += chunk;

						if (!replaced) {
							const index = buffer.indexOf(target);
							if (index !== -1) {
								// Заменяем только первое вхождение
								buffer = buffer.replace(target, replacement);
								replaced = true;
							}
						}

						output.write(buffer);
						buffer = ""; // очищаем, т.к. мы уже записали
					});

					input.on("error", reject);
					output.on("error", reject);

					input.on("end", () => {
						// на всякий случай дописываем остаток
						if (buffer) {
							output.write(buffer);
						}

						output.end();

						resolve(true);
					});
				});
			}

			await fs.unlink(tempPath);
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
			// Guarantee integrity
			await this.#validate();

			// Read the current files from the directory(in case they were changed manually)
			const updatedFiles = await this.#readAllFromDestination();

			const zipBuffer = await Zip.create(updatedFiles);

			await this.#cleanup();

			this.destroyed = true;

			return zipBuffer;

		} finally {
			this.#isProcessing = false;
		}
	}

	/**
	 * Writes the modified Excel template to a writable stream.
	 *
	 * @param {Writable} output - The writable stream to write to.
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async saveStream(output: Writable): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			// Guarantee integrity
			await this.#validate();

			await Zip.createWithStream(Array.from(this.fileKeys), this.destination, output);

			await this.#cleanup();

			this.destroyed = true;
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
	async set(key: string, content: Buffer | string): Promise<void> {
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
	 * Validates the template by checking all required files exist.
	 *
	 * @returns {Promise<void>}
	 * @throws {Error} If the template instance has been destroyed.
	 * @throws {Error} If any required files are missing.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async validate(): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			await this.#validate();
		} finally {
			this.#isProcessing = false;
		}
	}

	/** Static methods */

	/**
	 * Creates a Template instance from an Excel file source.
	 * Removes any existing files in the destination directory.
	 *
	 * @param {Object} data - The data to create the template from.
	 * @param {string} data.source - The path or buffer of the Excel file.
	 * @param {string} data.destination - The path to save the template to.
	 * @param {boolean} data.isUniqueDestination - Whether to add a random UUID to the destination path.
	 * @returns {Promise<Template>} A new Template instance.
	 * @throws {Error} If reading or writing files fails.
	 * @experimental This API is experimental and might change in future versions.
	 */
	static async from(data: {
		destination: string;
		source: string | Buffer;
		isUniqueDestination?: boolean;
	}): Promise<TemplateFs> {
		const { destination, isUniqueDestination = true, source } = data;

		if (!destination) {
			throw new Error("Destination is required");
		}

		// add random uuid to destination
		const destinationWithUuid = isUniqueDestination
			? path.join(destination, crypto.randomUUID())
			: destination;

		const buffer = typeof source === "string"
			? await fs.readFile(source)
			: source;

		const files = await Zip.read(buffer);

		// if destination exists, remove it
		await fs.rm(destinationWithUuid, { force: true, recursive: true });

		// Write all files to the file system, preserving exact paths
		await fs.mkdir(destinationWithUuid, { recursive: true });
		await Promise.all(
			Object.entries(files).map(async ([filePath, content]) => {
				const fullPath = path.join(destinationWithUuid, ...filePath.split("/"));
				await fs.mkdir(path.dirname(fullPath), { recursive: true });
				await fs.writeFile(fullPath, content);
			}),
		);

		return new TemplateFs(new Set(Object.keys(files)), destinationWithUuid);
	}
}
