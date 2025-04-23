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

	constructor(files: Record<string, Buffer>) {
		this.files = files;
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
	): { sheet: string; shared: string } {
		const {
			initialMergeCells,
			mergeCellMatches,
			modifiedXml,
		} = Utils.processMergeCells(sheetXml);

		const {
			sharedIndexMap,
			sharedStrings,
			sharedStringsHeader,
			sheetMergeCells,
		} = Utils.processSharedStrings(sharedStringsXml);

		const { lastIndex, resultRows, rowShift } = Utils.processRows({
			mergeCellMatches,
			replacements,
			sharedIndexMap,
			sharedStrings,
			sheetMergeCells,
			sheetXml: modifiedXml,
		});

		return Utils.processMergeFinalize({
			initialMergeCells,
			lastIndex,
			mergeCellMatches,
			resultRows,
			rowShift,
			sharedStrings,
			sharedStringsHeader,
			sheetMergeCells,
			sheetXml: modifiedXml,
		});
	}

	#getXml(fileKey: string): string {
		if (!this.files[fileKey]) {
			throw new Error(`${fileKey} not found`);
		}

		return Xml.extractXmlFromSheet(this.files[fileKey]);
	}

	/**
	 * Get the path of the sheet with the given name inside the workbook.
	 * @param sheetName The name of the sheet to find.
	 * @returns The path of the sheet inside the workbook.
	 * @throws {Error} If the sheet is not found.
	 */
	async #getSheetPath(sheetName: string): Promise<string> {
		// Read XML workbook to find sheet name and path
		const workbookXml = this.#getXml("xl/workbook.xml");
		const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name="${sheetName}"[^>]+r:id="([^"]+)"[^>]*/>`));

		if (!sheetMatch) throw new Error(`Sheet "${sheetName}" not found`);

		const rId = sheetMatch[1];

		const relsXml = this.#getXml("xl/_rels/workbook.xml.rels");
		const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"[^>]*/>`));

		if (!relMatch) throw new Error(`Relationship "${rId}" not found`);

		return "xl/" + relMatch[1]!.replace(/^\/?xl\//, "");
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

	async #substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void> {
		const sharedStringsPath = "xl/sharedStrings.xml";
		const sheetPath = await this.#getSheetPath(sheetName);

		let sharedStringsContent = "";
		let sheetContent = "";

		if (this.files[sharedStringsPath]) {
			sharedStringsContent = this.#getXml(sharedStringsPath);
		}

		if (this.files[sheetPath]) {
			sheetContent = this.#getXml(sheetPath);

			const TABLE_REGEX = /\$\{table:([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/g;

			const hasTablePlaceholders = TABLE_REGEX.test(sharedStringsContent) || TABLE_REGEX.test(sheetContent);

			if (hasTablePlaceholders) {
				const result = this.#expandTableRows(sheetContent, sharedStringsContent, replacements);

				sheetContent = result.sheet;
				sharedStringsContent = result.shared;
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
			// Read workbook.xml and find the source sheet
			const workbookXmlPath = "xl/workbook.xml";

			const workbookXml = this.#getXml(workbookXmlPath);

			// Find the source sheet
			const sheetRegex = new RegExp(`<sheet[^>]+name="${sourceName}"[^>]+r:id="([^"]+)"[^>]*/>`);
			const sheetMatch = workbookXml.match(sheetRegex);
			if (!sheetMatch) throw new Error(`Sheet "${sourceName}" not found`);

			const sourceRId = sheetMatch[1];

			// Check if a sheet with the new name already exists
			if (new RegExp(`<sheet[^>]+name="${newName}"`).test(workbookXml)) {
				throw new Error(`Sheet "${newName}" already exists`);
			}

			// Read workbook.rels
			// Find the source sheet path by rId
			const relsXmlPath = "xl/_rels/workbook.xml.rels";
			const relsXml = this.#getXml(relsXmlPath);
			const relRegex = new RegExp(`<Relationship[^>]+Id="${sourceRId}"[^>]+Target="([^"]+)"[^>]*/>`);
			const relMatch = relsXml.match(relRegex);
			if (!relMatch) throw new Error(`Relationship "${sourceRId}" not found`);

			const sourceTarget = relMatch[1]; // sheetN.xml

			if (!sourceTarget) throw new Error(`Relationship "${sourceRId}" not found`);

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
			const contentTypesXml = this.#getXml(contentTypesPath);
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
	substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			return this.#substitute(sheetName, replacements);
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
			const workbookXml = this.#getXml("xl/workbook.xml");
			const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name="${sheetName}"[^>]+r:id="([^"]+)"[^>]*/>`));

			if (!sheetMatch || !sheetMatch[1]) {
				throw new Error(`Sheet "${sheetName}" not found`);
			}

			const rId = sheetMatch[1];
			const relsXml = this.#getXml("xl/_rels/workbook.xml.rels");
			const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"[^>]*/>`));

			if (!relMatch || !relMatch[1]) {
				throw new Error(`Relationship "${rId}" not found`);
			}

			const sheetPath = "xl/" + relMatch[1].replace(/^\/?xl\//, "");
			const sheetXml = this.#getXml(sheetPath);

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

			await this.#set(sheetPath, Buffer.from(updatedXml));
		} finally {
			this.#isProcessing = false;
		}
	}

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

			const workbookXml = this.#getXml("xl/workbook.xml");
			const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name="${sheetName}"[^>]+r:id="([^"]+)"[^>]*/>`));
			if (!sheetMatch) throw new Error(`Sheet "${sheetName}" not found`);

			const rId = sheetMatch[1];
			const relsXml = this.#getXml("xl/_rels/workbook.xml.rels");
			const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"[^>]*/>`));
			if (!relMatch || !relMatch[1]) throw new Error(`Relationship "${rId}" not found`);

			const sheetPath = "xl/" + relMatch[1].replace(/^\/?xl\//, "");
			const sheetXml = this.#getXml(sheetPath);

			const output = new MemoryWriteStream();

			let inserted = false;

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

				const { rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

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
				await Utils.writeRowsToStream(output, rows, maxRowNumber);
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

				const { rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, Utils.getMaxRowNumber(innerRows));

				if (innerRows) {
					const filtered = Utils.getRowsAbove(innerRowsMap, actualRowNumber);
					if (filtered) output.write(filtered);
				}

				output.write(closeTag);
				output.write(afterRows);
				inserted = true;
			}

			if (!inserted) throw new Error("Failed to locate <sheetData> for insertion");

			// ← теперь мы не собираем строку, а собираем Buffer
			this.files[sheetPath] = output.toBuffer();

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

			// Очистка всех буферов
			for (const key in this.files) {
				if (this.files.hasOwnProperty(key)) {
					this.files[key] = Buffer.alloc(0); // Заменяем на пустой буфер
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
