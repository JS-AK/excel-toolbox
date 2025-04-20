import * as fs from "node:fs/promises";
import * as fsSync from "node:fs";
import * as path from "node:path";
import * as readline from "node:readline";
import { Writable } from "node:stream";

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

	async copySheet(sourceName: string, newName: string): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			// Read workbook.xml and find the source sheet
			const workbookXmlPath = "xl/workbook.xml";
			const workbookXml = Xml.extractXmlFromSheet(await this.#readFile(workbookXmlPath));

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
			const relsXml = Xml.extractXmlFromSheet(await this.#readFile(relsXmlPath));
			const relRegex = new RegExp(`<Relationship[^>]+Id="${sourceRId}"[^>]+Target="([^"]+)"[^>]*/>`);
			const relMatch = relsXml.match(relRegex);
			if (!relMatch) throw new Error(`Relationship "${sourceRId}" not found`);

			const sourceTarget = relMatch[1]; // sheetN.xml

			if (!sourceTarget) throw new Error(`Relationship "${sourceRId}" not found`);

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
			const contentTypesPath = "[Content_Types].xml";
			const contentTypesXml = Xml.extractXmlFromSheet(await this.#readFile(contentTypesPath));
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

	substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void> {
		return this.#substitute(sheetName, replacements);
	}

	async #substitute(sheetName: string, replacements: Record<string, unknown>): Promise<void> {
		this.#ensureNotProcessing();
		this.#ensureNotDestroyed();

		this.#isProcessing = true;

		try {
			const sharedStringsPath = "xl/sharedStrings.xml";
			const sheetPath = `xl/worksheets/${sheetName}.xml`;

			let sharedStringsContent = "";
			let sheetContent = "";

			if (this.fileKeys.has(sharedStringsPath)) {
				sharedStringsContent = Xml.extractXmlFromSheet(await this.#readFile(sharedStringsPath));
			}

			if (this.fileKeys.has(sheetPath)) {
				sheetContent = Xml.extractXmlFromSheet(await this.#readFile(sheetPath));

				const TABLE_REGEX = /\$\{table:([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/g;

				const hasTablePlaceholders = TABLE_REGEX.test(sharedStringsContent) || TABLE_REGEX.test(sheetContent);

				if (hasTablePlaceholders) {
					const result = this.expandTableRows(sheetContent, sharedStringsContent, replacements);

					sheetContent = result.sheet;
					sharedStringsContent = result.shared;
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

		} finally {
			this.#isProcessing = false;
		}
	}

	expandTableRows(
		sheetXml: string,
		sharedStringsXml: string,
		replacements: Record<string, unknown>,
	): { sheet: string; shared: string } {
		const {
			initialMergeCells,
			mergeCellMatches,
			modifiedXml,
		} = processMergeCells(sheetXml);

		const {
			sharedIndexMap,
			sharedStrings,
			sharedStringsHeader,
			sheetMergeCells,
		} = processSharedStrings(sharedStringsXml);

		const { lastIndex, resultRows, rowShift } = processRows({
			mergeCellMatches,
			replacements,
			sharedIndexMap,
			sharedStrings,
			sheetMergeCells,
			sheetXml: modifiedXml,
		});

		return processBuild({
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

	resolveValue<T extends Record<string, unknown>>(obj: T, key: string): T | undefined {
		const parts = key.split(".");

		let current = obj;

		for (const part of parts) {
			if (current == null || typeof current !== "object" || Array.isArray(current)) {
				return undefined;
			}

			current = current[part] as T;
		}

		return current;
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
			const workbookXml = Xml.extractXmlFromSheet(await this.#readFile("xl/workbook.xml"));
			const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name="${sheetName}"[^>]+r:id="([^"]+)"[^>]*/>`));

			if (!sheetMatch || !sheetMatch[1]) {
				throw new Error(`Sheet "${sheetName}" not found`);
			}

			const rId = sheetMatch[1];
			const relsXml = Xml.extractXmlFromSheet(await this.#readFile("xl/_rels/workbook.xml.rels"));
			const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"[^>]*/>`));

			if (!relMatch || !relMatch[1]) {
				throw new Error(`Relationship "${rId}" not found`);
			}

			const sheetPath = "xl/" + relMatch[1].replace(/^\/?xl\//, "");
			const sheetXmlRaw = await this.#readFile(sheetPath);
			const sheetXml = Xml.extractXmlFromSheet(sheetXmlRaw);

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

			await this.#set(sheetPath, updatedXml);
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
			const workbookXml = Xml.extractXmlFromSheet(await this.#readFile("xl/workbook.xml"));
			const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name="${sheetName}"[^>]+r:id="([^"]+)"[^>]*/>`));

			if (!sheetMatch) throw new Error(`Sheet "${sheetName}" not found`);

			const rId = sheetMatch[1];
			const relsXml = Xml.extractXmlFromSheet(await this.#readFile("xl/_rels/workbook.xml.rels"));
			const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"[^>]*/>`));

			if (!relMatch) throw new Error(`Relationship "${rId}" not found`);

			// Path to the desired sheet (sheet1.xml)
			const sheetPath = "xl/" + relMatch[1]!.replace(/^\/?xl\//, "");

			// The temporary file for writing
			const fullPath = path.join(this.destination, ...sheetPath.split("/"));
			const tempPath = fullPath + ".tmp";

			// Streams for reading and writing
			const input = fsSync.createReadStream(fullPath, { encoding: "utf-8" });
			const output = fsSync.createWriteStream(tempPath, { encoding: "utf-8" });

			// Inserted rows flag
			let inserted = false;

			const rl = readline.createInterface({
				// Process all line breaks
				crlfDelay: Infinity,
				input,
			});

			let isCollecting = false;
			let collection = "";

			for await (const line of rl) {
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

						const { rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

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
					const { rowNumber: actualRowNumber } = await Utils.writeRowsToStream(output, rows, maxRowNumber);

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
					await Utils.writeRowsToStream(output, rows, maxRowNumber);

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

			// Move the temporary file to the original location
			await fs.rename(tempPath, fullPath);

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

	/**
	 * Creates a Template instance from an Excel file source.
	 * Removes any existing files in the destination directory.
	 *
	 * @param {Object} data - The data to create the template from.
	 * @param {string} data.source - The path or buffer of the Excel file.
	 * @param {string} data.destination - The path to save the template to.
	 * @returns {Promise<Template>} A new Template instance.
	 * @throws {Error} If reading or writing files fails.
	 * @experimental This API is experimental and might change in future versions.
	 */
	static async from(data: {
		destination: string;
		source: string | Buffer;
	}): Promise<TemplateFs> {
		const { destination, source } = data;

		if (!destination) {
			throw new Error("Destination is required");
		}

		const buffer = typeof source === "string"
			? await fs.readFile(source)
			: source;

		const files = await Zip.read(buffer);

		// if destination exists, remove it
		await fs.rm(destination, { force: true, recursive: true });

		// Write all files to the file system, preserving exact paths
		await fs.mkdir(destination, { recursive: true });
		await Promise.all(
			Object.entries(files).map(async ([filePath, content]) => {
				const fullPath = path.join(destination, ...filePath.split("/"));
				await fs.mkdir(path.dirname(fullPath), { recursive: true });
				await fs.writeFile(fullPath, content);
			}),
		);

		return new TemplateFs(new Set(Object.keys(files)), destination);
	}
}

function processMergeCells(sheetXml: string) {
	// Regular expression for finding <mergeCells> block
	const mergeCellsBlockRegex = /<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/;

	// Find the first <mergeCells> block (if there are multiple, in xlsx usually there is only one)
	const mergeCellsBlockMatch = sheetXml.match(mergeCellsBlockRegex);

	const initialMergeCells: string[] = [];
	const mergeCellMatches: { from: string; to: string }[] = [];

	if (mergeCellsBlockMatch) {
		const mergeCellsBlock = mergeCellsBlockMatch[0];
		initialMergeCells.push(mergeCellsBlock);

		// Extract <mergeCell ref="A1:B2"/> from this block
		const mergeCellRegex = /<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"\/>/g;
		for (const match of mergeCellsBlock.matchAll(mergeCellRegex)) {
			mergeCellMatches.push({ from: match[1]!, to: match[2]! });
		}
	}

	// Remove the <mergeCells> block from the XML
	const modifiedXml = sheetXml.replace(mergeCellsBlockRegex, "");

	return {
		initialMergeCells,
		mergeCellMatches,
		modifiedXml,
	};
};

function processSharedStrings(sharedStringsXml: string) {
	// Final list of merged cells with all changes
	const sheetMergeCells: string[] = [];

	// Array for storing shared strings
	const sharedStrings: string[] = [];
	const sharedStringsHeader = Utils.extractXmlDeclaration(sharedStringsXml);

	// Map for fast lookup of shared string index by content
	const sharedIndexMap = new Map<string, number>();

	// Regular expression for finding <si> elements (shared string items)
	const siRegex = /<si>([\s\S]*?)<\/si>/g;

	// Parse sharedStringsXml and fill sharedStrings and sharedIndexMap
	for (const match of sharedStringsXml.matchAll(siRegex)) {
		const content = match[1];

		if (!content) throw new Error("Shared index not found");

		const fullSi = `<si>${content}</si>`;
		sharedIndexMap.set(content, sharedStrings.length);
		sharedStrings.push(fullSi);
	}

	return {
		sharedIndexMap,
		sharedStrings,
		sharedStringsHeader,
		sheetMergeCells,
	};
};

function processRows(data: {
	replacements: Record<string, unknown>;
	sharedIndexMap: Map<string, number>;
	mergeCellMatches: { from: string; to: string }[];
	sharedStrings: string[];
	sheetMergeCells: string[];
	sheetXml: string;
}) {
	const {
		mergeCellMatches,
		replacements,
		sharedIndexMap,
		sharedStrings,
		sheetMergeCells,
		sheetXml,
	} = data;
	const TABLE_REGEX = /\$\{table:([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/g;

	// Array for storing resulting XML rows
	const resultRows: string[] = [];

	// Previous position of processed part of XML
	let lastIndex = 0;

	// Shift for row numbers
	let rowShift = 0;

	// Regular expression for finding <row> elements
	const rowRegex = /<row[^>]*?>[\s\S]*?<\/row>/g;

	// Process each <row> element
	for (const match of sheetXml.matchAll(rowRegex)) {
		// Full XML row
		const fullRow = match[0];

		// Start position of the row in XML
		const matchStart = match.index!;

		// End position of the row in XML
		const matchEnd = matchStart + fullRow.length;

		// Add the intermediate XML chunk (if any) between the previous and the current row
		if (lastIndex !== matchStart) {
			resultRows.push(sheetXml.slice(lastIndex, matchStart));
		}

		lastIndex = matchEnd;

		// Get row number from r attribute
		const originalRowNumber = parseInt(fullRow.match(/<row[^>]* r="(\d+)"/)?.[1] ?? "1", 10);

		// Update row number based on rowShift
		const shiftedRowNumber = originalRowNumber + rowShift;

		// Find shared string indexes in cells of the current row
		const sharedValueIndexes: number[] = [];

		// Regular expression for finding a cell
		const cellRegex = /<c[^>]*?r="([A-Z]+\d+)"[^>]*?>([\s\S]*?)<\/c>/g;

		for (const cell of fullRow.matchAll(cellRegex)) {
			const cellTag = cell[0];
			// Check if the cell is a shared string
			const isShared = /t="s"/.test(cellTag);
			const valueMatch = cellTag.match(/<v>(\d+)<\/v>/);

			if (isShared && valueMatch) {
				sharedValueIndexes.push(parseInt(valueMatch[1]!, 10));
			}
		}

		// Get the text content of shared strings by their indexes
		const sharedTexts = sharedValueIndexes.map(i => sharedStrings[i]?.replace(/<\/?si>/g, "") ?? "");

		// Find table placeholders in shared strings
		const tablePlaceholders = sharedTexts.flatMap(e => [...e.matchAll(TABLE_REGEX)]);

		// If there are no table placeholders, just shift the row
		if (tablePlaceholders.length === 0) {
			const updatedRow = fullRow
				.replace(/(<row[^>]* r=")(\d+)(")/, `$1${shiftedRowNumber}$3`)
				.replace(/<c r="([A-Z]+)(\d+)"/g, (_, col) => `<c r="${col}${shiftedRowNumber}"`);

			resultRows.push(updatedRow);

			// Update mergeCells for regular row with rowShift
			const calculatedRowNumber = originalRowNumber + rowShift;

			for (const { from, to } of mergeCellMatches) {
				const [, fromCol, fromRow] = from.match(/^([A-Z]+)(\d+)$/)!;
				const [, toCol] = to.match(/^([A-Z]+)(\d+)$/)!;

				if (Number(fromRow) === calculatedRowNumber) {
					const newFrom = `${fromCol}${shiftedRowNumber}`;
					const newTo = `${toCol}${shiftedRowNumber}`;

					sheetMergeCells.push(`<mergeCell ref="${newFrom}:${newTo}"/>`);
				}
			}

			continue;
		}

		// Get the table name from the first placeholder
		const firstMatch = tablePlaceholders[0];
		const tableName = firstMatch?.[1];
		if (!tableName) throw new Error("Table name not found");

		// Get data for replacement from replacements
		const array = replacements[tableName];
		if (!array) continue;
		if (!Array.isArray(array)) throw new Error("Table data is not an array");

		const tableRowStart = shiftedRowNumber;

		// Find mergeCells to duplicate (mergeCells that start with the current row)
		const mergeCellsToDuplicate = mergeCellMatches.filter(({ from }) => {
			const match = from.match(/^([A-Z]+)(\d+)$/);

			if (!match) return false;

			// Row number of the merge cell start position is in the second group
			const rowNumber = match[2];

			return Number(rowNumber) === tableRowStart;
		});

		// Change the current row to multiple rows from the data array
		for (let i = 0; i < array.length; i++) {
			const rowData = array[i];
			let newRow = fullRow;

			// Replace placeholders in shared strings with real data
			sharedValueIndexes.forEach((originalIdx, idx) => {
				const originalText = sharedTexts[idx];
				if (!originalText) throw new Error("Shared value not found");

				// Replace placeholders ${tableName.field} with real data from array data
				const replacedText = originalText.replace(TABLE_REGEX, (_, tbl, field) =>
					tbl === tableName ? String(rowData?.[field] ?? "") : "",
				);

				// Add new text to shared strings if it doesn't exist
				let newIndex: number;
				if (sharedIndexMap.has(replacedText)) {
					newIndex = sharedIndexMap.get(replacedText)!;
				} else {
					newIndex = sharedStrings.length;
					sharedIndexMap.set(replacedText, newIndex);
					sharedStrings.push(`<si>${replacedText}</si>`);
				}

				// Replace the shared string index in the cell
				newRow = newRow.replace(`<v>${originalIdx}</v>`, `<v>${newIndex}</v>`);
			});

			// Update row number and cell references
			const newRowNum = shiftedRowNumber + i;
			newRow = newRow
				.replace(/<row[^>]* r="\d+"/, rowTag => rowTag.replace(/r="\d+"/, `r="${newRowNum}"`))
				.replace(/<c r="([A-Z]+)\d+"/g, (_, col) => `<c r="${col}${newRowNum}"`);

			resultRows.push(newRow);

			// Add duplicate mergeCells for new rows
			for (const { from, to } of mergeCellsToDuplicate) {
				const [, colFrom, rowFrom] = from.match(/^([A-Z]+)(\d+)$/)!;
				const [, colTo, rowTo] = to.match(/^([A-Z]+)(\d+)$/)!;
				const newFrom = `${colFrom}${Number(rowFrom) + i}`;
				const newTo = `${colTo}${Number(rowTo) + i}`;

				sheetMergeCells.push(`<mergeCell ref="${newFrom}:${newTo}"/>`);
			}
		}

		// It increases the row shift by the number of added rows minus one replaced
		rowShift += array.length - 1;

		const delta = array.length - 1;

		const calculatedRowNumber = originalRowNumber + rowShift - array.length + 1;

		if (delta > 0) {
			for (const merge of mergeCellMatches) {
				const fromRow = parseInt(merge.from.match(/\d+$/)![0], 10);
				if (fromRow > calculatedRowNumber) {
					merge.from = merge.from.replace(/\d+$/, r => `${parseInt(r) + delta}`);
					merge.to = merge.to.replace(/\d+$/, r => `${parseInt(r) + delta}`);
				}
			}
		}
	}

	return { lastIndex, resultRows, rowShift };
};

function processBuild(data: {
	initialMergeCells: string[];
	lastIndex: number;
	mergeCellMatches: { from: string; to: string }[];
	resultRows: string[];
	rowShift: number;
	sharedStrings: string[];
	sharedStringsHeader: string | null;
	sheetMergeCells: string[];
	sheetXml: string;
}) {
	const {
		initialMergeCells,
		lastIndex,
		mergeCellMatches,
		resultRows,
		rowShift,
		sharedStrings,
		sharedStringsHeader,
		sheetMergeCells,
		sheetXml,
	} = data;

	for (const { from, to } of mergeCellMatches) {
		const [, fromCol, fromRow] = from.match(/^([A-Z]+)(\d+)$/)!;
		const [, toCol, toRow] = to.match(/^([A-Z]+)(\d+)$/)!;

		const fromRowNum = Number(fromRow);
		// These rows have already been processed, don't add duplicates
		if (fromRowNum <= lastIndex) continue;

		const newFrom = `${fromCol}${fromRowNum + rowShift}`;
		const newTo = `${toCol}${Number(toRow) + rowShift}`;

		sheetMergeCells.push(`<mergeCell ref="${newFrom}:${newTo}"/>`);
	}

	resultRows.push(sheetXml.slice(lastIndex));

	// Form XML for mergeCells if there are any
	const mergeXml = sheetMergeCells.length
		? `<mergeCells count="${sheetMergeCells.length}">${sheetMergeCells.join("")}</mergeCells>`
		: initialMergeCells;

	// Insert mergeCells before the closing sheetData tag
	const sheetWithMerge = resultRows.join("").replace(/<\/sheetData>/, `</sheetData>${mergeXml}`);

	// Return modified sheet XML and shared strings
	return {
		shared: `${sharedStringsHeader}\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStrings.join("")}</sst>`,
		sheet: Utils.updateDimension(sheetWithMerge),
	};
}
