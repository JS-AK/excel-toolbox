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
		this.#ensureNotDestroyed();

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

		await this.set(sheetPath, updatedXml);
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
		this.#ensureNotDestroyed();

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
	}

	/**
	 * Saves the modified Excel template to a buffer.
	 *
	 * @returns {Promise<Buffer>} The modified Excel template as a buffer.
	 * @throws {Error} If the template instance has been destroyed.
	 * @experimental This API is experimental and might change in future versions.
	 */
	async save(): Promise<Buffer> {
		this.#ensureNotDestroyed();

		// Guarantee integrity
		await this.validate();

		// Read the current files from the directory(in case they were changed manually)
		const updatedFiles = await this.#readAllFromDestination();

		const zipBuffer = await Zip.create(updatedFiles);

		await this.#cleanup();

		this.destroyed = true;

		return zipBuffer;
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
		this.#ensureNotDestroyed();

		// Guarantee integrity
		await this.validate();

		await Zip.createWithStream(Array.from(this.fileKeys), this.destination, output);

		await this.#cleanup();

		this.destroyed = true;
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
		this.#ensureNotDestroyed();

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
	async validate(): Promise<void> {
		this.#ensureNotDestroyed();

		for (const key of this.fileKeys) {
			const fullPath = path.join(this.destination, ...key.split("/"));
			try {
				await fs.access(fullPath);
			} catch {
				throw new Error(`Missing file in template directory: ${key}`);
			}
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
