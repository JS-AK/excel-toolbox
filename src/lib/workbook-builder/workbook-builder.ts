import type { Writable } from "node:stream";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import * as Zip from "../zip/index.js";
import { columnIndexToLetter, updateDimension } from "../template/utils/index.js";

import * as Utils from "./utils/index.js";

import * as Types from "./types/index.js";

import * as Default from "./default/index.js";
import * as MergeCells from "./merge-cells/index.js";
import * as SharedStringRef from "./shared-string-ref/index.js";
import * as StyleRef from "./style-ref/index.js";

/**
 * Builds Excel workbooks by composing sheets, styles, shared strings and merges,
 * and provides methods to save to file or stream.
 *
 * @experimental This API is experimental and might change in future versions.
 */
export class WorkbookBuilder {
	/** In-memory representation of workbook files to be zipped. */
	#files: Utils.ExcelFiles;

	/** Collection of sheets keyed by sheet name. */
	#sheets: Map<string, Types.SheetData> = new Map();

	/** Shared strings storage used by cells of type "s". */
	#sharedStrings: string[] = [];
	/** Map for lookup of shared string indices (key = string, value = index). */
	#sharedStringMap: Map<string, number> = new Map();

	/** Workbook style collections. */
	#borders: NonNullable<Types.XmlNode["children"]>;
	#cellXfs: Types.CellXf[];
	#fills: NonNullable<Types.XmlNode["children"]>;
	#fonts: NonNullable<Types.XmlNode["children"]>;
	#numFmts: { formatCode: string; id: number }[];

	/** Map of serialized style JSON to style index (xf). */
	#styleMap = new Map<string, number>();

	/** Map caches for fast de-duplication of style components. */
	#fontMap = new Map<string, number>();
	#fillMap = new Map<string, number>();
	#borderMap = new Map<string, number>();

	/** Merge cell ranges grouped by sheet name. */
	#mergeCells: Map<string, Types.MergeCell[]> = new Map();

	/**
	 * Creates a new workbook with a default sheet and initial style collections.
	 *
	 * @param options.defaultSheetName - The name for the initial sheet
	 */
	constructor({
		defaultSheetName = Default.sheetName(),
	} = {}) {

		this.#files = Utils.initializeFiles(defaultSheetName);

		// Initialize base style collections
		this.#borders = [Default.border()];
		this.#fills = [Default.fill()];
		this.#fonts = [Default.font()];
		this.#numFmts = [];
		this.#cellXfs = [Default.cellXf()];

		// Seed component maps with defaults at index 0
		this.#fontMap.set(JSON.stringify(this.#fonts[0]), 0);
		this.#fillMap.set(JSON.stringify(this.#fills[0]), 0);
		this.#borderMap.set(JSON.stringify(this.#borders[0]), 0);

		const sheet = Utils.createSheet(defaultSheetName, {
			addMerge: this.#addMerge.bind(this),
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			removeMerge: this.#removeMerge.bind(this),
		});

		this.#sheets.set(Default.sheetName(), sheet);
	}

	/** Returns the internal sheets map. */
	protected get sheets() {
		return this.#sheets;
	}

	/** Returns the shared strings array. */
	protected get sharedStrings() {
		return this.#sharedStrings;
	}

	/** Replaces the shared strings array. */
	protected set sharedStrings(sharedStrings: string[]) {
		this.#sharedStrings = sharedStrings;
	}

	/** Returns the shared string index map. */
	protected get sharedStringMap() {
		return this.#sharedStringMap;
	}

	/** Returns the borders collection. */
	protected get borders() {
		return this.#borders;
	}

	/** Returns the border cache map (serialized xml -> index). */
	protected get bordersMap() {
		return this.#borderMap;
	}

	/** Returns the cellXfs (style records). */
	protected get cellXfs() {
		return this.#cellXfs;
	}

	/** Returns the fills collection. */
	protected get fills() {
		return this.#fills;
	}

	/** Returns the fill cache map (serialized xml -> index). */
	protected get fillsMap() {
		return this.#fillMap;
	}

	/** Returns the fonts collection. */
	protected get fonts() {
		return this.#fonts;
	}

	/** Returns the font cache map (serialized xml -> index). */
	protected get fontsMap() {
		return this.#fontMap;
	}

	/** Returns the number formats collection. */
	protected get numFmts() {
		return this.#numFmts;
	}

	/** Returns the mapping from serialized style JSON to style index. */
	protected get styleMap() {
		return this.#styleMap;
	}

	/** Returns the merge ranges, grouped by sheet name. */
	protected get mergeCells() {
		return this.#mergeCells;
	}

	/** Shared strings */

	/** Adds a shared string (or returns existing index) and tracks its usage by sheet. */
	#addSharedString(str: string, sheetName: string): number {
		return SharedStringRef.add.bind(this)({ sheetName, str });
	}

	/** -------------- */

	/** Style refs */

	/** Adds a style or returns an existing style index. */
	#addOrGetStyle(style: Types.CellStyle) {
		return StyleRef.addOrGet.bind(this)({ style });
	};

	/** ---------- */

	/** Merge cells */

	/** Adds a merge range to a sheet. */
	#addMerge(payload: Types.MergeCell & { sheetName: string }) {
		return MergeCells.add.bind(this)(payload);
	}

	/** Removes a merge range from a sheet. */
	#removeMerge(payload: Types.MergeCell & { sheetName: string }) {
		return MergeCells.remove.bind(this)(payload);
	}

	/** Removes all merge ranges for a sheet. */
	#removeSheetMerges(sheetName: string) {
		this.mergeCells.delete(sheetName);
	}

	/** ----------- */

	/** Adds or replaces a logical file content in the in-memory file map. */
	#addFile(key: string, value: Utils.ExcelFileContent): void {
		this.#files[key] = value;
	}

	/** Updates the docProps/app.xml content based on current sheet names. */
	#updateAppXml() {
		this.#addFile(
			Utils.FILE_PATHS.APP,
			Utils.buildAppXml({ sheetNames: Array.from(this.#sheets.keys()) }),
		);
	}

	/** Updates the xl/workbook.xml content based on current sheets. */
	#updateWorkbookXml() {
		this.#addFile(
			Utils.FILE_PATHS.WORKBOOK,
			Utils.buildWorkbookXml(Array.from(this.#sheets.values())),
		);
	}

	/** Updates the xl/_rels/workbook.xml.rels relationships for sheets. */
	#updateWorkbookRels() {
		this.#addFile(
			Utils.FILE_PATHS.WORKBOOK_RELS,
			Utils.buildWorkbookRels(this.#sheets.size),
		);
	}

	/** Updates [Content_Types].xml with sheet overrides. */
	#updateContentTypes() {
		this.#addFile(
			Utils.FILE_PATHS.CONTENT_TYPES,
			Utils.buildContentTypesXml(this.#sheets.size),
		);
	}

	/** Public methods */

	/**
	 * Adds a new sheet to the workbook.
	 *
	 * @throws Error if a sheet with the same name already exists
	 * @param sheetName - Sheet name to add
	 * @returns The created sheet data
	 */
	addSheet(sheetName: string): Types.SheetData {
		if (this.getSheet(sheetName)) {
			throw new Error("Sheet with this name already exists");
		}

		const sheet = Utils.createSheet(sheetName, {
			addMerge: this.#addMerge.bind(this),
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			removeMerge: this.#removeMerge.bind(this),
		});

		this.#sheets.set(sheetName, sheet);

		// Add entry to app.xml
		this.#updateAppXml();

		// Add entry to workbook.xml
		this.#updateWorkbookXml();

		// Add relationship in workbook.xml.rels
		this.#updateWorkbookRels();

		// Add Override in [Content_Types].xml
		this.#updateContentTypes();

		return sheet;
	}

	/** Returns a sheet by name if it exists. */
	getSheet(sheetName: string): Types.SheetData | undefined {
		return this.#sheets.get(sheetName);
	}

	/**
	 * Removes a sheet by name.
	 * If cleanup is enabled, also removes associated shared strings and styles.
	 *
	 * @param sheetName - Sheet name to remove
	 * @returns True if the sheet existed and was removed
	 */
	removeSheet(sheetName: string): true {
		const sheet = this.#sheets.get(sheetName);

		if (!sheet) {
			throw new Error("Sheet not found: " + sheetName);
		}

		// Remove its merges
		this.#removeSheetMerges(sheetName);

		// Remove from collection
		this.#sheets.delete(sheetName);

		this.#updateAppXml();
		this.#updateWorkbookXml();
		this.#updateWorkbookRels();
		this.#updateContentTypes();

		return true;
	}

	/**
	 * Returns a snapshot of the workbook internals for inspection and tests.
	 * The returned structure is deeply frozen to avoid accidental mutations.
	 */
	getInfo(): {
		mergeCells: Map<string, Types.MergeCell[]>;

		sheetsNames: string[];

		sharedStringMap: Map<string, number>;
		sharedStrings: string[];

		styles: {
			borders: NonNullable<Types.XmlNode["children"]>;
			cellXfs: Types.CellXf[];
			fills: NonNullable<Types.XmlNode["children"]>;
			fonts: NonNullable<Types.XmlNode["children"]>;
			numFmts: { formatCode: string; id: number }[];
			styleMap: Map<string, number>;
		};
	} {
		function deepFreeze<T>(obj: T): T {
			if (obj === null || obj === undefined) {
				return obj;
			}

			if (typeof obj !== "object") {
				// string | number | boolean | symbol
				return obj;
			}

			if (Array.isArray(obj)) {
				return Object.freeze(obj.map(item => deepFreeze(item))) as T;
			}

			if (obj instanceof Map) {
				const frozenMap = new Map(
					Array.from(obj.entries()).map(([k, v]) => [k, deepFreeze(v)]),
				);
				return Object.freeze(frozenMap) as T;
			}

			if (obj instanceof Set) {
				const frozenSet = new Set(Array.from(obj.values()).map(v => deepFreeze(v)));
				return Object.freeze(frozenSet) as T;
			}

			// XmlNode or generic object
			const frozenObj: Record<string, unknown> = {};
			for (const [k, v] of Object.entries(obj)) {
				frozenObj[k] = deepFreeze(v);
			}
			return Object.freeze(frozenObj) as T;
		}

		return deepFreeze({
			mergeCells: new Map(this.#mergeCells),

			sheetsNames: Array.from(this.#sheets.values()).map((sheet) => sheet.name),

			sharedStringMap: new Map(this.#sharedStringMap),
			sharedStrings: [...this.#sharedStrings],

			styles: {
				borders: [...this.#borders],
				cellXfs: [...this.#cellXfs],
				fills: [...this.#fills],
				fonts: [...this.#fonts],
				numFmts: [...this.#numFmts],
				styleMap: new Map(this.#styleMap),
			},
		});
	}

	/**
	 * Generates workbook XML parts in-memory and writes a .xlsx zip to disk.
	 *
	 * @param path - Absolute or relative file path to write
	 */
	async saveToFile(path: string): Promise<void> {
		let index = 0;

		for (const sheet of this.#sheets.values()) {
			const merges = this.#mergeCells.get(sheet.name) || [];
			const preparedMerges = merges.map(
				merge => `${columnIndexToLetter(merge.startCol)}${merge.startRow}:${columnIndexToLetter(merge.endCol)}${merge.endRow}`,
			);

			const xml = Utils.buildWorksheetXml(sheet.rows, preparedMerges);
			const filePath = `xl/worksheets/sheet${++index}.xml`;

			this.#addFile(filePath, updateDimension(xml));
		}

		if (this.#sharedStrings.length) {
			const xml = Utils.buildSharedStringsXml(this.#sharedStrings);

			this.#addFile(Utils.FILE_PATHS.SHARED_STRINGS, xml);
		}

		// Styles
		this.#addFile(Utils.FILE_PATHS.STYLES, Utils.buildStylesXml({
			borders: this.#borders,
			cellXfs: this.#cellXfs,
			fills: this.#fills,
			fonts: this.#fonts,
			numFmts: this.#numFmts,
		}));

		const zipBuffer = await Zip.create(this.#files);

		await fs.writeFile(path, zipBuffer);
	}

	/**
	 * Saves the workbook to a writable stream using temporary files.
	 * This method creates temporary files from the in-memory data and streams them
	 * to the output stream, avoiding loading the entire file into memory.
	 *
	 * @param output - Writable stream to receive the Excel file
	 * @param options.destination - Optional existing directory to use instead of a system temp directory
	 * @param options.cleanup - Optional flag to cleanup the temporary directory after saving (default: true)
	 * @returns Promise that resolves when the file has been fully written
	 */
	async saveToStream(
		output: Writable,
		options?: {
			destination?: string;
			cleanup?: boolean;
		},
	): Promise<void> {
		const { cleanup = true, destination } = options ?? {};

		// Determine a temp directory to assemble ZIP contents
		let tempDir = "";
		if (destination) {
			// Create a random subdirectory inside provided destination
			tempDir = path.join(destination, crypto.randomUUID());

			await fs.mkdir(tempDir, { recursive: true });
		} else {
			// Create a temp directory in OS temp
			tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "excel-toolbox-", crypto.randomUUID()));
		}

		let index = 0;

		const usedFileKeys: string[] = [];

		// Write "xl/worksheets/sheet*.xml"
		for (const sheet of this.#sheets.values()) {
			const merges = this.#mergeCells.get(sheet.name) || [];
			const filePath = `xl/worksheets/sheet${++index}.xml`;

			usedFileKeys.push(filePath);

			const fullPath = path.join(tempDir, ...filePath.split("/"));

			await Utils.writeWorksheetXml(fullPath, sheet.rows, merges);

			this.#addFile(filePath, "");
		}

		// Write "xl/sharedStrings.xml"
		if (this.#sharedStrings.length) {
			usedFileKeys.push(Utils.FILE_PATHS.SHARED_STRINGS);

			const fullPath = path.join(tempDir, ...Utils.FILE_PATHS.SHARED_STRINGS.split("/"));

			await Utils.writeSharedStringsXml(fullPath, this.#sharedStrings);
		}

		// Write "xl/styles.xml"
		{
			usedFileKeys.push(Utils.FILE_PATHS.STYLES);

			const fullPath = path.join(tempDir, ...Utils.FILE_PATHS.STYLES.split("/"));

			await Utils.writeStylesXml(fullPath, {
				borders: this.#borders,
				cellXfs: this.#cellXfs,
				fills: this.#fills,
				fonts: this.#fonts,
				numFmts: this.#numFmts,
			});
		}

		try {
			// Write all files from memory to temporary files
			const fileKeys: string[] = [];

			for (const [key, value] of Object.entries(this.#files)) {
				if (usedFileKeys.includes(key)) {
					fileKeys.push(key);

					continue;
				}

				const fullPath = path.join(tempDir, ...key.split("/"));

				await fs.mkdir(path.dirname(fullPath), { recursive: true });
				await fs.writeFile(fullPath, value);

				fileKeys.push(key);
			}

			// Create ZIP archive and stream to output
			await Zip.createWithStream(fileKeys, tempDir, output);
		} finally {
			// Clean up temporary files
			if (cleanup) {
				await fs.rm(tempDir, { force: true, recursive: true });
			}
		}
	}
}
