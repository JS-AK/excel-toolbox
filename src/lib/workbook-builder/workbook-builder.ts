import fs from "node:fs/promises";

import * as Utils from "./utils/index.js";
import * as Zip from "../zip/index.js";
import { updateDimension } from "../template/utils/update-dimension.js";

import * as Default from "./default/index.js";
import * as MergeCells from "./merge-cells/index.js";
import * as SharedStringRef from "./shared-string-ref/index.js";
import * as StyleRef from "./style-ref/index.js";
import { FILE_PATHS } from "./utils/constants.js";
import { columnIndexToLetter } from "../template/utils/column-index-to-letter.js";

export class WorkbookBuilder {
	// Нужна ли глубокая очистка
	#cleanupUnused: boolean;

	// Все что касается листов
	#files: Utils.ExcelFiles;

	// Все что касается листов
	#sheets: Map<string, Utils.SheetData> = new Map();

	// Все что касается shared strings
	#sharedStrings: string[] = [];
	#sharedStringMap: Map<string, number> = new Map(); // key = строка, value = индекс в массиве
	#sharedStringRefs: Map<string, Set<string>> = new Map(); // key = строка, value = множество листов

	// Все что касается styles
	#borders: NonNullable<Utils.XmlNode["children"]>;
	#cellXfs: Utils.CellXfs;
	#fills: NonNullable<Utils.XmlNode["children"]>;
	#fonts: NonNullable<Utils.XmlNode["children"]>;
	#numFmts: { formatCode: string; id: number }[];
	#styleMap = new Map<string, number>(); // JSON -> styleIndex

	// Все что касается merge cells
	#mergeCells: Map<string, MergeCells.MergeCell[]> = new Map();

	constructor({ cleanupUnused = false, defaultSheetName = Default.sheetName() } = {}) {
		this.#cleanupUnused = cleanupUnused;

		this.#files = Utils.initializeFiles(Default.sheetName());

		// Initial styles initialization
		this.#borders = [Default.border()];
		this.#fills = [Default.fill()];
		this.#fonts = [Default.font()];
		this.#numFmts = [];
		this.#cellXfs = [Default.cellXf()];

		const sheet = Utils.createSheet(defaultSheetName, {
			addMerge: this.#addMerge.bind(this),
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			cleanupUnused: this.#cleanupUnused,
			removeMerge: this.#removeMerge.bind(this),
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
			removeStyleRef: this.#removeStyleRef.bind(this),
		});

		this.#sheets.set(Default.sheetName(), sheet);
	}

	get sheets() {
		return this.#sheets;
	}

	get sharedStrings() {
		return this.#sharedStrings;
	}

	set sharedStrings(sharedStrings: string[]) {
		this.#sharedStrings = sharedStrings;
	}

	get sharedStringMap() {
		return this.#sharedStringMap;
	}

	get sharedStringRefs() {
		return this.#sharedStringRefs;
	}

	get borders() {
		return this.#borders;
	}

	get cellXfs() {
		return this.#cellXfs;
	}

	get fills() {
		return this.#fills;
	}

	get fonts() {
		return this.#fonts;
	}

	get numFmts() {
		return this.#numFmts;
	}

	get styleMap() {
		return this.#styleMap;
	}

	get mergeCells() {
		return this.#mergeCells;
	}

	/** Shared strings */

	#addSharedString(str: string, sheetName: string): number {
		return SharedStringRef.add.bind(this)({ sheetName, str });
	}

	#removeSharedStringRef(strIdx: number, sheetName: string): boolean {
		return SharedStringRef.remove.bind(this)({ sheetName, strIdx });
	}

	#removeSheetSharedStrings(sheetName: string) {
		return SharedStringRef.removeAllFromSheet.bind(this)({ sheetName });
	}

	/** -------------- */

	/** Style refs */

	#addOrGetStyle(style: Utils.CellStyle) {
		return StyleRef.addOrGet.bind(this)({ style });
	};

	#removeStyleRef(style: Utils.CellStyle): boolean {
		return StyleRef.remove.bind(this)({ style });
	}

	#removeSheetStyleRefs(sheetName: string): boolean {
		return StyleRef.removeAllFromSheet.bind(this)({ sheetName });
	}

	/** ---------- */

	/** Merge cells */

	#addMerge(payload: MergeCells.MergeCell & { sheetName: string }) {
		return MergeCells.add.bind(this)(payload);
	}

	#removeMerge(payload: MergeCells.MergeCell & { sheetName: string }) {
		return MergeCells.remove.bind(this)(payload);
	}

	#removeSheetMerges(sheetName: string) {
		this.mergeCells.delete(sheetName);
	}

	/** ----------- */

	#addFile(key: string, value: Utils.ExcelFileContent): void {
		this.#files[key] = value;
	}

	#updateAppXml() {
		this.#addFile(
			FILE_PATHS.APP,
			Utils.buildAppXml({ sheetNames: Array.from(this.#sheets.keys()) }),
		);
	}

	#updateWorkbookXml() {
		this.#addFile(
			FILE_PATHS.WORKBOOK,
			Utils.buildWorkbookXml(Array.from(this.#sheets.values())),
		);
	}

	#updateWorkbookRels() {
		this.#addFile(
			FILE_PATHS.WORKBOOK_RELS,
			Utils.buildWorkbookRels(this.#sheets.size),
		);
	}

	#updateContentTypes() {
		this.#addFile(
			FILE_PATHS.CONTENT_TYPES,
			Utils.buildContentTypesXml(this.#sheets.size),
		);
	}

	/** Public methods */

	addSheet(sheetName: string): Utils.SheetData {
		if (this.getSheet(sheetName)) {
			throw new Error("Sheet with this name already exists");
		}

		const sheet = Utils.createSheet(sheetName, {
			addMerge: this.#addMerge.bind(this),
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			cleanupUnused: this.#cleanupUnused,
			removeMerge: this.#removeMerge.bind(this),
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
			removeStyleRef: this.#removeStyleRef.bind(this),
		});

		this.#sheets.set(sheetName, sheet);

		// Добавляем запись в app.xml
		this.#updateAppXml();

		// Добавляем запись в workbook.xml
		this.#updateWorkbookXml();

		// Добавляем связь в workbook.xml.rels
		this.#updateWorkbookRels();

		// Добавляем Override в Content_Types.xml
		this.#updateContentTypes();

		return sheet;
	}

	getSheet(sheetName: string): Utils.SheetData | undefined {
		return this.#sheets.get(sheetName);
	}

	removeSheet(sheetName: string): boolean {
		const sheet = this.#sheets.get(sheetName);

		if (!sheet) {
			// Лист с таким именем не найден
			return false;
		}

		if (this.#cleanupUnused) {
			// Удаляем его shared strings
			this.#removeSheetSharedStrings(sheetName);

			// Удаляем его style refs
			this.#removeSheetStyleRefs(sheetName);
		}

		// Удаляем его merges
		this.#removeSheetMerges(sheetName);

		// Удаляем из коллекции
		this.#sheets.delete(sheetName);

		this.#updateAppXml();
		this.#updateWorkbookXml();
		this.#updateWorkbookRels();
		this.#updateContentTypes();

		return true;
	}

	getInfo() {
		return {
			sheetsNames: Array.from(this.#sheets.values()).map((sheet) => sheet.name),

			sharedStringMap: this.#sharedStringMap,
			sharedStringRefs: this.#sharedStringRefs,
			sharedStrings: this.#sharedStrings,

			styles: {
				borders: JSON.stringify(this.#borders),
				cellXfs: JSON.stringify(this.#cellXfs),
				fills: JSON.stringify(this.#fills),
				fonts: JSON.stringify(this.#fonts),
				numFmts: JSON.stringify(this.#numFmts),
				styleMap: this.#styleMap,
			},
		};
	}

	async saveToFile(path: string) {
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

			this.#addFile(FILE_PATHS.SHARED_STRINGS, xml);
		}

		// Styles
		this.#addFile(FILE_PATHS.STYLES, Utils.buildStylesXml({
			borders: this.#borders,
			cellXfs: this.#cellXfs,
			fills: this.#fills,
			fonts: this.#fonts,
			numFmts: this.#numFmts,
		}));

		const zipBuffer = await Zip.create(this.#files);

		await fs.writeFile(path, zipBuffer);
	}
}
