import fs from "node:fs/promises";

import * as Utils from "./utils/index.js";
import * as Zip from "../zip/index.js";
import { updateDimension } from "../template/utils/update-dimension.js";

import * as Default from "./default/index.js";
import * as SharedStringRef from "./shared-string-ref/index.js";
import * as StyleRef from "./style-ref/index.js";
import { FILE_PATHS } from "./utils/constants.js";

export type CellValue = string | number | Date;

export class WorkbookBuilder {
	#cleanupUnused: boolean;

	#files: Utils.ExcelFiles;
	#sheets: Map<string, Utils.SheetData> = new Map();
	#sharedStrings: string[] = [];
	#sharedStringRefs: Map<string, Set<string>> = new Map(); // key = строка, value = множество листов

	#borders: NonNullable<Utils.XmlNode["children"]>;
	#cellXfs: Utils.CellXfs;
	#fills: NonNullable<Utils.XmlNode["children"]>;
	#fonts: NonNullable<Utils.XmlNode["children"]>;
	#numFmts: { formatCode: string; id: number }[];
	#styleMap = new Map<string, number>(); // JSON -> styleIndex

	constructor({ cleanupUnused = false } = {}) {
		this.#cleanupUnused = cleanupUnused;

		this.#files = Utils.initializeFiles(Default.sheetName());

		// Initial styles initialization
		this.#borders = Default.borders();
		this.#fills = Default.fills();
		this.#fonts = Default.fonts();
		this.#numFmts = Default.numFmts();
		this.#cellXfs = Default.cellXfs();

		const sheet = Utils.createSheet(Default.sheetName(), {
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			cleanupUnused: this.#cleanupUnused,
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

	#addFile(key: string, value: Utils.ExcelFileContent): void {
		this.#files[key] = value;
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
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			cleanupUnused: this.#cleanupUnused,
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
			removeStyleRef: this.#removeStyleRef.bind(this),
		});

		this.#sheets.set(sheetName, sheet);

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
			this.#removeSheetSharedStrings(sheetName);
			this.#removeSheetStyleRefs(sheetName);
		}

		// Удаляем из коллекции
		this.#sheets.delete(sheetName);

		this.#updateWorkbookXml();
		this.#updateWorkbookRels();
		this.#updateContentTypes();

		return true;
	}

	getInfo() {
		return {
			sheetsNames: Array.from(this.#sheets.values()).map((sheet) => sheet.name),

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
		// 1. Пройтись по всем sheets
		Array.from(this.#sheets.values()).forEach((sheet, index) => {
			const xml = Utils.buildWorksheetXml(sheet.rows);
			const filePath = `xl/worksheets/sheet${index + 1}.xml`;

			this.#addFile(filePath, updateDimension(xml));
		});

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
