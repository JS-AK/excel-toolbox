import fs from "node:fs/promises";

import * as Utils from "./utils/index.js";
import * as Zip from "../zip/index.js";
import { FILE_PATHS } from "./utils/constants.js";
import { updateDimension } from "../template/utils/update-dimension.js";

export type CellValue = string | number | Date;

export class WorkbookBuilder {
	#files: Utils.ExcelFiles;
	#sheets: Map<string, Utils.SheetData> = new Map();
	#sharedStrings: string[] = [];
	#sharedStringRefs: Map<string, Set<string>> = new Map(); // key = строка, value = множество листов

	constructor() {
		this.#files = Utils.initializeFiles();

		const sheet = Utils.createSheet("Sheet1", this.#addSharedString.bind(this), this.#removeSharedStringRef.bind(this));

		this.#sheets.set("Sheet1", sheet);
	}

	#addSharedString(str: string, sheetName: string): number {
		let idx = this.#sharedStrings.indexOf(str);

		if (idx === -1) {
			idx = this.#sharedStrings.length;
			this.#sharedStrings.push(str);
			this.#sharedStringRefs.set(str, new Set([sheetName]));
		} else {
			// Добавляем имя листа в Set, если ещё нет
			this.#sharedStringRefs.get(str)?.add(sheetName);
		}

		return idx;
	}

	#removeSheetSharedStrings(sheetName: string) {
		for (const [str, sheetsSet] of this.#sharedStringRefs) {
			sheetsSet.delete(sheetName);

			if (sheetsSet.size === 0) {
				// Удаляем строку из рефов и из массива sharedStrings
				this.#sharedStringRefs.delete(str);

				const idx = this.#sharedStrings.indexOf(str);

				if (idx !== -1) {
					this.#sharedStrings.splice(idx, 1);
				}
			}
		}

		// После удаления из массива sharedStrings нужна переиндексация
		this.#reindexSharedStrings();
	}

	#removeSharedStringRef(strIdx: number, sheetName: string): boolean {
		const str = this.#sharedStrings[strIdx];
		if (!str) return false;

		const refs = this.#sharedStringRefs.get(str);
		if (!refs) return false;

		refs.delete(sheetName);

		if (refs.size === 0) {
			// Удаляем строку
			this.#sharedStringRefs.delete(str);
			this.#sharedStrings.splice(strIdx, 1);
			this.#reindexSharedStrings(); // чтобы обновить индексы в ссылках и в ячейках всех листов!
		}

		return true;
	}

	#reindexSharedStrings() {
		// Создаем новую Map для быстрого поиска индексов
		const newSharedStrings = [...this.#sharedStrings];
		const newRefs = new Map<string, Set<string>>();

		for (const str of newSharedStrings) {
			const oldSet = this.#sharedStringRefs.get(str);

			if (oldSet) {
				newRefs.set(str, oldSet);
			} else {
				newRefs.set(str, new Set());
			}
		}

		this.#sharedStrings = newSharedStrings;
		this.#sharedStringRefs = newRefs;
	}

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

	addSheet(name: string): Utils.SheetData {
		const sheet = Utils.createSheet(name, this.#addSharedString.bind(this), this.#removeSharedStringRef.bind(this));

		this.#sheets.set(name, sheet);

		// Добавляем запись в workbook.xml
		this.#updateWorkbookXml();

		// Добавляем связь в workbook.xml.rels
		this.#updateWorkbookRels();

		// Добавляем Override в Content_Types.xml
		this.#updateContentTypes();

		return sheet;
	}

	getSheet(name: string): Utils.SheetData | undefined {
		return this.#sheets.get(name);
	}

	removeSheet(name: string): boolean {
		const sheet = this.#sheets.get(name);

		if (!sheet) {
			// Лист с таким именем не найден
			return false;
		}

		// Удаляем из коллекции
		this.#sheets.delete(name);

		this.#removeSheetSharedStrings(name);
		this.#updateWorkbookXml();
		this.#updateWorkbookRels();
		this.#updateContentTypes();

		return true;
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

		console.log(this.#files);

		const zipBuffer = await Zip.create(this.#files);

		await fs.writeFile(path, zipBuffer);
	}
}
