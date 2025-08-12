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

	#borders: NonNullable<Utils.XmlNode["children"]>;
	#cellXfs: Utils.CellXfs;
	#fills: NonNullable<Utils.XmlNode["children"]>;
	#fonts: NonNullable<Utils.XmlNode["children"]>;
	#numFmts: { formatCode: string; id: number }[];
	#styleMap: Map<string, number> = new Map(); // JSON -> styleIndex

	constructor() {
		this.#files = Utils.initializeFiles();

		// Начальная инициализация стилей

		// Бордюры — базовый пустой бордер
		this.#borders = [
			{
				children: [
					{ tag: "left" },
					{ tag: "right" },
					{ tag: "top" },
					{ tag: "bottom" },
				],
				tag: "border",
			},
		];

		// Заливки — первый элемент пустой patternFill (обязательно!)
		this.#fills = [
			{
				children: [
					{
						attrs: { patternType: "none" },
						tag: "patternFill",
					},
				],
				tag: "fill",
			},
		];

		// Шрифты — пустой базовый шрифт (как в buildStylesXml)
		this.#fonts = [
			{
				children: [
					{ attrs: { val: "11" }, tag: "sz" },
					{ attrs: { theme: "1" }, tag: "color" },
					{ attrs: { val: "Calibri" }, tag: "name" },
				],
				tag: "font",
			},
		];

		// Форматы чисел — пустой (пока нет своих)
		this.#numFmts = [];

		// cellXfs — базовый стиль, ссылающийся на индексы выше:
		// borderId=0, fillId=0, fontId=0, numFmtId=0
		this.#cellXfs = [
			{
				borderId: 0,
				fillId: 0,
				fontId: 0,
				numFmtId: 0,
			},
		];

		const sheet = Utils.createSheet("Sheet1", {
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
		});

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

	#addUnique(arr: unknown[], item: unknown): number {
		const idx = arr.findIndex(x => JSON.stringify(x) === JSON.stringify(item));

		if (idx !== -1) return idx;

		arr.push(item);

		return arr.length - 1;
	}

	#addNumFmt(formatCode: string) {
		// 164+ зарезервировано для кастомных форматов
		const existing = this.#numFmts.find(nf => nf.formatCode === formatCode);

		if (existing) return existing.id;

		const id = 164 + this.#numFmts.length;

		this.#numFmts.push({ formatCode, id });

		return id;
	}

	#fontToXml(font?: Utils.CellStyle["font"]): Utils.XmlNode {
		if (!font) return this.#fonts[0] as Utils.XmlNode;

		const children = [];
		if (font.size) children.push({
			attrs: { val: String(font.size) },
			tag: "sz",
		});
		if (font.color) {
			const colorVal = font.color.startsWith("#") ? font.color.slice(1) : font.color;
			if (colorVal.length === 6) {
				children.push({
					attrs: { rgb: "FF" + colorVal.toUpperCase() }, // добавляем FF - непрозрачность
					tag: "color",
				});
			} else if (colorVal.length === 8) {
				children.push({
					attrs: { rgb: colorVal.toUpperCase() },
					tag: "color",
				});
			} else {
				throw new Error(`Некорректный цвет: ${font.color}`);
			}
		}
		if (font.name) children.push({
			attrs: { val: font.name },
			tag: "name",
		});
		if (font.bold) children.push({ tag: "b" });
		if (font.italic) children.push({ tag: "i" });
		if (font.underline) {
			const val = font.underline === true ? "single" : font.underline;
			children.push({
				attrs: { val },
				tag: "u",
			});
		}

		return {
			children,
			tag: "font",
		};
	}

	#fillToXml(fill?: Utils.CellStyle["fill"]) {
		if (!fill) return this.#fills[0] as Utils.XmlNode;

		const patternType = fill.patternType ?? "solid";
		const children = [];

		const attrs: unknown = { patternType };
		const fillChildren = [];
		if (fill.fgColor) {
			const colorVal = fill.fgColor.startsWith("#") ? fill.fgColor.slice(1) : fill.fgColor;
			fillChildren.push({
				attrs: { rgb: colorVal },
				tag: "fgColor",
			});
		}
		if (fill.bgColor) {
			const colorVal = fill.bgColor.startsWith("#") ? fill.bgColor.slice(1) : fill.bgColor;
			fillChildren.push({
				attrs: { rgb: colorVal },
				tag: "bgColor",
			});
		}
		children.push({
			attrs,
			children: fillChildren,
			tag: "patternFill",
		});

		return {
			children,
			tag: "fill",
		};
	}

	#borderToXml(border?: Utils.CellStyle["border"]) {
		const children = [];
		for (const side of ["left", "right", "top", "bottom"] as const) {
			const b = border?.[side];
			if (b) {
				const attrs: unknown = { style: b.style };
				const sideChildren = b.color
					? [{
						attrs: { rgb: b.color.replace(/^#/, "") },
						tag: "color",
					}]
					: [];
				children.push({
					attrs,
					children: sideChildren,
					tag: side,
				});
			} else {
				children.push({ tag: side });
			}
		}
		return {
			children,
			tag: "border",
		};
	}

	#addOrGetStyle(style: Utils.CellStyle) {
		// Конвертируем каждую часть
		const fontId = this.#addUnique(this.#fonts, this.#fontToXml(style.font));
		const fillId = this.#addUnique(this.#fills, this.#fillToXml(style.fill));
		const borderId = this.#addUnique(this.#borders, this.#borderToXml(style.border));
		const numFmtId = style.numberFormat ? this.#addNumFmt(style.numberFormat) : 0;

		const xfKey = JSON.stringify({
			alignment: style.alignment ?? null,  // включаем alignment в ключ
			borderId,
			fillId,
			fontId,
			numFmtId,
		});

		if (this.#styleMap.has(xfKey)) {
			return this.#styleMap.get(xfKey)!;
		}

		const index = this.#cellXfs.length;

		this.#cellXfs.push({
			alignment: style.alignment,  // сохраняем alignment
			borderId,
			fillId,
			fontId,
			numFmtId,
		});

		this.#styleMap.set(xfKey, index);

		return index;
	};

	addSheet(name: string): Utils.SheetData {
		const sheet = Utils.createSheet(name, {
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
		});

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
