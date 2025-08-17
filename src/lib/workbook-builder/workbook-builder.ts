import fs from "node:fs/promises";

import * as Utils from "./utils/index.js";
import * as Zip from "../zip/index.js";
import { FILE_PATHS } from "./utils/constants.js";
import { updateDimension } from "../template/utils/update-dimension.js";

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

	#bordersUsageMap = new Map<string, Map<string, number>>(); // border -> Set of pages
	#cellXfsUsageMap = new Map<string, Map<string, number>>(); // cellXf -> Set of pages
	#fillsUsageMap = new Map<string, Map<string, number>>(); // fill -> Set of pages
	#fontsUsageMap = new Map<string, Map<string, number>>(); // font -> Set of pages
	#numFmtsUsageMap = new Map<string, Map<string, number>>(); // numFmt -> Set of pages

	constructor({ cleanupUnused = false } = {}) {
		this.#cleanupUnused = cleanupUnused;

		this.#files = Utils.initializeFiles();

		// Initial styles initialization

		// Borders — basic empty border
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
			cleanupUnused: this.#cleanupUnused,
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
			removeStyleRef: this.#removeStyleRef.bind(this),
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
		// 1. Собираем строки, которые нужно удалить
		const stringsToRemove: string[] = [];

		for (const [str, sheetsSet] of this.#sharedStringRefs) {
			sheetsSet.delete(sheetName);
			if (sheetsSet.size === 0) {
				stringsToRemove.push(str);
			}
		}

		if (stringsToRemove.length === 0) return;

		// 2. Строим карту старых индексов → новых
		const oldToNew = new Map<number, number>();
		let newIdx = 0;

		for (let oldIdx = 0; oldIdx < this.#sharedStrings.length; oldIdx++) {
			const str = this.#sharedStrings[oldIdx];
			if (!str) continue; // пропускаем, если undefined
			if (stringsToRemove.includes(str)) {
				// Удаляем строку из рефов
				this.#sharedStringRefs.delete(str);
				continue; // индекс не учитывается
			}
			oldToNew.set(oldIdx, newIdx++);
		}

		// 3. Обновляем массив sharedStrings
		this.#sharedStrings = this.#sharedStrings.filter(s => !stringsToRemove.includes(s));

		// 4. Обновляем индексы в ячейках на всех листах
		for (const sheet of this.#sheets.values()) {
			for (const row of sheet.rows.values()) {
				for (const cell of row.cells.values()) {
					if (cell.type === "s" && typeof cell.value === "number") {
						const newIdx = oldToNew.get(cell.value);
						if (newIdx !== undefined) {
							cell.value = newIdx;
						} else {
							// Если cell.value была удалённой строкой, можно поставить 0 или null
							cell.value = 0;
						}
					}
				}
			}
		}
	}

	#removeSharedStringRef(strIdx: number, sheetName: string): boolean {
		const str = this.#sharedStrings[strIdx];
		if (!str) return false;

		const refs = this.#sharedStringRefs.get(str);
		if (!refs) return false;

		refs.delete(sheetName);

		if (refs.size === 0) {
			// Строим карту старых индексов → новых до удаления
			const oldToNew = new Map<number, number>();
			for (let i = 0; i < this.#sharedStrings.length; i++) {
				if (i < strIdx) oldToNew.set(i, i);
				else if (i > strIdx) oldToNew.set(i, i - 1);
				// i === strIdx — эта строка будет удалена, индекса нет
			}

			// Удаляем строку из массива и рефов
			this.#sharedStrings.splice(strIdx, 1);
			this.#sharedStringRefs.delete(str);

			// Обновляем индексы на всех листах
			for (const sheet of this.#sheets.values()) {
				for (const row of sheet.rows.values()) {
					for (const cell of row.cells.values()) {
						if (cell.type === "s" && typeof cell.value === "number") {
							const newIdx = oldToNew.get(cell.value);
							if (newIdx !== undefined) {
								cell.value = newIdx;
							} else {
								// На всякий случай, если cell.value был удалённой строкой
								cell.value = 0; // или null, по логике твоего приложения
							}
						}
					}
				}
			}
		}

		return true;
	}

	#removeStyleRef2(style: Utils.CellStyle, sheetName: string): boolean {
		const styleIndex = style.index;

		if (!styleIndex) {
			throw new Error("Invalid styleIndex");
		}

		let removedSomething = false;

		const fillId = this.#cellXfs[styleIndex]?.fillId;
		const fontId = this.#cellXfs[styleIndex]?.fontId;
		const borderId = this.#cellXfs[styleIndex]?.borderId;
		const numFmtId = this.#cellXfs[styleIndex]?.numFmtId;

		// Удаляем ссылку на cellXfs по индексу
		if (this.#removeFromUsageMap(this.#cellXfsUsageMap, sheetName, this.#cellXfs[styleIndex])) {
			this.#cellXfs.splice(styleIndex, 1);

			// Найдем и удалим из styleMap ключ для этого индекса
			for (const [key, idx] of this.#styleMap.entries()) {
				if (idx === styleIndex) {
					this.#styleMap.delete(key);
					break;
				}
			}

			// Переиндексация ячеек во всех листах
			for (const sheet of this.#sheets.values()) {
				for (const row of sheet.rows.values()) {
					for (const cell of row.cells.values()) {
						if (cell.style?.index !== undefined && cell.style.index > styleIndex) {
							cell.style.index -= 1;
						}
					}
				}
			}

			removedSomething = true;
		}

		// игнорим 0 в том числе
		if (fillId && this.#removeFromUsageMap(this.#fillsUsageMap, sheetName, style.fill)) {
			this.#fills.splice(fillId, 1);
			removedSomething = true;
		}

		// игнорим 0 в том числе
		if (fontId && this.#removeFromUsageMap(this.#fontsUsageMap, sheetName, style.font)) {
			this.#fonts.splice(fontId, 1);
			removedSomething = true;
		}

		// игнорим 0 в том числе
		if (borderId && this.#removeFromUsageMap(this.#bordersUsageMap, sheetName, style.border)) {
			this.#borders.splice(borderId, 1);
			removedSomething = true;
		}

		// не игнорим 0 в том числе
		if (numFmtId !== undefined) {
			const nf = this.#numFmts.find(nf => nf.id === numFmtId);
			if (nf && this.#removeFromUsageMap(this.#numFmtsUsageMap, sheetName, nf.formatCode)) {
				const idx = this.#numFmts.indexOf(nf);
				if (idx !== -1) {
					this.#numFmts.splice(idx, 1);
					removedSomething = true;
				}
			}
		}

		return removedSomething;
	}

	#reindexStyleMapAfterRemoval(removedIndex: number) {
		const updates: Array<[string, number]> = [];
		for (const [key, idx] of this.#styleMap.entries()) {
			if (idx === removedIndex) {
				this.#styleMap.delete(key);
			} else if (idx > removedIndex) {
				updates.push([key, idx - 1]);
			}
		}
		for (const [key, newIdx] of updates) {
			this.#styleMap.set(key, newIdx);
		}
	}

	#removeStyleRef(style: Utils.CellStyle, sheetName: string): boolean {
		const styleIndex = style.index;
		if (styleIndex === undefined || styleIndex === null) {
			throw new Error("Invalid styleIndex");
		}

		let removedSomething = false;

		// Снимем части до splice — индексы ещё валидны
		const xf = this.#cellXfs[styleIndex];
		const fillId = xf?.fillId;
		const fontId = xf?.fontId;
		const borderId = xf?.borderId;
		const numFmtId = xf?.numFmtId;

		// Удаляем сам xf (ключ в usageMap — это xf, а не cell.style!)
		if (xf && this.#removeFromUsageMap(this.#cellXfsUsageMap, sheetName, xf)) {
			this.#cellXfs.splice(styleIndex, 1);

			// почин: переиндексация styleMap после splice
			this.#reindexStyleMapAfterRemoval(styleIndex);

			// переиндексация ссылок в ячейках на всех листах
			for (const sheet of this.#sheets.values()) {
				for (const row of sheet.rows.values()) {
					for (const cell of row.cells.values()) {
						if (cell.style?.index !== undefined && cell.style.index > styleIndex) {
							cell.style.index -= 1;
						}
					}
				}
			}

			removedSomething = true;
		}

		// части стиля — удаляем только если больше нигде не используются
		if (fillId && this.#removeFromUsageMap(this.#fillsUsageMap, sheetName, style.fill)) {
			this.#fills.splice(fillId, 1);
			removedSomething = true;
		}
		if (fontId && this.#removeFromUsageMap(this.#fontsUsageMap, sheetName, style.font)) {
			this.#fonts.splice(fontId, 1);
			removedSomething = true;
		}
		if (borderId && this.#removeFromUsageMap(this.#bordersUsageMap, sheetName, style.border)) {
			this.#borders.splice(borderId, 1);
			removedSomething = true;
		}
		if (numFmtId !== undefined) {
			const nf = this.#numFmts.find(n => n.id === numFmtId);
			if (nf && this.#removeFromUsageMap(this.#numFmtsUsageMap, sheetName, nf.formatCode)) {
				const idx = this.#numFmts.indexOf(nf);
				if (idx !== -1) {
					this.#numFmts.splice(idx, 1);
					removedSomething = true;
				}
			}
		}

		return removedSomething;
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

	#fillToXml(fill?: Utils.CellStyle["fill"]): Utils.XmlNode {
		if (!fill) return this.#fills[0] as Utils.XmlNode;

		const patternType = fill.patternType ?? "solid";
		const children: Utils.XmlNode["children"] = [];

		const attrs = { patternType };
		const fillChildren: Utils.XmlNode["children"] = [];

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

	#borderToXml(border?: Utils.CellStyle["border"]): Utils.XmlNode {
		const children: Utils.XmlNode["children"] = [];

		for (const side of ["left", "right", "top", "bottom"] as const) {
			const b = border?.[side];
			if (b) {
				const attrs = { style: b.style };
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

	#addOrGetStyle(style: Utils.CellStyle, name: string) {
		// Конвертируем каждую часть
		const fontId = this.#addUnique(this.#fonts, this.#fontToXml(style.font));
		const fillId = this.#addUnique(this.#fills, this.#fillToXml(style.fill));
		const borderId = this.#addUnique(this.#borders, this.#borderToXml(style.border));
		const numFmtId = style.numberFormat ? this.#addNumFmt(style.numberFormat) : 0;

		const xf = {
			alignment: style.alignment,
			borderId,
			fillId,
			fontId,
			numFmtId,
		};

		const xfKey = JSON.stringify(xf);

		if (this.#styleMap.has(xfKey)) {
			return this.#styleMap.get(xfKey)!;
		}

		const index = this.#cellXfs.length;

		this.#cellXfs.push(xf);

		this.#styleMap.set(xfKey, index);

		this.#updateUsageMap(this.#cellXfsUsageMap, name, xf);
		this.#updateUsageMap(this.#fillsUsageMap, name, style.fill);
		this.#updateUsageMap(this.#fontsUsageMap, name, style.font);
		this.#updateUsageMap(this.#bordersUsageMap, name, style.border);
		this.#updateUsageMap(this.#numFmtsUsageMap, name, style.numberFormat);

		return index;
	};

	#updateUsageMap<T>(
		map: Map<string, Map<string, number>>,
		pageName: string,
		key?: T,
	) {
		if (key === undefined || key === null) return;

		const k = JSON.stringify(key);

		if (!map.has(k)) {
			map.set(k, new Map<string, number>());
		}

		const pageCounts = map.get(k)!;
		pageCounts.set(pageName, (pageCounts.get(pageName) ?? 0) + 1);
	}

	#removeFromUsageMap<T>(
		map: Map<string, Map<string, number>>,
		pageName: string,
		key?: T,
	): boolean {
		if (key === undefined || key === null) {
			return false;
		}

		const k = JSON.stringify(key);
		const pageCounts = map.get(k);
		if (!pageCounts) {
			return false;
		}

		const currentCount = pageCounts.get(pageName) ?? 0;
		if (currentCount <= 1) {
			pageCounts.delete(pageName);
		} else {
			pageCounts.set(pageName, currentCount - 1);
		}

		if (pageCounts.size === 0) {
			map.delete(k);
			return true; // стиль больше нигде не используется
		}

		return false; // ещё есть страницы, где используется
	}

	#removeStylesRefForSheet(sheetName: string): boolean {
		const sheet = this.#sheets.get(sheetName);
		if (!sheet) return false;

		const stylesToRemove: Utils.CellStyle[] = [];

		let removedSomething = false;

		for (const row of sheet.rows.values()) {
			for (const cell of row.cells.values()) {
				if (cell.style?.index !== undefined) {

					stylesToRemove.push(cell.style);
				}
			}
		}

		for (const style of stylesToRemove) {
			const removed = this.#removeStyleRef(style, sheetName);

			if (removed) {
				removedSomething = true;
			}
		}

		return removedSomething;
	}

	addSheet(name: string): Utils.SheetData {
		const sheet = Utils.createSheet(name, {
			addOrGetStyle: this.#addOrGetStyle.bind(this),
			addSharedString: this.#addSharedString.bind(this),
			cleanupUnused: this.#cleanupUnused,
			removeSharedStringRef: this.#removeSharedStringRef.bind(this),
			removeStyleRef: this.#removeStyleRef.bind(this),
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

		if (this.#cleanupUnused) {
			this.#removeSheetSharedStrings(name);
			this.#removeStylesRefForSheet(name);
		}

		// Удаляем из коллекции
		this.#sheets.delete(name);

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
