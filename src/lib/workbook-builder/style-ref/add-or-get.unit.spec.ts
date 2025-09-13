import { beforeEach, describe, expect, it } from "vitest";

import { WorkbookBuilder } from "../workbook-builder.js";
import { addOrGet } from "./add-or-get.js";

import { CellStyle } from "../utils/sheet.js";

describe("addOrGet()", () => {
	let wb: WorkbookBuilder;

	beforeEach(() => {
		wb = new WorkbookBuilder();
	});

	it("добавляет новый стиль и возвращает его индекс", () => {
		const style: CellStyle = {
			alignment: { horizontal: "center" },
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
			numberFormat: "dd/mm/yyyy",
		};

		const idx = addOrGet.call(wb, { sheetName: "Sheet1", style });

		expect(idx).toBe(1); // дефолтный стиль уже занимает 0

		expect(wb.cellXfs.at(-1)).toEqual({ alignment: { horizontal: "center" }, borderId: 1, fillId: 1, fontId: 1, numFmtId: 164 });
		expect(wb.fills.at(-1)).toEqual({ children: [{ attrs: { patternType: "solid" }, children: [], tag: "patternFill" }], tag: "fill" });
		expect(wb.fonts.at(-1)).toEqual({
			children: [
				{ attrs: { val: "11" }, tag: "sz" },
				{ attrs: { theme: "1" }, tag: "color" },
				{ attrs: { val: "Arial" }, tag: "name" },
			],
			tag: "font",
		});
		expect(wb.borders.at(-1)).toEqual({
			children: [
				{ tag: "left" },
				{ tag: "right" },
				{ tag: "top" },
				{ attrs: { style: "medium" }, children: [{ attrs: { rgb: "FFFF0000" }, tag: "color" }], tag: "bottom" },
			], tag: "border",
		});
		expect(wb.numFmts.at(-1)).toEqual({ formatCode: "dd/mm/yyyy", id: 164 });

		expect(wb.styleMap.get(JSON.stringify({ alignment: { horizontal: "center" }, borderId: 1, fillId: 1, fontId: 1, numFmtId: 164 }))).toBe(1);

		expect(wb.cellXfs).toHaveLength(2);
		expect(wb.fills).toHaveLength(2);
		expect(wb.fonts).toHaveLength(2);
		expect(wb.borders).toHaveLength(2);
		expect(wb.numFmts).toHaveLength(1);

		expect(wb.styleMap.size).toBe(1);
	});

	it("повторное добавление того же стиля возвращает тот же индекс", () => {
		const style: CellStyle = {
			alignment: { horizontal: "center" },
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
			numberFormat: "dd/mm/yyyy",
		};

		const idx1 = addOrGet.call(wb, { sheetName: "Sheet1", style });
		const idx2 = addOrGet.call(wb, { sheetName: "Sheet1", style });

		expect(idx1).toBe(idx2);
		expect(wb.cellXfs).toHaveLength(2);
		expect(wb.fills).toHaveLength(2);
		expect(wb.fonts).toHaveLength(2);
		expect(wb.borders).toHaveLength(2);
		expect(wb.numFmts).toHaveLength(1);
		expect(wb.styleMap.size).toBe(1);
	});

	it("разные стили создают разные индексы", () => {
		const style1: CellStyle = {
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
			numberFormat: "dd/mm/yyyy",
		};
		const style2: CellStyle = {
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Times New Roman" },
			numberFormat: "dd/mm/yyyy",
		};

		const idx1 = addOrGet.call(wb, { sheetName: "Sheet1", style: style1 });
		const idx2 = addOrGet.call(wb, { sheetName: "Sheet1", style: style2 });

		expect(idx1).toBe(1);
		expect(idx2).toBe(2);
		expect(wb.cellXfs).toHaveLength(3);
		expect(wb.fonts).toHaveLength(3);
		expect(wb.styleMap.size).toBe(2);
	});

	it("если numberFormat не задан, то numFmtId = 0", () => {
		const style: CellStyle = {
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
		};

		const idx = addOrGet.call(wb, { sheetName: "Sheet1", style });

		expect(idx).toBe(1);
		expect(wb.cellXfs[idx].numFmtId).toBe(0);
	});
});
