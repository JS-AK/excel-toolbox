import { beforeEach, describe, expect, it } from "vitest";

import type { CellStyle } from "../types/index.js";

import { WorkbookBuilder } from "../workbook-builder.js";
import { addOrGet } from "./add-or-get.js";

describe("addOrGet()", () => {
	let wb: WorkbookBuilder;

	beforeEach(() => {
		wb = new WorkbookBuilder();
	});

	it("adds a new style and returns its index", () => {
		const style: CellStyle = {
			alignment: { horizontal: "center" },
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
			numberFormat: "dd/mm/yyyy",
		};

		const idx = addOrGet.call(wb, { sheetName: "Sheet1", style });

		expect(idx).toBe(1); // default style occupies index 0

		expect(wb.getInfo().styles.cellXfs.at(-1)).toEqual({ alignment: { horizontal: "center" }, borderId: 1, fillId: 1, fontId: 1, numFmtId: 164 });
		expect(wb.getInfo().styles.fills.at(-1)).toEqual({ children: [{ attrs: { patternType: "solid" }, children: [], tag: "patternFill" }], tag: "fill" });
		expect(wb.getInfo().styles.fonts.at(-1)).toEqual({
			children: [
				{ attrs: { val: "11" }, tag: "sz" },
				{ attrs: { theme: "1" }, tag: "color" },
				{ attrs: { val: "Arial" }, tag: "name" },
			],
			tag: "font",
		});
		expect(wb.getInfo().styles.borders.at(-1)).toEqual({
			children: [
				{ tag: "left" },
				{ tag: "right" },
				{ tag: "top" },
				{ attrs: { style: "medium" }, children: [{ attrs: { rgb: "FFFF0000" }, tag: "color" }], tag: "bottom" },
			], tag: "border",
		});
		expect(wb.getInfo().styles.numFmts.at(-1)).toEqual({ formatCode: "dd/mm/yyyy", id: 164 });

		expect(wb.getInfo().styles.styleMap.get(JSON.stringify({ alignment: { horizontal: "center" }, borderId: 1, fillId: 1, fontId: 1, numFmtId: 164 }))).toBe(1);

		expect(wb.getInfo().styles.cellXfs).toHaveLength(2);
		expect(wb.getInfo().styles.fills).toHaveLength(2);
		expect(wb.getInfo().styles.fonts).toHaveLength(2);
		expect(wb.getInfo().styles.borders).toHaveLength(2);
		expect(wb.getInfo().styles.numFmts).toHaveLength(1);

		expect(wb.getInfo().styles.styleMap.size).toBe(1);
	});

	it("adding the same style again returns the same index", () => {
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
		expect(wb.getInfo().styles.cellXfs).toHaveLength(2);
		expect(wb.getInfo().styles.fills).toHaveLength(2);
		expect(wb.getInfo().styles.fonts).toHaveLength(2);
		expect(wb.getInfo().styles.borders).toHaveLength(2);
		expect(wb.getInfo().styles.numFmts).toHaveLength(1);
		expect(wb.getInfo().styles.styleMap.size).toBe(1);
	});

	it("different styles produce different indices", () => {
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
		expect(wb.getInfo().styles.cellXfs).toHaveLength(3);
		expect(wb.getInfo().styles.fonts).toHaveLength(3);
		expect(wb.getInfo().styles.styleMap.size).toBe(2);
	});

	it("if numberFormat is not set, then numFmtId = 0", () => {
		const style: CellStyle = {
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
		};

		const idx = addOrGet.call(wb, { sheetName: "Sheet1", style });

		expect(idx).toBe(1);
		expect(wb.getInfo().styles.cellXfs[idx].numFmtId).toBe(0);
	});
});
