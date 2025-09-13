import { beforeEach, describe, expect, it } from "vitest";

import { WorkbookBuilder } from "../workbook-builder.js";
import { addOrGet } from "./add-or-get.js";
import { remove } from "./remove.js";

import { CellStyle } from "../utils/sheet.js";

describe("remove()", () => {
	let wb: WorkbookBuilder;

	beforeEach(() => {
		wb = new WorkbookBuilder();
	});

	function makeStyle(): CellStyle {
		return {
			alignment: { horizontal: "center" },
			border: { bottom: { color: "FFFF0000", style: "medium" } },
			fill: { patternType: "solid" },
			font: { name: "Arial" },
			numberFormat: "dd/mm/yyyy",
		};
	}

	it("удаляет стиль по индексу", () => {
		const style = makeStyle();
		const idx = addOrGet.call(wb, { style });

		style.index = idx;

		expect(wb.fills).toHaveLength(2);
		expect(wb.fonts).toHaveLength(2);
		expect(wb.borders).toHaveLength(2);
		expect(wb.numFmts).toHaveLength(1);
		expect(wb.styleMap.size).toBe(1);
		expect(wb.cellXfs).toHaveLength(2);

		const removed = remove.call(wb, { style });

		expect(removed).toBe(true);
		expect(wb.cellXfs.findIndex(xf => xf.numFmtId === 164)).toBe(-1);
		expect([...wb.styleMap.values()]).not.toContain(idx);

		expect(wb.fills).toHaveLength(2); // остался только дефолт
		expect(wb.fonts).toHaveLength(2);
		expect(wb.borders).toHaveLength(2);
		expect(wb.numFmts).toHaveLength(1);
		expect(wb.styleMap.size).toBe(0);
		expect(wb.cellXfs).toHaveLength(1);
	});

	it("не удаляет части стиля, если они используются в другом xf", () => {
		const style1 = makeStyle();
		const idx1 = addOrGet.call(wb, { sheetName: "Sheet1", style: style1 });
		style1.index = idx1;

		const style2: CellStyle = { ...makeStyle(), font: { name: "Times New Roman" } };
		const idx2 = addOrGet.call(wb, { sheetName: "Sheet1", style: style2 });
		style2.index = idx2;

		// Удаляем первый стиль
		const removed = remove.call(wb, { sheetName: "Sheet1", style: style1 });
		expect(removed).toBe(true);

		// части всё ещё должны быть на месте, потому что второй стиль их использует
		expect(wb.fonts.length).toBeGreaterThan(1);
		expect(wb.fills.length).toBeGreaterThan(1);
	});

	it("возвращает false если ничего не было удалено", () => {
		const style = makeStyle();
		style.index = 999; // несуществующий индекс
		const removed = remove.call(wb, { sheetName: "Sheet1", style });
		expect(removed).toBe(false);
	});
});
