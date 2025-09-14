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

	it("validates existing style index and returns true", () => {
		const style = makeStyle();
		const idx = addOrGet.call(wb, { style });
		style.index = idx;

		const result = remove.call(wb, { style });

		expect(result).toBe(true);
	});

	it("throws error when styleIndex is not a number", () => {
		const style = makeStyle();
		// @ts-expect-error Testing invalid type
		style.index = "invalid";

		expect(() => {
			remove.call(wb, { style });
		}).toThrow("Invalid styleIndex");
	});

	it("throws error when styleIndex is undefined", () => {
		const style = makeStyle();
		style.index = undefined;

		expect(() => {
			remove.call(wb, { style });
		}).toThrow("Invalid styleIndex");
	});

	it("throws error when styleIndex is null", () => {
		const style = makeStyle();
		// @ts-expect-error Testing invalid type
		style.index = null;

		expect(() => {
			remove.call(wb, { style });
		}).toThrow("Invalid styleIndex");
	});

	it("throws error when style doesn't exist at the given index", () => {
		const style = makeStyle();
		style.index = 999; // non-existent index

		expect(() => {
			remove.call(wb, { style });
		}).toThrow("Invalid styleIndex");
	});

	it("throws error when trying to remove default style (index 0)", () => {
		const style = makeStyle();
		style.index = 0; // default style index

		expect(() => {
			remove.call(wb, { style });
		}).toThrow("Invalid styleIndex: 0 is the default style");
	});

	it("validates style at index 1 (first custom style)", () => {
		const style = makeStyle();
		const idx = addOrGet.call(wb, { style });
		style.index = idx;

		const result = remove.call(wb, { style });

		expect(result).toBe(true);
		expect(idx).toBe(1); // Should be index 1, not 0
	});
});
