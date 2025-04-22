import { describe, expect, it } from "vitest";

import { getRowsBelow } from "./get-rows-below.js";

describe("getRowsBelow", () => {
	it("should return empty string for empty map", () => {
		const map = new Map<number, string>();
		expect(getRowsBelow(map, 5)).toBe("");
	});

	it("should return rows below maxRow", () => {
		const map = new Map([
			[1, "<row r=\"1\">Row 1</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
			[3, "<row r=\"3\">Row 3</row>"],
			[4, "<row r=\"4\">Row 4</row>"],
		]);
		expect(getRowsBelow(map, 3)).toBe("<row r=\"1\">Row 1</row><row r=\"2\">Row 2</row>");
	});

	it("should return all rows when maxRow is above last row", () => {
		const map = new Map([
			[1, "<row r=\"1\">Row 1</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
		]);
		expect(getRowsBelow(map, 10)).toBe("<row r=\"1\">Row 1</row><row r=\"2\">Row 2</row>");
	});

	it("should return empty string when maxRow is below first row", () => {
		const map = new Map([
			[5, "<row r=\"5\">Row 5</row>"],
			[6, "<row r=\"6\">Row 6</row>"],
		]);
		expect(getRowsBelow(map, 1)).toBe("");
	});

	it("should handle non-consecutive row numbers", () => {
		const map = new Map([
			[10, "<row r=\"10\">Row 10</row>"],
			[20, "<row r=\"20\">Row 20</row>"],
			[30, "<row r=\"30\">Row 30</row>"],
		]);
		expect(getRowsBelow(map, 25)).toBe("<row r=\"10\">Row 10</row><row r=\"20\">Row 20</row>");
	});

	it("should maintain original order of rows", () => {
		const map = new Map([
			[3, "<row r=\"3\">Row 3</row>"],
			[1, "<row r=\"1\">Row 1</row>"],
			[4, "<row r=\"4\">Row 4</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
		]);
		expect(getRowsBelow(map, 4)).toBe("<row r=\"3\">Row 3</row><row r=\"1\">Row 1</row><row r=\"2\">Row 2</row>");
	});

	it("should handle edge case where maxRow equals a row number", () => {
		const map = new Map([
			[1, "<row r=\"1\">Row 1</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
			[3, "<row r=\"3\">Row 3</row>"],
		]);
		expect(getRowsBelow(map, 3)).toBe("<row r=\"1\">Row 1</row><row r=\"2\">Row 2</row>");
	});

	it("should handle large row numbers", () => {
		const map = new Map([
			[1000000, "<row r=\"1000000\">Big row</row>"],
			[1000001, "<row r=\"1000001\">Bigger row</row>"],
		]);
		expect(getRowsBelow(map, 1000001)).toBe("<row r=\"1000000\">Big row</row>");
	});

	it("should handle empty row content", () => {
		const map = new Map([
			[1, ""],
			[2, "<row r=\"2\"></row>"],
			[3, ""],
		]);
		expect(getRowsBelow(map, 3)).toBe("<row r=\"2\"></row>");
	});

	it("should handle exact boundary condition", () => {
		const map = new Map([
			[1, "Row 1"],
			[2, "Row 2"],
			[3, "Row 3"],
		]);
		expect(getRowsBelow(map, 2)).toBe("Row 1");
	});
});
