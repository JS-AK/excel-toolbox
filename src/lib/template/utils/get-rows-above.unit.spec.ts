import { describe, expect, it } from "vitest";

import { getRowsAbove } from "./get-rows-above.js";

describe("getRowsAbove", () => {
	it("should return empty string for empty map", () => {
		const map = new Map<number, string>();
		expect(getRowsAbove(map, 1)).toBe("");
	});

	it("should return rows above minRow", () => {
		const map = new Map([
			[1, "<row r=\"1\">Row 1</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
			[3, "<row r=\"3\">Row 3</row>"],
			[4, "<row r=\"4\">Row 4</row>"],
		]);
		expect(getRowsAbove(map, 2)).toBe("<row r=\"3\">Row 3</row><row r=\"4\">Row 4</row>");
	});

	it("should return all rows when minRow is below first row", () => {
		const map = new Map([
			[5, "<row r=\"5\">Row 5</row>"],
			[6, "<row r=\"6\">Row 6</row>"],
		]);
		expect(getRowsAbove(map, 0)).toBe("<row r=\"5\">Row 5</row><row r=\"6\">Row 6</row>");
	});

	it("should return empty string when minRow is above all rows", () => {
		const map = new Map([
			[1, "<row r=\"1\">Row 1</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
		]);
		expect(getRowsAbove(map, 5)).toBe("");
	});

	it("should handle non-consecutive row numbers", () => {
		const map = new Map([
			[10, "<row r=\"10\">Row 10</row>"],
			[20, "<row r=\"20\">Row 20</row>"],
			[30, "<row r=\"30\">Row 30</row>"],
		]);
		expect(getRowsAbove(map, 15)).toBe("<row r=\"20\">Row 20</row><row r=\"30\">Row 30</row>");
	});

	it("should maintain original order of rows", () => {
		const map = new Map([
			[3, "<row r=\"3\">Row 3</row>"],
			[1, "<row r=\"1\">Row 1</row>"], // Insertion order matters
			[4, "<row r=\"4\">Row 4</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
		]);
		expect(getRowsAbove(map, 2)).toBe("<row r=\"3\">Row 3</row><row r=\"4\">Row 4</row>");
	});

	it("should handle edge case where minRow equals a row number", () => {
		const map = new Map([
			[1, "<row r=\"1\">Row 1</row>"],
			[2, "<row r=\"2\">Row 2</row>"],
			[3, "<row r=\"3\">Row 3</row>"],
		]);
		expect(getRowsAbove(map, 2)).toBe("<row r=\"3\">Row 3</row>");
	});

	it("should handle large row numbers", () => {
		const map = new Map([
			[1000000, "<row r=\"1000000\">Big row</row>"],
			[1000001, "<row r=\"1000001\">Bigger row</row>"],
		]);
		expect(getRowsAbove(map, 1000000)).toBe("<row r=\"1000001\">Bigger row</row>");
	});

	it("should handle empty row content", () => {
		const map = new Map([
			[1, ""],
			[2, "<row r=\"2\"></row>"],
			[3, ""],
		]);
		expect(getRowsAbove(map, 1)).toBe("<row r=\"2\"></row>");
	});
});
