import { describe, expect, it } from "vitest";

import { getMaxRowNumber } from "./get-max-row-number.js";

describe("getMaxRowNumber", () => {
	it("should return 1 for empty string", () => {
		expect(getMaxRowNumber("")).toBe(1);
	});

	it("should return 1 when no rows are present", () => {
		expect(getMaxRowNumber("<sheetData></sheetData>")).toBe(1);
		expect(getMaxRowNumber("<row></row>")).toBe(1); // missing r attribute
	});

	it("should find the max row number from r attributes", () => {
		const xml = `
      <sheetData>
        <row r="1"></row>
        <row r="2"></row>
        <row r="3"></row>
      </sheetData>
    `;
		expect(getMaxRowNumber(xml)).toBe(4);
	});

	it("should handle non-consecutive row numbers", () => {
		const xml = `
      <sheetData>
        <row r="5"></row>
        <row r="2"></row>
        <row r="7"></row>
      </sheetData>
    `;
		expect(getMaxRowNumber(xml)).toBe(8);
	});

	it("should ignore rows without r attribute", () => {
		const xml = `
      <sheetData>
        <row></row>
        <row r="3"></row>
        <row></row>
        <row r="5"></row>
      </sheetData>
    `;
		expect(getMaxRowNumber(xml)).toBe(6);
	});

	it("should handle malformed row tags", () => {
		const xml = `
      <sheetData>
        <row r="1">
        <row r="2" >
        <row r="3" / >
        <row r="abc"></row> <!-- invalid number -->
        <row r="4"></row>
      </sheetData>
    `;
		expect(getMaxRowNumber(xml)).toBe(5);
	});

	it("should handle very large row numbers", () => {
		const xml = "<row r=\"1048576\"></row>"; // Excel max rows
		expect(getMaxRowNumber(xml)).toBe(1048577);
	});

	it("should handle multiple row declarations", () => {
		const xml = `
      <row r="1"></row>
      <row r="2"></row>
      <row r="3"></row>
      <row r="1"></row> <!-- duplicate -->
      <row r="5"></row>
    `;
		expect(getMaxRowNumber(xml)).toBe(6);
	});

	it("should ignore case in attribute names", () => {
		const xml = `
      <row R="1"></row>
      <row r="2"></row>
      <row R="3"></row>
    `;
		expect(getMaxRowNumber(xml)).toBe(4);
	});

	it("should handle self-closing row tags", () => {
		const xml = `
      <row r="1"/>
      <row r="2" />
      <row r="3"/>
    `;
		expect(getMaxRowNumber(xml)).toBe(4);
	});
});
