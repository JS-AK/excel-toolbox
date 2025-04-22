import { describe, expect, it } from "vitest";

import { parseRows } from "./parse-rows.js";

describe("parseRows", () => {
	it("should return empty map for empty input", () => {
		expect(parseRows("")).toEqual(new Map());
	});

	it("should parse self-closing row tags", () => {
		const xml = `
      <row r="1"/>
      <row r="2" />
      <row r="3" custom="attr"/>
    `;
		const expected = new Map([
			[1, "<row r=\"1\"/>"],
			[2, "<row r=\"2\" />"],
			[3, "<row r=\"3\" custom=\"attr\"/>"],
		]);
		expect(parseRows(xml)).toEqual(expected);
	});

	it("should parse regular row tags with content", () => {
		const xml = `
      <row r="1"><c r="A1">Data</c></row>
      <row r="2" hidden="1">
        <c r="B2">More data</c>
      </row>
    `;
		const expected = new Map([
			[1, "<row r=\"1\"><c r=\"A1\">Data</c></row>"],
			[2, "<row r=\"2\" hidden=\"1\">\n        <c r=\"B2\">More data</c>\n      </row>"],
		]);
		expect(parseRows(xml)).toEqual(expected);
	});

	it("should handle mixed self-closing and regular tags", () => {
		const xml = `
      <row r="1"/>
      <row r="2"><c r="A2"/></row>
      <row r="3" custom="attr"></row>
    `;
		const expected = new Map([
			[1, "<row r=\"1\"/>"],
			[2, "<row r=\"2\"><c r=\"A2\"/></row>"],
			[3, "<row r=\"3\" custom=\"attr\"></row>"],
		]);
		expect(parseRows(xml)).toEqual(expected);
	});

	it("should handle complex row content", () => {
		const xml = `
      <row r="1">
        <c r="A1">
          <v>100</v>
        </c>
        <c r="B1">
          <v>200</v>
        </c>
      </row>
      <row r="2"><c r="A2"><v>Text</v></c></row>
    `;
		const expected = new Map([
			[1, "<row r=\"1\">\n        <c r=\"A1\">\n          <v>100</v>\n        </c>\n        <c r=\"B1\">\n          <v>200</v>\n        </c>\n      </row>"],
			[2, "<row r=\"2\"><c r=\"A2\"><v>Text</v></c></row>"],
		]);
		expect(parseRows(xml)).toEqual(expected);
	});

	it("should ignore non-row elements", () => {
		const xml = `
      <sheetData>
        <row r="1"/>
        <col min="1" max="10"/>
        <row r="2"/>
      </sheetData>
    `;
		const expected = new Map([
			[1, "<row r=\"1\"/>"],
			[2, "<row r=\"2\"/>"],
		]);
		expect(parseRows(xml)).toEqual(expected);
	});

	it("should handle large row numbers", () => {
		const xml = "<row r=\"1048576\"><c r=\"A1048576\"/></row>";
		const expected = new Map([
			[1048576, "<row r=\"1048576\"><c r=\"A1048576\"/></row>"],
		]);
		expect(parseRows(xml)).toEqual(expected);
	});
});
