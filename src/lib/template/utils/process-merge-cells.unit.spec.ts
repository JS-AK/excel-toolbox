import { describe, expect, it } from "vitest";

import { processMergeCells } from "./process-merge-cells.js";

describe("processMergeCells", () => {
	it("should handle XML without merge cells", () => {
		const xml = "<worksheet><sheetData></sheetData></worksheet>";
		const result = processMergeCells(xml);

		expect(result.initialMergeCells).toEqual([]);
		expect(result.mergeCellMatches).toEqual([]);
		expect(result.modifiedXml).toBe(xml);
	});

	it("should extract single merge cell", () => {
		const xml = `
      <worksheet>
        <mergeCells count="1">
          <mergeCell ref="A1:B2"/>
        </mergeCells>
        <sheetData></sheetData>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		expect(result.initialMergeCells).toEqual([
			"<mergeCells count=\"1\">\n          <mergeCell ref=\"A1:B2\"/>\n        </mergeCells>",
		]);
		expect(result.mergeCellMatches).toEqual([
			{ from: "A1", to: "B2" },
		]);
		expect(result.modifiedXml).toContain("<sheetData></sheetData>");
		expect(result.modifiedXml).not.toContain("mergeCells");
	});

	it("should extract multiple merge cells", () => {
		const xml = `
      <worksheet>
        <mergeCells count="2">
          <mergeCell ref="A1:B2"/>
          <mergeCell ref="C3:D4"/>
        </mergeCells>
        <sheetData></sheetData>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		expect(result.mergeCellMatches).toEqual([
			{ from: "A1", to: "B2" },
			{ from: "C3", to: "D4" },
		]);
	});

	it("should handle merge cells with attributes", () => {
		const xml = `
      <worksheet>
        <mergeCells count="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <mergeCell ref="A1:B2"/>
        </mergeCells>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		expect(result.initialMergeCells[0]).toContain("xmlns=\"http://schemas.openxmlformats.org");
		expect(result.mergeCellMatches).toEqual([
			{ from: "A1", to: "B2" },
		]);
	});

	it("should handle multiple mergeCells blocks", () => {
		const xml = `
      <worksheet>
        <mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>
        <sheetData></sheetData>
        <mergeCells count="1"><mergeCell ref="C3:D4"/></mergeCells>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		// Only first mergeCells block is processed
		expect(result.initialMergeCells).toEqual([
			"<mergeCells count=\"1\"><mergeCell ref=\"A1:B2\"/></mergeCells>",
		]);
		expect(result.mergeCellMatches).toEqual([
			{ from: "A1", to: "B2" },
		]);
		expect(result.modifiedXml).toContain("<sheetData></sheetData>");
		expect(result.modifiedXml).toContain("<mergeCells count=\"1\"><mergeCell ref=\"C3:D4\"/></mergeCells>");
	});

	it("should handle malformed merge cell references", () => {
		const xml = `
      <worksheet>
        <mergeCells count="1">
          <mergeCell ref="A1:B2:invalid"/>
          <mergeCell ref="C3:D4"/>
        </mergeCells>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		expect(result.mergeCellMatches).toEqual([
			{ from: "C3", to: "D4" }, // Only valid ref is captured
		]);
	});

	it("should preserve XML outside mergeCells block", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:Z100"/>
        <mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>
        <sheetData>
          <row r="1"><c r="A1"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		expect(result.modifiedXml).toContain("<dimension ref=\"A1:Z100\"/>");
		expect(result.modifiedXml).toContain("<sheetData>");
		expect(result.modifiedXml).not.toContain("mergeCells");
	});

	it("should handle empty mergeCells block", () => {
		const xml = `
      <worksheet>
        <mergeCells count="0"></mergeCells>
      </worksheet>
    `;
		const result = processMergeCells(xml);

		expect(result.initialMergeCells).toEqual([
			"<mergeCells count=\"0\"></mergeCells>",
		]);
		expect(result.mergeCellMatches).toEqual([]);
	});
});
