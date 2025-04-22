import { describe, expect, it } from "vitest";

import { updateDimension } from "./update-dimension.js";

describe("updateDimension", () => {
	it("should return original XML when no cell references found", () => {
		const xml = "<worksheet><dimension ref=\"A1:Z100\"/></worksheet>";
		expect(updateDimension(xml)).toBe(xml);
	});

	it("should update dimension for single cell", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:Z100"/>
        <sheetData>
          <row r="1"><c r="B2"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = updateDimension(xml);
		expect(result).toContain("<dimension ref=\"B2:B2\"");
	});

	it("should calculate correct dimension for multiple cells", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:A1"/>
        <sheetData>
          <row r="1"><c r="A1"/><c r="C3"/></row>
          <row r="2"><c r="B2"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = updateDimension(xml);
		expect(result).toContain("<dimension ref=\"A1:C3\"");
	});

	it("should handle non-contiguous cells", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:A1"/>
        <sheetData>
          <row r="1"><c r="Z100"/></row>
          <row r="2"><c r="A1"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = updateDimension(xml);
		expect(result).toContain("<dimension ref=\"A1:Z100\"");
	});

	it("should handle large cell references", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:A1"/>
        <sheetData>
          <row r="1048576"><c r="XFD1048576"/></row>
          <row r="1"><c r="A1"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = updateDimension(xml);
		expect(result).toContain("<dimension ref=\"A1:XFD1048576\"");
	});

	it("should preserve other attributes in dimension tag", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:Z100" xmlns="test"/>
        <sheetData>
          <row r="1"><c r="B2"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = updateDimension(xml);
		expect(result).toContain("<dimension ref=\"B2:B2\" xmlns=\"test\"");
	});

	it("should handle multiple cell references in same element", () => {
		const xml = `
      <worksheet>
        <dimension ref="A1:A1"/>
        <sheetData>
          <row r="1"><c r="A1"/><c r="Z100"/></row>
        </sheetData>
      </worksheet>
    `;
		const result = updateDimension(xml);
		expect(result).toContain("<dimension ref=\"A1:Z100\"");
	});
});
