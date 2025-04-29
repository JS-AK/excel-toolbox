import { describe, expect, it } from "vitest";

import { shiftRowIndices } from "./shift-row-indices.js";

describe("shiftRowIndices", () => {
	it("shifts a single row and cell reference down", () => {
		const input = ["<row r=\"1\"><c r=\"A1\"/></row>"];
		const expected = ["<row r=\"3\"><c r=\"A3\"/></row>"];
		expect(shiftRowIndices(input, 2)).toEqual(expected);
	});

	it("shifts multiple rows and cells up", () => {
		const input = [
			"<row r=\"5\"><c r=\"B5\"/><c r=\"C5\"/></row>",
			"<row r=\"6\"><c r=\"A6\"/></row>",
		];
		const expected = [
			"<row r=\"3\"><c r=\"B3\"/><c r=\"C3\"/></row>",
			"<row r=\"4\"><c r=\"A4\"/></row>",
		];
		expect(shiftRowIndices(input, -2)).toEqual(expected);
	});

	it("returns original if offset is zero", () => {
		const input = ["<row r=\"2\"><c r=\"A2\"/></row>"];
		expect(shiftRowIndices(input, 0)).toEqual(input);
	});

	it("handles mixed columns and multi-digit row numbers", () => {
		const input = ["<row r=\"10\"><c r=\"Z10\"/><c r=\"AA10\"/></row>"];
		const expected = ["<row r=\"13\"><c r=\"Z13\"/><c r=\"AA13\"/></row>"];
		expect(shiftRowIndices(input, 3)).toEqual(expected);
	});

	it("does not modify unrelated attributes", () => {
		const input = ["<row r=\"7\" custom=\"yes\"><c r=\"C7\" style=\"bold\"/></row>"];
		const expected = ["<row r=\"9\" custom=\"yes\"><c r=\"C9\" style=\"bold\"/></row>"];
		expect(shiftRowIndices(input, 2)).toEqual(expected);
	});
});
