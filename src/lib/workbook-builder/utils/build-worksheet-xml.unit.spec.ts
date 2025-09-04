import { describe, expect, it } from "vitest";

import { trimAndJoinMultiline } from "../../utils/trim-and-join-multiline.js";

import { CellData, RowData } from "./sheet.js";
import { buildWorksheetXml } from "./build-worksheet-xml.js";

describe("buildWorksheetXml", () => {
	it("should build empty worksheet XML", () => {
		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(), separator: "" });
		expect(result).toContain("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		expect(result).toContain("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
		expect(result).toContain("<dimension ref=\"A1:A1\"/>");
		expect(result).toContain("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>");
		expect(result).toContain("<sheetFormatPr defaultRowHeight=\"15\"/>");
		expect(result).toContain("<sheetData/>");
	});

	it("should build worksheet with rows and cells", () => {
		const rows = new Map<number, RowData>([
			[1, { cells: new Map<string, CellData>([["A", { type: "str", value: "Test" }]]) }],
			[2, { cells: new Map<string, CellData>([["B", { type: "n", value: 123 }]]) }],
		]);

		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(rows), separator: "" });
		expect(result).toContain("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>Test</v></c></row>");
		expect(result).toContain("<row r=\"2\"><c r=\"B2\" t=\"n\"><v>123</v></c></row>");
	});

	it("should handle different cell types correctly", () => {
		const rows = new Map<number, RowData>([
			[
				1,
				{
					cells: new Map<string, CellData>([
						["A", { type: "inlineStr", value: "Inline" }],
						["B", { type: "b", value: true }],
						["C", { type: "s", value: "Shared" }],
						["D", { type: "n", value: 42.5 }],
						["E", { type: "e", value: "Error" }],
					]),
				},
			],
		]);

		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(rows), separator: "" });
		expect(result).toContain("<c r=\"A1\" t=\"inlineStr\"><is><t>Inline</t></is></c>");
		expect(result).toContain("<c r=\"B1\" t=\"b\"><v>1</v></c>");
		expect(result).toContain("<c r=\"C1\" t=\"s\"><v>Shared</v></c>");
		expect(result).toContain("<c r=\"D1\" t=\"n\"><v>42.5</v></c>");
		expect(result).toContain("<c r=\"E1\" t=\"e\"><v>Error</v></c>");
	});

	it("should include merge cells when provided", () => {
		const merges = ["A1:B2", "C3:D4"];
		const result = buildWorksheetXml(new Map(), merges);
		expect(result).toContain("<mergeCells count=\"2\">");
		expect(result).toContain("<mergeCell ref=\"A1:B2\"/>");
		expect(result).toContain("<mergeCell ref=\"C3:D4\"/>");
	});

	it("should not include mergeCells element when no merges provided", () => {
		const result = buildWorksheetXml();
		expect(result).not.toContain("mergeCells");
	});

	it("should handle cells with styles", () => {
		const rows = new Map<number, RowData>([
			[1, { cells: new Map<string, CellData>([["A", { style: { index: 1 }, type: "str", value: "Styled" }]]) }],
		]);

		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(rows), separator: "" });
		expect(result).toContain("<c r=\"A1\" s=\"1\" t=\"str\"><v>Styled</v></c>");
	});

	it("should skip cells with undefined values", () => {
		const rows = new Map<number, RowData>([
			[1, { cells: new Map<string, CellData>([["A", { type: "str", value: undefined as unknown as string }]]) }],
		]);

		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(rows), separator: "" });
		expect(result).toContain("<row r=\"1\"><c r=\"A1\" t=\"str\"/></row>");
	});
});

describe("buildCellChildren", () => {
	it("should handle boolean values correctly", () => {
		const rows = new Map<number, RowData>([
			[1, { cells: new Map<string, CellData>([["A", { type: "b", value: false }]]) }],
		]);

		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(rows), separator: "" });
		expect(result).toContain("<c r=\"A1\" t=\"b\"><v>0</v></c>");
	});

	it("should handle numeric values correctly", () => {
		const rows = new Map<number, RowData>([
			[1, { cells: new Map<string, CellData>([["A", { type: "n", value: 3.14 }]]) }],
		]);

		const result = trimAndJoinMultiline({ inputString: buildWorksheetXml(rows), separator: "" });
		expect(result).toContain("<c r=\"A1\" t=\"n\"><v>3.14</v></c>");
	});
});
