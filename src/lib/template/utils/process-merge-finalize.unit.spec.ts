import { describe, expect, it } from "vitest";

import { processMergeFinalize } from "./process-merge-finalize.js";

describe("processMergeFinalize", () => {
	const baseData = {
		initialMergeCells: [],
		lastIndex: 0,
		mergeCellMatches: [],
		resultRows: [],
		rowShift: 0,
		sharedStrings: ["<si><t>Test</t></si>"],
		sharedStringsHeader: "<?xml version=\"1.0\"?>",
		sheetMergeCells: [],
		sheetXml: "<worksheet><sheetData></sheetData></worksheet>",
	};

	it("should handle empty merge cells", () => {
		const result = processMergeFinalize(baseData);
		expect(result.sheet).toContain("<sheetData></sheetData>");
		expect(result.shared).toContain("<sst");
		expect(result.shared).toContain("Test");
	});

	it("should process merge cells with row shift", () => {
		const data = JSON.parse(JSON.stringify({
			...baseData,
			mergeCellMatches: [
				{ from: "A1", to: "B2" },
				{ from: "C3", to: "D4" },
			],
			rowShift: 5,
		}));
		const result = processMergeFinalize(data);
		expect(result.sheet).toContain("<mergeCell ref=\"A6:B7\"/>");
		expect(result.sheet).toContain("<mergeCell ref=\"C8:D9\"/>");
	});

	it("should skip merge cells below lastIndex", () => {
		const data = JSON.parse(JSON.stringify({
			...baseData,
			lastIndex: 3,
			mergeCellMatches: [
				{ from: "A1", to: "B2" },  // Should be skipped
				{ from: "C4", to: "D5" },   // Should be included
			],
			rowShift: 1,
		}));
		const result = processMergeFinalize(data);
		expect(result.sheet).not.toContain("A2:B3");
		expect(result.sheet).toContain("<mergeCell ref=\"C5:D6\"/>");
	});

	it("should preserve initial merge cells when no new ones", () => {
		const initialMerge = "<mergeCells count=\"1\"><mergeCell ref=\"X1:Y2\"/></mergeCells>";
		const data = JSON.parse(JSON.stringify({
			...baseData,
			initialMergeCells: [initialMerge],
		}));
		const result = processMergeFinalize(data);

		expect(result.sheet).toContain(initialMerge);
	});

	it("should combine initial and new merge cells", () => {
		const initialMerge = "<mergeCells count=\"1\"><mergeCell ref=\"X1:Y2\"/></mergeCells>";
		const data = JSON.parse(JSON.stringify({
			...baseData,
			initialMergeCells: [initialMerge],
			mergeCellMatches: [
				{ from: "A1", to: "B2" },
			],
			rowShift: 0,
		}));
		const result = processMergeFinalize(data);

		expect(result.sheet).toContain("A1:B2");
		expect(result.sheet).toContain("count=\"1\"");
	});

	it("should properly format shared strings XML", () => {
		const data = {
			...baseData,
			sharedStrings: [
				"<si><t>First</t></si>",
				"<si><t>Second</t></si>",
			],
		};
		const result = processMergeFinalize(data);
		expect(result.shared).toContain("count=\"2\" uniqueCount=\"2\"");
		expect(result.shared).toContain("<t>First</t>");
		expect(result.shared).toContain("<t>Second</t>");
	});

	it("should update sheet dimensions", () => {
		const data = JSON.parse(JSON.stringify({
			...baseData,
			sheetXml: `
        <worksheet>
          <dimension ref="A1:B2"/>
          <sheetData>
            <row r="1"><c r="A1"/><c r="B1"/></row>
            <row r="2"><c r="A2"/><c r="C2"/></row>
          </sheetData>
        </worksheet>
      `,
		}));
		const result = processMergeFinalize(data);
		expect(result.sheet).toContain("<dimension ref=\"A1:C2\"");
	});
});
