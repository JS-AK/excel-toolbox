import { beforeEach, describe, expect, it } from "vitest";

import { ProcessMergeFinalizeData, processMergeFinalize } from "./process-merge-finalize.js";

describe("processMergeFinalize", () => {
	let baseData: ProcessMergeFinalizeData;

	beforeEach(() => {
		// Reset the state before each test
		baseData = {
			initialMergeCells: [],
			lastIndex: 0,
			resultRows: [],
			sharedStrings: ["<si><t>Test</t></si>"],
			sharedStringsHeader: "<?xml version=\"1.0\"?>",
			sheetMergeCells: [],
			sheetXml: "<worksheet><sheetData></sheetData></worksheet>",
		};
	});

	it("should handle empty merge cells", () => {
		const result = processMergeFinalize(baseData);
		expect(result.sheet).toContain("<sheetData></sheetData>");
		expect(result.shared).toContain("<sst");
		expect(result.shared).toContain("Test");
	});

	it("should preserve initial merge cells when no new ones", () => {
		const initialMerge = "<mergeCells count=\"1\"><mergeCell ref=\"X1:Y2\"/></mergeCells>";
		const result = processMergeFinalize({
			...baseData,
			initialMergeCells: [initialMerge],
		});

		expect(result.sheet).toContain(initialMerge);
	});

	it("should properly format shared strings XML", () => {
		const result = processMergeFinalize({
			...baseData,
			sharedStrings: [
				"<si><t>First</t></si>",
				"<si><t>Second</t></si>",
			],
		});
		expect(result.shared).toContain("count=\"2\" uniqueCount=\"2\"");
		expect(result.shared).toContain("<t>First</t>");
		expect(result.shared).toContain("<t>Second</t>");
	});

	it("should update sheet dimensions", () => {
		const result = processMergeFinalize({
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
		});
		expect(result.sheet).toContain("<dimension ref=\"A1:C2\"");
	});
});
