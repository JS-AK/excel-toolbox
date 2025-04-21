import { beforeEach, describe, expect, it } from "vitest";

import { processRows } from "./process-rows.js";

describe("processRows", () => {
	let baseData: {
		replacements: Record<string, unknown>;
		sharedIndexMap: Map<string, number>;
		mergeCellMatches: { from: string; to: string }[];
		sharedStrings: string[];
		sheetMergeCells: string[];
		sheetXml: string;
	};

	beforeEach(() => {
		// Сбрасываем состояние перед каждым тестом
		baseData = {
			mergeCellMatches: [],
			replacements: {},
			sharedIndexMap: new Map(),
			sharedStrings: ["<si><t>Test</t></si>"],
			sheetMergeCells: [],
			sheetXml: "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>",
		};
	});

	it("should process basic rows without replacements", () => {
		const result = processRows({ ...baseData });
		expect(result.resultRows[1]).toContain("r=\"1\"");
		expect(result.rowShift).toBe(0);
	});

	it("should process table placeholders in shared strings", () => {
		const data = {
			...baseData,
			replacements: { users: [{ name: "Alice" }, { name: "Bob" }] },
			sharedStrings: ["<si><t>${table:users.name}</t></si>"],
		};

		const result = processRows(data);

		expect(result.resultRows.length).toBe(3);
		expect(result.rowShift).toBe(1); // 2 rows - 1 original
		expect(data.sharedStrings).toContain("<si><t>Alice</t></si>");
		expect(data.sharedStrings).toContain("<si><t>Bob</t></si>");
	});

	it("should update merge cells for expanded rows", () => {
		const data = {
			...baseData,
			mergeCellMatches: [{ from: "A1", to: "B1" }],
			replacements: { data: [{ value: "First" }, { value: "Second" }] },
			sharedStrings: ["<si><t>${table:data.value}</t></si>"],
			sheetXml: `
        <worksheet>
          <sheetData>
            <row r="1"><c r="A1" t="s"><v>0</v></c></row>
          </sheetData>
        </worksheet>
      `,
		};

		const result = processRows(data);

		expect(data.sheetMergeCells).toContain("<mergeCell ref=\"A1:B1\"/>");
		expect(data.sheetMergeCells).toContain("<mergeCell ref=\"A2:B2\"/>");
		expect(result.rowShift).toBe(1);
	});

	it("should skip processing if table not found in replacements", () => {
		const data = {
			...baseData,
			sharedStrings: ["<si><t>${table:missing.value}</t></si>"],
		};

		const result = processRows(data);

		expect(result.resultRows.length).toBe(1);
		expect(result.rowShift).toBe(0);
	});

	it("should handle multiple placeholders in one row", () => {
		const data = {
			...baseData,
			replacements: { users: [{ age: 30, name: "Alice" }, { age: 25, name: "Bob" }] },
			sharedStrings: ["<si><t>${table:users.name} - ${table:users.age}</t></si>"],
		};
		processRows(data);

		expect(data.sharedStrings).toContain("<si><t>Alice - 30</t></si>");
		expect(data.sharedStrings).toContain("<si><t>Bob - 25</t></si>");
	});

	it("should adjust merge cells after expanded rows", () => {
		const data = {
			...baseData,
			mergeCellMatches: [{ from: "A3", to: "B3" }],
			replacements: { data: [{ value: "First" }, { value: "Second" }] },
			sharedStrings: ["<si><t>${table:data.value}</t></si>"],
			sheetXml: `
        <worksheet>
          <sheetData>
            <row r="1"><c r="A1" t="s"><v>0</v></c></row>
            <row r="3"><c r="A3">Data</c></row>
          </sheetData>
        </worksheet>
      `,
		};

		const result = processRows(data);

		expect(data.sheetMergeCells).toContain("<mergeCell ref=\"A4:B4\"/>");
		expect(result.rowShift).toBe(1);
	});

	it("should throw error for non-array table data", () => {
		const data = {
			...baseData,
			replacements: { data: { value: "Not an array" } },
			sharedStrings: ["<si><t>${table:data.value}</t></si>"],
		};

		expect(() => processRows(data)).toThrow("Table data is not an array");
	});
});
