import { beforeEach, describe, expect, it } from "vitest";

import { WorkbookBuilder } from "../workbook-builder.js";
import { add } from "./add.js";

describe("add (optimized with Map)", () => {
	let wb: WorkbookBuilder;

	beforeEach(() => {
		wb = new WorkbookBuilder();
		wb.addSheet("Sheet2");
	});

	it("should add new shared string and return correct index", () => {
		const idx = add.call(wb, { sheetName: "Sheet1", str: "Hello World" });

		expect(idx).toBe(0);
		expect(wb.getInfo().sharedStrings).toHaveLength(1);
		expect(wb.getInfo().sharedStrings[0]).toBe("Hello World");
		expect(wb.getInfo().sharedStringMap.get("Hello World")).toBe(0);
		// expect(wb.getInfo().sharedStringRefs.get("Hello World")?.has("Sheet1")).toBe(true);
	});

	it("should return existing index for duplicate string", () => {
		const idx1 = add.call(wb, { sheetName: "Sheet1", str: "Hello World" });
		const idx2 = add.call(wb, { sheetName: "Sheet2", str: "Hello World" });

		expect(idx1).toBe(idx2);
		expect(wb.getInfo().sharedStrings).toHaveLength(1);
		expect(wb.getInfo().sharedStringMap.get("Hello World")).toBe(0);
		// expect(wb.getInfo().sharedStringRefs.get("Hello World")?.has("Sheet1")).toBe(true);
		// expect(wb.getInfo().sharedStringRefs.get("Hello World")?.has("Sheet2")).toBe(true);
	});

	it("should handle multiple different strings", () => {
		const idx1 = add.call(wb, { sheetName: "Sheet1", str: "Hello" });
		const idx2 = add.call(wb, { sheetName: "Sheet1", str: "World" });
		const idx3 = add.call(wb, { sheetName: "Sheet1", str: "Test" });

		expect(idx1).toBe(0);
		expect(idx2).toBe(1);
		expect(idx3).toBe(2);
		expect(wb.getInfo().sharedStrings).toHaveLength(3);
		expect(wb.getInfo().sharedStringMap.get("Hello")).toBe(0);
		expect(wb.getInfo().sharedStringMap.get("World")).toBe(1);
		expect(wb.getInfo().sharedStringMap.get("Test")).toBe(2);
	});

	it("should maintain Map consistency after operations", () => {
		// Add some strings
		add.call(wb, { sheetName: "Sheet1", str: "First" });
		add.call(wb, { sheetName: "Sheet1", str: "Second" });
		add.call(wb, { sheetName: "Sheet1", str: "Third" });

		// Verify Map is consistent with array
		expect(wb.getInfo().sharedStringMap.size).toBe(wb.getInfo().sharedStrings.length);
		for (let i = 0; i < wb.getInfo().sharedStrings.length; i++) {
			const str = wb.getInfo().sharedStrings[i];
			expect(wb.getInfo().sharedStringMap.get(str)).toBe(i);
		}
	});
});
