import { beforeEach, describe, expect, it } from "vitest";

import { WorkbookBuilder } from "../workbook-builder.js";
import { add } from "./add.js";
import { removeAllFromSheet } from "./remove-all-from-sheet.js";

describe("removeAllFromSheet (optimized with Map)", () => {
	let wb: WorkbookBuilder;

	beforeEach(() => {
		wb = new WorkbookBuilder();
		wb.addSheet("Sheet2");
		wb.addSheet("Sheet3");
	});

	describe("basic functionality", () => {
		it("should remove all references for a specific sheet", () => {
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			add.call(wb, { sheetName: "Sheet1", str: "World" });
			add.call(wb, { sheetName: "Sheet2", str: "Hello" });
			add.call(wb, { sheetName: "Sheet2", str: "Test" });

			expect(wb.getInfo().sharedStrings).toEqual(["Hello", "World", "Test"]);

			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.getInfo().sharedStrings).toEqual(["Hello", "World", "Test"]);
		});

		it("should handle empty sheet gracefully", () => {
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });

			removeAllFromSheet.call(wb, { sheetName: "EmptySheet" });

			expect(wb.getInfo().sharedStrings).toEqual(["Hello"]);
		});

		it("should remove all strings when sheet is the only user", () => {
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			add.call(wb, { sheetName: "Sheet1", str: "World" });
			add.call(wb, { sheetName: "Sheet1", str: "Test" });

			expect(wb.getInfo().sharedStrings).toHaveLength(3);
			expect(wb.getInfo().sharedStringMap.size).toBe(3);

			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.getInfo().sharedStrings).toHaveLength(3);
			expect(wb.getInfo().sharedStringMap.size).toBe(3);
		});
	});

	describe("index reordering", () => {
		it("should correctly reorder indices after removing multiple strings", () => {
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet1", str: "B" });
			add.call(wb, { sheetName: "Sheet2", str: "A" });
			add.call(wb, { sheetName: "Sheet1", str: "C" });
			add.call(wb, { sheetName: "Sheet1", str: "D" });

			expect(wb.getInfo().sharedStrings).toEqual(["A", "B", "C", "D"]);
			expect(wb.getInfo().sharedStringMap.get("A")).toBe(0);
			expect(wb.getInfo().sharedStringMap.get("B")).toBe(1);
			expect(wb.getInfo().sharedStringMap.get("C")).toBe(2);
			expect(wb.getInfo().sharedStringMap.get("D")).toBe(3);

			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.getInfo().sharedStrings).toEqual(["A", "B", "C", "D"]);
			expect(wb.getInfo().sharedStringMap.get("A")).toBe(0);
			expect(wb.getInfo().sharedStringMap.has("B")).toBe(true);
			expect(wb.getInfo().sharedStringMap.has("C")).toBe(true);
			expect(wb.getInfo().sharedStringMap.has("D")).toBe(true);
		});

		it("should handle complex reordering scenario", () => {
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet2", str: "B" });
			add.call(wb, { sheetName: "Sheet1", str: "C" });
			add.call(wb, { sheetName: "Sheet2", str: "D" });
			add.call(wb, { sheetName: "Sheet1", str: "E" });
			add.call(wb, { sheetName: "Sheet2", str: "F" });

			expect(wb.getInfo().sharedStrings).toEqual(["A", "B", "C", "D", "E", "F"]);

			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.getInfo().sharedStrings).toEqual(["A", "B", "C", "D", "E", "F"]);
			expect(wb.getInfo().sharedStringMap.get("B")).toBe(1);
			expect(wb.getInfo().sharedStringMap.get("D")).toBe(3);
			expect(wb.getInfo().sharedStringMap.get("F")).toBe(5);
		});
	});

	describe("Map consistency", () => {
		it("should maintain Map consistency after bulk removal", () => {
			const count = 100;
			for (let i = 0; i < count; i++) {
				add.call(wb, { sheetName: "Sheet1", str: `String${i}` });
			}

			for (let i = 0; i < count; i += 2) {
				add.call(wb, { sheetName: "Sheet2", str: `String${i}` });
			}

			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.getInfo().sharedStringMap.size).toBe(wb.getInfo().sharedStrings.length);
			for (let i = 0; i < wb.getInfo().sharedStrings.length; i++) {
				const str = wb.getInfo().sharedStrings[i];
				expect(wb.getInfo().sharedStringMap.get(str)).toBe(i);
			}

			expect(wb.getInfo().sharedStrings.length).toBe(100);
		});

		it("should handle multiple sheet removals correctly", () => {
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet2", str: "A" });
			add.call(wb, { sheetName: "Sheet3", str: "A" });
			add.call(wb, { sheetName: "Sheet1", str: "B" });
			add.call(wb, { sheetName: "Sheet2", str: "B" });

			expect(wb.getInfo().sharedStrings).toEqual(["A", "B"]);

			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });
			expect(wb.getInfo().sharedStrings).toEqual(["A", "B"]);

			removeAllFromSheet.call(wb, { sheetName: "Sheet2" });
			expect(wb.getInfo().sharedStrings).toEqual(["A", "B"]);

			removeAllFromSheet.call(wb, { sheetName: "Sheet3" });
			expect(wb.getInfo().sharedStrings).toHaveLength(2);
		});
	});

	describe("performance", () => {
		it("should handle large number of strings efficiently", () => {
			const count = 5000;
			for (let i = 0; i < count; i++) {
				add.call(wb, { sheetName: "Sheet1", str: `String${i}` });
			}

			for (let i = 0; i < count; i += 3) {
				add.call(wb, { sheetName: "Sheet2", str: `String${i}` });
			}

			const startTime = performance.now();
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });
			const endTime = performance.now();

			expect(endTime - startTime).toBeLessThan(200); // Less than 200ms

			expect(wb.getInfo().sharedStrings.length).toBe(5000);
			expect(wb.getInfo().sharedStringMap.size).toBe(wb.getInfo().sharedStrings.length);
		});
	});

	describe("edge cases", () => {
		it("should handle empty sharedStrings array", () => {
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.getInfo().sharedStrings).toHaveLength(0);
			expect(wb.getInfo().sharedStringMap.size).toBe(0);
		});
	});
});
