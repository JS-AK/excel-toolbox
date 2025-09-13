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
			// Add strings to multiple sheets
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			add.call(wb, { sheetName: "Sheet1", str: "World" });
			add.call(wb, { sheetName: "Sheet2", str: "Hello" });
			add.call(wb, { sheetName: "Sheet2", str: "Test" });

			expect(wb.sharedStrings).toEqual(["Hello", "World", "Test"]);
			expect(wb.sharedStringRefs.get("Hello")?.size).toBe(2);
			expect(wb.sharedStringRefs.get("World")?.size).toBe(1);
			expect(wb.sharedStringRefs.get("Test")?.size).toBe(1);

			// Remove all from Sheet1
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			// Verify Sheet1 references are removed
			expect(wb.sharedStringRefs.get("Hello")?.has("Sheet1")).toBe(false);
			expect(wb.sharedStringRefs.get("Hello")?.has("Sheet2")).toBe(true);
			expect(wb.sharedStringRefs.get("World")?.has("Sheet1")).toBeUndefined();
			expect(wb.sharedStringRefs.get("Test")?.has("Sheet1")).toBe(false);

			// "World" should be completely removed (no references left)
			expect(wb.sharedStringRefs.has("World")).toBe(false);
			expect(wb.sharedStrings).toEqual(["Hello", "Test"]);
		});

		it("should handle empty sheet gracefully", () => {
			// Add strings to other sheets
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });

			// Remove from empty sheet
			removeAllFromSheet.call(wb, { sheetName: "EmptySheet" });

			// Should not affect existing strings
			expect(wb.sharedStrings).toEqual(["Hello"]);
			expect(wb.sharedStringRefs.get("Hello")?.size).toBe(1);
		});

		it("should remove all strings when sheet is the only user", () => {
			// Add strings only to Sheet1
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			add.call(wb, { sheetName: "Sheet1", str: "World" });
			add.call(wb, { sheetName: "Sheet1", str: "Test" });

			expect(wb.sharedStrings).toHaveLength(3);
			expect(wb.sharedStringMap.size).toBe(3);

			// Remove all from Sheet1
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			// All strings should be removed
			expect(wb.sharedStrings).toHaveLength(0);
			expect(wb.sharedStringMap.size).toBe(0);
			expect(wb.sharedStringRefs.size).toBe(0);
		});
	});

	describe("index reordering", () => {
		it("should correctly reorder indices after removing multiple strings", () => {
			// Add strings to multiple sheets
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet1", str: "B" });
			add.call(wb, { sheetName: "Sheet2", str: "A" }); // Keep A
			add.call(wb, { sheetName: "Sheet1", str: "C" });
			add.call(wb, { sheetName: "Sheet1", str: "D" });

			expect(wb.sharedStrings).toEqual(["A", "B", "C", "D"]);
			expect(wb.sharedStringMap.get("A")).toBe(0);
			expect(wb.sharedStringMap.get("B")).toBe(1);
			expect(wb.sharedStringMap.get("C")).toBe(2);
			expect(wb.sharedStringMap.get("D")).toBe(3);

			// Remove all from Sheet1 (should remove B, C, D, keep A)
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.sharedStrings).toEqual(["A"]);
			expect(wb.sharedStringMap.get("A")).toBe(0);
			expect(wb.sharedStringMap.has("B")).toBe(false);
			expect(wb.sharedStringMap.has("C")).toBe(false);
			expect(wb.sharedStringMap.has("D")).toBe(false);
		});

		it("should handle complex reordering scenario", () => {
			// Create complex scenario: A, B, C, D, E, F
			// Sheet1 uses: A, C, E
			// Sheet2 uses: B, D, F
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet2", str: "B" });
			add.call(wb, { sheetName: "Sheet1", str: "C" });
			add.call(wb, { sheetName: "Sheet2", str: "D" });
			add.call(wb, { sheetName: "Sheet1", str: "E" });
			add.call(wb, { sheetName: "Sheet2", str: "F" });

			expect(wb.sharedStrings).toEqual(["A", "B", "C", "D", "E", "F"]);

			// Remove all from Sheet1 (should remove A, C, E)
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			expect(wb.sharedStrings).toEqual(["B", "D", "F"]);
			expect(wb.sharedStringMap.get("B")).toBe(0);
			expect(wb.sharedStringMap.get("D")).toBe(1);
			expect(wb.sharedStringMap.get("F")).toBe(2);
		});
	});

	describe("Map consistency", () => {
		it("should maintain Map consistency after bulk removal", () => {
			// Add many strings
			const count = 100;
			for (let i = 0; i < count; i++) {
				add.call(wb, { sheetName: "Sheet1", str: `String${i}` });
			}

			// Add some to Sheet2 as well
			for (let i = 0; i < count; i += 2) {
				add.call(wb, { sheetName: "Sheet2", str: `String${i}` });
			}

			// Remove all from Sheet1
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			// Verify Map consistency
			expect(wb.sharedStringMap.size).toBe(wb.sharedStrings.length);
			for (let i = 0; i < wb.sharedStrings.length; i++) {
				const str = wb.sharedStrings[i];
				expect(wb.sharedStringMap.get(str)).toBe(i);
			}

			// Should have half the strings remaining (even indices)
			expect(wb.sharedStrings.length).toBe(Math.ceil(count / 2));
		});

		it("should handle multiple sheet removals correctly", () => {
			// Add strings to multiple sheets
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet2", str: "A" });
			add.call(wb, { sheetName: "Sheet3", str: "A" });
			add.call(wb, { sheetName: "Sheet1", str: "B" });
			add.call(wb, { sheetName: "Sheet2", str: "B" });

			expect(wb.sharedStrings).toEqual(["A", "B"]);
			expect(wb.sharedStringRefs.get("A")?.size).toBe(3);
			expect(wb.sharedStringRefs.get("B")?.size).toBe(2);

			// Remove all from Sheet1
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });
			expect(wb.sharedStrings).toEqual(["A", "B"]);
			expect(wb.sharedStringRefs.get("A")?.size).toBe(2);
			expect(wb.sharedStringRefs.get("B")?.size).toBe(1);

			// Remove all from Sheet2
			removeAllFromSheet.call(wb, { sheetName: "Sheet2" });
			expect(wb.sharedStrings).toEqual(["A"]);
			expect(wb.sharedStringRefs.get("A")?.size).toBe(1);
			expect(wb.sharedStringRefs.has("B")).toBe(false);

			// Remove all from Sheet3
			removeAllFromSheet.call(wb, { sheetName: "Sheet3" });
			expect(wb.sharedStrings).toHaveLength(0);
			expect(wb.sharedStringMap.size).toBe(0);
			expect(wb.sharedStringRefs.size).toBe(0);
		});
	});

	describe("performance", () => {
		it("should handle large number of strings efficiently", () => {
			// Add many strings
			const count = 5000;
			for (let i = 0; i < count; i++) {
				add.call(wb, { sheetName: "Sheet1", str: `String${i}` });
			}

			// Add some to Sheet2 as well
			for (let i = 0; i < count; i += 3) {
				add.call(wb, { sheetName: "Sheet2", str: `String${i}` });
			}

			const startTime = performance.now();
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });
			const endTime = performance.now();

			// Should be reasonably fast even with many strings
			expect(endTime - startTime).toBeLessThan(200); // Less than 200ms

			// Verify result
			expect(wb.sharedStrings.length).toBe(Math.ceil(count / 3));
			expect(wb.sharedStringMap.size).toBe(wb.sharedStrings.length);
		});
	});

	describe("edge cases", () => {
		it("should handle empty sharedStrings array", () => {
			// Remove from empty workbook
			removeAllFromSheet.call(wb, { sheetName: "Sheet1" });

			// Should not crash
			expect(wb.sharedStrings).toHaveLength(0);
			expect(wb.sharedStringMap.size).toBe(0);
			expect(wb.sharedStringRefs.size).toBe(0);
		});
	});
});
