import { beforeEach, describe, expect, it } from "vitest";

import { WorkbookBuilder } from "../workbook-builder.js";
import { add } from "./add.js";
import { remove } from "./remove.js";

describe("remove (optimized with Map)", () => {
	let wb: WorkbookBuilder;

	beforeEach(() => {
		wb = new WorkbookBuilder();
		wb.addSheet("Sheet2");
	});

	describe("basic functionality", () => {
		it("should return false for non-existent string index", () => {
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 999 });
			expect(result).toBe(false);
		});

		it("should return false for non-existent sheet reference", () => {
			// Add a string
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });

			// Try to remove from non-existent sheet
			const result = remove.call(wb, { sheetName: "NonExistentSheet", strIdx: 0 });
			expect(result).toBe(false);
		});

		it("should remove reference but keep string when multiple sheets use it", () => {
			// Add same string to multiple sheets
			const idx1 = add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			const idx2 = add.call(wb, { sheetName: "Sheet2", str: "Hello" });

			expect(idx1).toBe(idx2);
			expect(wb.sharedStrings).toHaveLength(1);
			expect(wb.sharedStringRefs.get("Hello")?.size).toBe(2);

			// Remove reference from one sheet
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });

			expect(result).toBe(true);
			expect(wb.sharedStrings).toHaveLength(1); // String still exists
			expect(wb.sharedStringRefs.get("Hello")?.size).toBe(1); // Only one reference left
			expect(wb.sharedStringRefs.get("Hello")?.has("Sheet2")).toBe(true);
			expect(wb.sharedStringRefs.get("Hello")?.has("Sheet1")).toBe(false);
		});

		it("should completely remove string when no references left", () => {
			// Add string
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			expect(wb.sharedStrings).toHaveLength(1);
			expect(wb.sharedStringMap.get("Hello")).toBe(0);

			// Remove reference
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });

			expect(result).toBe(true);
			expect(wb.sharedStrings).toHaveLength(0);
			expect(wb.sharedStringMap.has("Hello")).toBe(false);
			expect(wb.sharedStringRefs.has("Hello")).toBe(false);
		});
	});

	describe("index reordering", () => {
		it("should correctly reorder indices when removing middle element", () => {
			// Add multiple strings
			const idx1 = add.call(wb, { sheetName: "Sheet1", str: "First" });
			const idx2 = add.call(wb, { sheetName: "Sheet1", str: "Second" });
			const idx3 = add.call(wb, { sheetName: "Sheet1", str: "Third" });

			expect(idx1).toBe(0);
			expect(idx2).toBe(1);
			expect(idx3).toBe(2);
			expect(wb.sharedStrings).toEqual(["First", "Second", "Third"]);

			// Remove middle element
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 1 });

			expect(result).toBe(true);
			expect(wb.sharedStrings).toEqual(["First", "Third"]);
			expect(wb.sharedStringMap.get("First")).toBe(0);
			expect(wb.sharedStringMap.get("Third")).toBe(1);
			expect(wb.sharedStringMap.has("Second")).toBe(false);
		});

		it("should correctly reorder indices when removing first element", () => {
			// Add multiple strings
			add.call(wb, { sheetName: "Sheet1", str: "First" });
			add.call(wb, { sheetName: "Sheet1", str: "Second" });
			add.call(wb, { sheetName: "Sheet1", str: "Third" });

			// Remove first element
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });

			expect(result).toBe(true);
			expect(wb.sharedStrings).toEqual(["Second", "Third"]);
			expect(wb.sharedStringMap.get("Second")).toBe(0);
			expect(wb.sharedStringMap.get("Third")).toBe(1);
			expect(wb.sharedStringMap.has("First")).toBe(false);
		});

		it("should correctly reorder indices when removing last element", () => {
			// Add multiple strings
			add.call(wb, { sheetName: "Sheet1", str: "First" });
			add.call(wb, { sheetName: "Sheet1", str: "Second" });
			add.call(wb, { sheetName: "Sheet1", str: "Third" });

			// Remove last element
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 2 });

			expect(result).toBe(true);
			expect(wb.sharedStrings).toEqual(["First", "Second"]);
			expect(wb.sharedStringMap.get("First")).toBe(0);
			expect(wb.sharedStringMap.get("Second")).toBe(1);
			expect(wb.sharedStringMap.has("Third")).toBe(false);
		});
	});

	describe("Map consistency", () => {
		it("should maintain Map consistency after removal", () => {
			// Add multiple strings
			add.call(wb, { sheetName: "Sheet1", str: "A" });
			add.call(wb, { sheetName: "Sheet1", str: "B" });
			add.call(wb, { sheetName: "Sheet1", str: "C" });
			add.call(wb, { sheetName: "Sheet1", str: "D" });

			// Remove middle element
			remove.call(wb, { sheetName: "Sheet1", strIdx: 1 });

			// Verify Map consistency
			expect(wb.sharedStringMap.size).toBe(wb.sharedStrings.length);
			for (let i = 0; i < wb.sharedStrings.length; i++) {
				const str = wb.sharedStrings[i];
				expect(wb.sharedStringMap.get(str)).toBe(i);
			}
		});

		it("should maintain Map consistency after multiple removals", () => {
			// Add many strings
			const strings = ["A", "B", "C", "D", "E", "F", "G", "H"];
			for (const str of strings) {
				add.call(wb, { sheetName: "Sheet1", str });
			}

			// Remove several elements
			remove.call(wb, { sheetName: "Sheet1", strIdx: 1 }); // Remove "B"
			remove.call(wb, { sheetName: "Sheet1", strIdx: 2 }); // Remove "D" (was at index 3)
			remove.call(wb, { sheetName: "Sheet1", strIdx: 4 }); // Remove "G" (was at index 6)

			// Verify Map consistency
			expect(wb.sharedStringMap.size).toBe(wb.sharedStrings.length);
			for (let i = 0; i < wb.sharedStrings.length; i++) {
				const str = wb.sharedStrings[i];
				expect(wb.sharedStringMap.get(str)).toBe(i);
			}

			// Verify remaining strings
			expect(wb.sharedStrings).toEqual(["A", "C", "E", "F", "H"]);
		});
	});

	describe("edge cases", () => {
		it("should handle removing from empty workbook", () => {
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });
			expect(result).toBe(false);
		});

		it("should handle negative index", () => {
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: -1 });
			expect(result).toBe(false);
		});

		it("should handle removing same string multiple times", () => {
			// Add string to multiple sheets
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			add.call(wb, { sheetName: "Sheet2", str: "Hello" });

			// Remove from first sheet
			const result1 = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });
			expect(result1).toBe(true);
			expect(wb.sharedStrings).toHaveLength(1);

			// Try to remove from first sheet again (should fail)
			const result2 = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });
			expect(result2).toBe(false);
		});
	});

	describe("performance characteristics", () => {
		it("should handle large number of strings efficiently", () => {
			// Add many strings
			const count = 1000;
			for (let i = 0; i < count; i++) {
				add.call(wb, { sheetName: "Sheet1", str: `String${i}` });
			}

			expect(wb.sharedStrings).toHaveLength(count);
			expect(wb.sharedStringMap.size).toBe(count);

			// Remove middle element
			const startTime = performance.now();
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: Math.floor(count / 2) });
			const endTime = performance.now();

			expect(result).toBe(true);
			expect(wb.sharedStrings).toHaveLength(count - 1);
			expect(wb.sharedStringMap.size).toBe(count - 1);

			// Should be fast even with many strings (Map lookup is O(1))
			expect(endTime - startTime).toBeLessThan(100); // Less than 100ms
		});
	});
});
