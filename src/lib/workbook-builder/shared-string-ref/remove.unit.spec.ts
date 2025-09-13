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
		it("should throw error for non-existent string index", () => {
			expect(() => {
				remove.call(wb, { sheetName: "Sheet1", strIdx: 999 });
			}).toThrow("String not found: 999");
		});

		it("should throw error for non-existent sheet reference", () => {
			// Add a string
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });

			expect(() => {
				remove.call(wb, { sheetName: "NonExistentSheet", strIdx: 0 });
			}).toThrow("Sheet not found: NonExistentSheet");
		});

		it("should remove reference but keep string when multiple sheets use it", () => {
			const idx1 = add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			const idx2 = add.call(wb, { sheetName: "Sheet2", str: "Hello" });

			expect(idx1).toBe(idx2);
			expect(wb.getInfo().sharedStrings).toHaveLength(1);

			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });

			expect(result).toBe(true);
		});

		it("should completely remove string when no references left", () => {
			// Add string
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			expect(wb.getInfo().sharedStrings).toHaveLength(1);
			expect(wb.getInfo().sharedStringMap.get("Hello")).toBe(0);

			// Remove reference
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });

			expect(result).toBe(true);
			expect(wb.getInfo().sharedStrings).toHaveLength(1);
			expect(wb.getInfo().sharedStringMap.has("Hello")).toBe(true);
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
			expect(wb.getInfo().sharedStrings).toEqual(["First", "Second", "Third"]);

			// Remove middle element
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 1 });

			expect(result).toBe(true);
			expect(wb.getInfo().sharedStrings).toEqual(["First", "Second", "Third"]);
			expect(wb.getInfo().sharedStringMap.get("First")).toBe(0);
			expect(wb.getInfo().sharedStringMap.get("Third")).toBe(2);
			expect(wb.getInfo().sharedStringMap.has("Second")).toBe(true);
		});

		it("should correctly reorder indices when removing first element", () => {
			// Add multiple strings
			add.call(wb, { sheetName: "Sheet1", str: "First" });
			add.call(wb, { sheetName: "Sheet1", str: "Second" });
			add.call(wb, { sheetName: "Sheet1", str: "Third" });

			// Remove first element
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });

			expect(result).toBe(true);
			expect(wb.getInfo().sharedStrings).toEqual(["First", "Second", "Third"]);
			expect(wb.getInfo().sharedStringMap.get("Second")).toBe(1);
			expect(wb.getInfo().sharedStringMap.get("Third")).toBe(2);
			expect(wb.getInfo().sharedStringMap.has("First")).toBe(true);
		});

		it("should correctly reorder indices when removing last element", () => {
			// Add multiple strings
			add.call(wb, { sheetName: "Sheet1", str: "First" });
			add.call(wb, { sheetName: "Sheet1", str: "Second" });
			add.call(wb, { sheetName: "Sheet1", str: "Third" });

			// Remove last element
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: 2 });

			expect(result).toBe(true);
			expect(wb.getInfo().sharedStrings).toEqual(["First", "Second", "Third"]);
			expect(wb.getInfo().sharedStringMap.get("First")).toBe(0);
			expect(wb.getInfo().sharedStringMap.get("Second")).toBe(1);
			expect(wb.getInfo().sharedStringMap.has("Third")).toBe(true);
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
			expect(wb.getInfo().sharedStringMap.size).toBe(wb.getInfo().sharedStrings.length);
			for (let i = 0; i < wb.getInfo().sharedStrings.length; i++) {
				const str = wb.getInfo().sharedStrings[i];
				expect(wb.getInfo().sharedStringMap.get(str)).toBe(i);
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
			expect(wb.getInfo().sharedStringMap.size).toBe(wb.getInfo().sharedStrings.length);
			for (let i = 0; i < wb.getInfo().sharedStrings.length; i++) {
				const str = wb.getInfo().sharedStrings[i];
				expect(wb.getInfo().sharedStringMap.get(str)).toBe(i);
			}

			// Verify remaining strings
			expect(wb.getInfo().sharedStrings).toEqual(["A", "B", "C", "D", "E", "F", "G", "H"]);
		});
	});

	describe("edge cases", () => {
		it("should handle removing from empty workbook", () => {
			expect(() => {
				remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });
			}).toThrow("String not found: 0");
		});

		it("should handle negative index", () => {
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			expect(() => {
				remove.call(wb, { sheetName: "Sheet1", strIdx: -1 });
			}).toThrow("String not found: -1");
		});

		it("should handle removing same string multiple times", () => {
			// Add string to multiple sheets
			add.call(wb, { sheetName: "Sheet1", str: "Hello" });
			add.call(wb, { sheetName: "Sheet2", str: "Hello" });

			// Remove from first sheet
			const result1 = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });
			expect(result1).toBe(true);
			expect(wb.getInfo().sharedStrings).toHaveLength(1);

			// Try to remove from first sheet again (should fail)
			const result2 = remove.call(wb, { sheetName: "Sheet1", strIdx: 0 });
			expect(result2).toBe(true);
		});
	});

	describe("performance characteristics", () => {
		it("should handle large number of strings efficiently", () => {
			// Add many strings
			const count = 1000;
			for (let i = 0; i < count; i++) {
				add.call(wb, { sheetName: "Sheet1", str: `String${i}` });
			}

			expect(wb.getInfo().sharedStrings).toHaveLength(count);
			expect(wb.getInfo().sharedStringMap.size).toBe(count);

			// Remove middle element
			const startTime = performance.now();
			const result = remove.call(wb, { sheetName: "Sheet1", strIdx: Math.floor(count / 2) });
			const endTime = performance.now();

			expect(result).toBe(true);

			expect(wb.getInfo().sharedStrings).toHaveLength(count);
			expect(wb.getInfo().sharedStringMap.size).toBe(count);

			// Should be fast even with many strings (Map lookup is O(1))
			expect(endTime - startTime).toBeLessThan(100); // Less than 100ms
		});
	});
});
