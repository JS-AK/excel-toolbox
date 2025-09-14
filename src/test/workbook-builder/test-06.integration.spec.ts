import { beforeEach, describe, expect, it } from "vitest";

import { WorkbookBuilder } from "../../lib/workbook-builder/workbook-builder";

describe("Shared String Ref Integration Tests", () => {
	let wb: WorkbookBuilder;

	beforeEach(async () => {
		wb = new WorkbookBuilder({ cleanupUnused: true });
		wb.addSheet("Sheet2");
	});

	describe("complete workflow", () => {
		it("should handle complex scenario with multiple sheets and strings", async () => {
			// Add strings to multiple sheets
			const sheet1 = wb.getSheet("Sheet1");

			if (!sheet1) {
				throw new Error("Sheet 'Sheet1' not found");
			}

			const sheet2 = wb.getSheet("Sheet2");

			if (!sheet2) {
				throw new Error("Sheet 'Sheet2' not found");
			}

			sheet1.setCell(1, "A", { type: "s", value: "Hello" });
			sheet1.setCell(2, "A", { type: "s", value: "World" });

			sheet2.setCell(1, "A", { type: "s", value: "Hello" }); // Duplicate
			sheet2.setCell(2, "A", { type: "s", value: "Test" });

			// Verify initial state
			expect(wb.getInfo().sharedStrings).toEqual(["Hello", "World", "Test"]);

			// Remove "Hello" from Sheet1 (should keep string, remove reference)
			sheet1.setCell(1, "A", { type: "inlineStr", value: "Hello" }); // Upsert with inline string

			expect(wb.getInfo().sharedStrings).toEqual(["Hello", "World", "Test"]);

			// Remove "World" from Sheet1 (should remove string completely)
			sheet1.setCell(2, "A", { type: "inlineStr", value: "World" }); // Upsert with inline string

			expect(wb.getInfo().sharedStrings).toEqual(["Hello", "World", "Test"]);
			expect(wb.getInfo().sharedStringMap.get("Hello")).toBe(0);
			expect(wb.getInfo().sharedStringMap.get("World")).toBe(1);
			expect(wb.getInfo().sharedStringMap.get("Test")).toBe(2);

			// Remove all from Sheet2 (should remove remaining strings)
			wb.removeSheet("Sheet2");

			expect(wb.getInfo().sharedStrings).toHaveLength(3);
			expect(wb.getInfo().sharedStringMap.size).toBe(3);
		});

		it("should maintain consistency during rapid add/remove operations", async () => {
			const sheet1 = wb.getSheet("Sheet1");

			if (!sheet1) {
				throw new Error("Sheet 'Sheet1' not found");
			}

			const sheet2 = wb.getSheet("Sheet2");

			if (!sheet2) {
				throw new Error("Sheet 'Sheet2' not found");
			}

			sheet1.setCell(1, "A", { type: "s", value: "A" });
			// shared-string-ref 0 -> A

			sheet1.setCell(2, "A", { type: "s", value: "B" });
			// shared-string-ref 1 -> B

			sheet2.setCell(1, "A", { type: "s", value: "A" });
			// shared-string-ref A exists -> 0

			sheet1.setCell(1, "A", { type: "inlineStr", value: "A" });
			// shared-string-ref A removed from Sheet1 (but exists in Sheet2)

			sheet1.setCell(3, "A", { type: "s", value: "C" });
			// shared-string-ref 2 -> C

			sheet2.setCell(1, "A", { type: "inlineStr", value: "A" });
			// shared-string-ref A removed from Sheet2

			sheet1.setCell(4, "A", { type: "s", value: "D" });
			// shared-string-ref 3 -> D

			sheet1.setCell(2, "A", { type: "inlineStr", value: "B" });
			// shared-string-ref B removed from Sheet1

			// Final state should be consistent
			expect(wb.getInfo().sharedStrings).toEqual(["A", "B", "C", "D"]);
			expect(wb.getInfo().sharedStringMap.size).toBe(4);

			expect(wb.getInfo().sharedStringMap.get("A")).toBe(0);
			expect(wb.getInfo().sharedStringMap.get("B")).toBe(1);
			expect(wb.getInfo().sharedStringMap.get("C")).toBe(2);
			expect(wb.getInfo().sharedStringMap.get("D")).toBe(3);
		});
	});

	describe("Map optimization verification", () => {
		it("should demonstrate O(1) lookup performance", () => {
			// Add many strings
			const count = 10000;
			const strings = Array.from({ length: count }, (_, i) => `String${i}`);

			// Add all strings
			const sheet = wb.getSheet("Sheet1");

			if (!sheet) {
				throw new Error("Sheet 'Sheet1' not found");
			}

			let row = 1;

			for (const str of strings) {
				sheet.setCell(row++, "A", { type: "s", value: str });
			}

			// Test lookup performance
			const startTime = performance.now();
			const info = wb.getInfo();

			for (let i = 0; i < 1000; i++) {
				const randomStr = strings[Math.floor(Math.random() * strings.length)];
				const idx = info.sharedStringMap.get(randomStr);
				expect(idx).toBeDefined();
				expect(info.sharedStrings[idx!]).toBe(randomStr);
			}
			const endTime = performance.now();

			// Should be very fast (O(1) lookup)
			expect(endTime - startTime).toBeLessThan(100); // Less than 100ms for 1000 lookups
		});

		it("should maintain Map consistency during bulk operations", () => {
			// Add many strings
			const count = 1000;

			const sheet = wb.getSheet("Sheet1");
			if (!sheet) {
				throw new Error("Sheet 'Sheet1' not found");
			}

			let row1 = 1;

			for (let i = 0; i < count; i++) {
				sheet.setCell(row1++, "A", { type: "s", value: `String${i}` });
			}

			let row2 = 1;

			// Remove every other string
			for (let i = count - 1; i >= 0; i -= 2) {
				sheet.setCell(row2++, "A", { type: "inlineStr", value: `String${i}` });
			}

			// Verify final consistency
			expect(wb.getInfo().sharedStringMap.size).toBe(wb.getInfo().sharedStrings.length);

			for (let i = 0; i < wb.getInfo().sharedStrings.length; i++) {
				const str = wb.getInfo().sharedStrings[i];
				expect(wb.getInfo().sharedStringMap.get(str)).toBe(i);
			}

			// Should have half the strings remaining
			expect(wb.getInfo().sharedStrings.length).toBe(count);
		});
	});
});
