import { describe, expect, it } from "vitest";

import type { MergeCell } from "./types.js";

import { WorkbookBuilder } from "../workbook-builder.js";
import { add } from "./add.js";
import { remove } from "./remove.js";

describe("merge-cells/remove", () => {
	describe("remove", () => {
		it("should remove a merge cell from a sheet", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			// Add merge cell first
			add.call(wb, { ...mergeCell, sheetName: "Sheet1" });
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(1);

			// Remove merge cell
			const result = remove.call(wb, { ...mergeCell, sheetName: "Sheet1" });

			expect(result).toBe(true);
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(0);
		});

		it("should throw error if sheet is not found", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			expect(() => {
				remove.call(wb, { ...mergeCell, sheetName: "NonExistentSheet" });
			}).toThrow("Sheet not found");
		});

		it("should throw error if merge cell does not exist", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			expect(() => {
				remove.call(wb, { ...mergeCell, sheetName: "Sheet1" });
			}).toThrow("Invalid merge cell");
		});

		it("should remove correct merge cell when multiple exist", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const merge1: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			const merge2: MergeCell = {
				endCol: 4,
				endRow: 4,
				startCol: 3,
				startRow: 3,
			};

			const merge3: MergeCell = {
				endCol: 6,
				endRow: 6,
				startCol: 5,
				startRow: 5,
			};

			// Add all merge cells
			add.call(wb, { ...merge1, sheetName: "Sheet1" });
			add.call(wb, { ...merge2, sheetName: "Sheet1" });
			add.call(wb, { ...merge3, sheetName: "Sheet1" });

			expect(wb.mergeCells.get("Sheet1")).toHaveLength(3);

			// Remove middle merge cell
			const result = remove.call(wb, { ...merge2, sheetName: "Sheet1" });

			expect(result).toBe(true);
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(2);
			expect(wb.mergeCells.get("Sheet1")).toContainEqual(merge1);
			expect(wb.mergeCells.get("Sheet1")).toContainEqual(merge3);
			expect(wb.mergeCells.get("Sheet1")).not.toContainEqual(merge2);
		});

		it("should handle removing from empty sheet", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			expect(() => {
				remove.call(wb, { ...mergeCell, sheetName: "Sheet1" });
			}).toThrow("Invalid merge cell");
		});

		it("should work with multiple sheets independently", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });
			wb.addSheet("Sheet2");

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			// Add same merge cell to both sheets
			add.call(wb, { ...mergeCell, sheetName: "Sheet1" });
			add.call(wb, { ...mergeCell, sheetName: "Sheet2" });

			expect(wb.mergeCells.get("Sheet1")).toHaveLength(1);
			expect(wb.mergeCells.get("Sheet2")).toHaveLength(1);

			// Remove from Sheet1 only
			const result = remove.call(wb, { ...mergeCell, sheetName: "Sheet1" });

			expect(result).toBe(true);
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(0);
			expect(wb.mergeCells.get("Sheet2")).toHaveLength(1);
			expect(wb.mergeCells.get("Sheet2")).toContainEqual(mergeCell);
		});

		it("should handle single cell merge removal", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const singleCellMerge: MergeCell = {
				endCol: 1,
				endRow: 1,
				startCol: 1,
				startRow: 1,
			};

			// Add single cell merge
			add.call(wb, { ...singleCellMerge, sheetName: "Sheet1" });
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(1);

			// Remove single cell merge
			const result = remove.call(wb, { ...singleCellMerge, sheetName: "Sheet1" });

			expect(result).toBe(true);
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(0);
		});

		it("should handle large merge range removal", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const largeMerge: MergeCell = {
				endCol: 1000,
				endRow: 1000,
				startCol: 1,
				startRow: 1,
			};

			// Add large merge
			add.call(wb, { ...largeMerge, sheetName: "Sheet1" });
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(1);

			// Remove large merge
			const result = remove.call(wb, { ...largeMerge, sheetName: "Sheet1" });

			expect(result).toBe(true);
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(0);
		});

		it("should not remove merge cell with different coordinates", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const originalMerge: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			const differentMerge: MergeCell = {
				endCol: 3,
				endRow: 3,
				startCol: 1,
				startRow: 1,
			};

			// Add original merge
			add.call(wb, { ...originalMerge, sheetName: "Sheet1" });
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(1);

			// Try to remove different merge
			expect(() => {
				remove.call(wb, { ...differentMerge, sheetName: "Sheet1" });
			}).toThrow("Invalid merge cell");

			// Original merge should still exist
			expect(wb.mergeCells.get("Sheet1")).toHaveLength(1);
			expect(wb.mergeCells.get("Sheet1")).toContainEqual(originalMerge);
		});
	});
});
