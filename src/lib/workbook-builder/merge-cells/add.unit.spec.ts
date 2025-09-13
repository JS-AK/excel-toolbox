import { describe, expect, it } from "vitest";

import type { MergeCell } from "./types.js";

import { WorkbookBuilder } from "../workbook-builder.js";
import { add } from "./add.js";

describe("merge-cells/add", () => {
	describe("add", () => {
		it("should add a merge cell to a sheet", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			const result = add.call(wb, { ...mergeCell, sheetName: "Sheet1" });

			expect(result).toEqual(mergeCell);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toContainEqual(mergeCell);
		});

		it("should return existing merge cell if identical range already exists", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const mergeCell: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			// Add first time
			const firstResult = add.call(wb, { ...mergeCell, sheetName: "Sheet1" });

			// Add second time (identical)
			const secondResult = add.call(wb, { ...mergeCell, sheetName: "Sheet1" });

			expect(firstResult).toBe(secondResult);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toHaveLength(1);
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
				add.call(wb, { ...mergeCell, sheetName: "NonExistentSheet" });
			}).toThrow("Sheet not found");
		});

		it("should throw error if merge intersects with existing merge cell", () => {
			const wb = new WorkbookBuilder();

			const existingMerge: MergeCell = {
				endCol: 3,
				endRow: 3,
				startCol: 1,
				startRow: 1,
			};

			const intersectingMerge: MergeCell = {
				endCol: 4,
				endRow: 4,
				startCol: 2,
				startRow: 2,
			};

			// Add existing merge
			add.call(wb, { ...existingMerge, sheetName: "Sheet1" });

			// Try to add intersecting merge
			expect(() => {
				add.call(wb, { ...intersectingMerge, sheetName: "Sheet1" });
			}).toThrow("Merge intersects existing merged cell");
		});

		it("should allow non-intersecting merge cells on the same sheet", () => {
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

			const result1 = add.call(wb, { ...merge1, sheetName: "Sheet1" });
			const result2 = add.call(wb, { ...merge2, sheetName: "Sheet1" });

			expect(result1).toEqual(merge1);
			expect(result2).toEqual(merge2);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toHaveLength(2);
		});

		it("should handle edge case: adjacent merge cells (touching but not intersecting)", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const merge1: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			const merge2: MergeCell = {
				endCol: 4,
				endRow: 2,
				startCol: 3, // Adjacent column
				startRow: 1,
			};

			const result1 = add.call(wb, { ...merge1, sheetName: "Sheet1" });
			const result2 = add.call(wb, { ...merge2, sheetName: "Sheet1" });

			expect(result1).toEqual(merge1);
			expect(result2).toEqual(merge2);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toHaveLength(2);
		});

		it("should handle single cell merge (start equals end)", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const singleCellMerge: MergeCell = {
				endCol: 1,
				endRow: 1,
				startCol: 1,
				startRow: 1,
			};

			const result = add.call(wb, { ...singleCellMerge, sheetName: "Sheet1" });

			expect(result).toEqual(singleCellMerge);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toContainEqual(singleCellMerge);
		});

		it("should handle large merge ranges", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });

			const largeMerge: MergeCell = {
				endCol: 1000,
				endRow: 1000,
				startCol: 1,
				startRow: 1,
			};

			const result = add.call(wb, { ...largeMerge, sheetName: "Sheet1" });

			expect(result).toEqual(largeMerge);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toContainEqual(largeMerge);
		});

		it("should work with multiple sheets independently", () => {
			const wb = new WorkbookBuilder({ defaultSheetName: "Sheet1" });
			wb.addSheet("Sheet2");

			const merge1: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			const merge2: MergeCell = {
				endCol: 2,
				endRow: 2,
				startCol: 1,
				startRow: 1,
			};

			// Add same range to different sheets
			const result1 = add.call(wb, { ...merge1, sheetName: "Sheet1" });
			const result2 = add.call(wb, { ...merge2, sheetName: "Sheet2" });

			expect(result1).toEqual(merge1);
			expect(result2).toEqual(merge2);
			expect(wb.getInfo().mergeCells.get("Sheet1")).toHaveLength(1);
			expect(wb.getInfo().mergeCells.get("Sheet2")).toHaveLength(1);
		});
	});
});
