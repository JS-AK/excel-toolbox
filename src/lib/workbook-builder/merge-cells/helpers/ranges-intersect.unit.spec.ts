/* eslint-disable sort-keys */
import { describe, expect, it } from "vitest";

import { MergeCell } from "../../types/index.js";

import { rangesIntersect } from "./ranges-intersect.js";

describe("rangesIntersect", () => {
	it("should return true for overlapping ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 3, endCol: 3 };
		const range2: MergeCell = { startRow: 2, startCol: 2, endRow: 4, endCol: 4 };

		expect(rangesIntersect(range1, range2)).toBe(true);
	});

	it("should return true for identical ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };

		expect(rangesIntersect(range1, range2)).toBe(true);
	});

	it("should return true when one range contains another", () => {
		const outer: MergeCell = { startRow: 1, startCol: 1, endRow: 5, endCol: 5 };
		const inner: MergeCell = { startRow: 2, startCol: 2, endRow: 3, endCol: 3 };

		expect(rangesIntersect(outer, inner)).toBe(true);
		expect(rangesIntersect(inner, outer)).toBe(true);
	});

	it("should return true for ranges that share an edge", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 2, startCol: 1, endRow: 3, endCol: 2 };

		expect(rangesIntersect(range1, range2)).toBe(true);
	});

	it("should return true for ranges that share a corner", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 2, startCol: 2, endRow: 3, endCol: 3 };

		expect(rangesIntersect(range1, range2)).toBe(true);
	});

	it("should return false for non-overlapping ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 4, startCol: 4, endRow: 5, endCol: 5 };

		expect(rangesIntersect(range1, range2)).toBe(false);
	});

	it("should return false for ranges separated by rows", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 5 };
		const range2: MergeCell = { startRow: 4, startCol: 1, endRow: 5, endCol: 5 };

		expect(rangesIntersect(range1, range2)).toBe(false);
	});

	it("should return false for ranges separated by columns", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 5, endCol: 2 };
		const range2: MergeCell = { startRow: 1, startCol: 4, endRow: 5, endCol: 5 };

		expect(rangesIntersect(range1, range2)).toBe(false);
	});

	it("should handle single cell ranges", () => {
		const cell1: MergeCell = { startRow: 1, startCol: 1, endRow: 1, endCol: 1 };
		const cell2: MergeCell = { startRow: 1, startCol: 1, endRow: 1, endCol: 1 };
		const cell3: MergeCell = { startRow: 2, startCol: 2, endRow: 2, endCol: 2 };

		expect(rangesIntersect(cell1, cell2)).toBe(true);
		expect(rangesIntersect(cell1, cell3)).toBe(false);
	});

	it("should handle ranges with zero coordinates", () => {
		const range1: MergeCell = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range3: MergeCell = { startRow: 2, startCol: 2, endRow: 3, endCol: 3 };

		expect(rangesIntersect(range1, range2)).toBe(true);
		expect(rangesIntersect(range1, range3)).toBe(false);
	});

	it("should handle large ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 1000, endCol: 1000 };
		const range2: MergeCell = { startRow: 500, startCol: 500, endRow: 1500, endCol: 1500 };
		const range3: MergeCell = { startRow: 2000, startCol: 2000, endRow: 3000, endCol: 3000 };

		expect(rangesIntersect(range1, range2)).toBe(true);
		expect(rangesIntersect(range1, range3)).toBe(false);
	});

	it("should handle edge case with maximum values", () => {
		const maxValue = Number.MAX_SAFE_INTEGER;
		const range1: MergeCell = { startRow: maxValue - 1, startCol: maxValue - 1, endRow: maxValue, endCol: maxValue };
		const range2: MergeCell = { startRow: maxValue, startCol: maxValue, endRow: maxValue, endCol: maxValue };
		const range3: MergeCell = { startRow: maxValue - 2, startCol: maxValue - 2, endRow: maxValue - 2, endCol: maxValue - 2 };

		expect(rangesIntersect(range1, range2)).toBe(true);
		expect(rangesIntersect(range1, range3)).toBe(false);
	});

	it("should handle partial overlaps in rows only", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 3, endCol: 2 };
		const range2: MergeCell = { startRow: 2, startCol: 5, endRow: 4, endCol: 6 };

		expect(rangesIntersect(range1, range2)).toBe(false);
	});

	it("should handle partial overlaps in columns only", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 3 };
		const range2: MergeCell = { startRow: 5, startCol: 2, endRow: 6, endCol: 4 };

		expect(rangesIntersect(range1, range2)).toBe(false);
	});

	it("should handle ranges that touch but do not overlap", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 3, startCol: 1, endRow: 4, endCol: 2 };

		expect(rangesIntersect(range1, range2)).toBe(false);
	});
});
