/* eslint-disable sort-keys */
import { describe, expect, it } from "vitest";

import { MergeCell } from "../../types/index.js";

import { rangesEqual } from "./ranges-equal.js";

describe("rangesEqual", () => {
	it("should return true for identical ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };

		expect(rangesEqual(range1, range2)).toBe(true);
	});

	it("should return false for different start rows", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 2, startCol: 1, endRow: 2, endCol: 2 };

		expect(rangesEqual(range1, range2)).toBe(false);
	});

	it("should return false for different end rows", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 3, endCol: 2 };

		expect(rangesEqual(range1, range2)).toBe(false);
	});

	it("should return false for different start columns", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 1, startCol: 2, endRow: 2, endCol: 2 };

		expect(rangesEqual(range1, range2)).toBe(false);
	});

	it("should return false for different end columns", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 3 };

		expect(rangesEqual(range1, range2)).toBe(false);
	});

	it("should return false for completely different ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 2, endCol: 2 };
		const range2: MergeCell = { startRow: 5, startCol: 5, endRow: 10, endCol: 10 };

		expect(rangesEqual(range1, range2)).toBe(false);
	});

	it("should handle single cell ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 1, endCol: 1 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 1, endCol: 1 };
		const range3: MergeCell = { startRow: 2, startCol: 2, endRow: 2, endCol: 2 };

		expect(rangesEqual(range1, range2)).toBe(true);
		expect(rangesEqual(range1, range3)).toBe(false);
	});

	it("should handle large ranges", () => {
		const range1: MergeCell = { startRow: 1, startCol: 1, endRow: 1000, endCol: 1000 };
		const range2: MergeCell = { startRow: 1, startCol: 1, endRow: 1000, endCol: 1000 };
		const range3: MergeCell = { startRow: 1, startCol: 1, endRow: 1001, endCol: 1000 };

		expect(rangesEqual(range1, range2)).toBe(true);
		expect(rangesEqual(range1, range3)).toBe(false);
	});

	it("should handle zero-based coordinates", () => {
		const range1: MergeCell = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
		const range2: MergeCell = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
		const range3: MergeCell = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };

		expect(rangesEqual(range1, range2)).toBe(true);
		expect(rangesEqual(range1, range3)).toBe(false);
	});

	it("should handle edge case with maximum values", () => {
		const maxValue = Number.MAX_SAFE_INTEGER;
		const range1: MergeCell = { startRow: maxValue, startCol: maxValue, endRow: maxValue, endCol: maxValue };
		const range2: MergeCell = { startRow: maxValue, startCol: maxValue, endRow: maxValue, endCol: maxValue };
		const range3: MergeCell = { startRow: maxValue - 1, startCol: maxValue, endRow: maxValue, endCol: maxValue };

		expect(rangesEqual(range1, range2)).toBe(true);
		expect(rangesEqual(range1, range3)).toBe(false);
	});
});
