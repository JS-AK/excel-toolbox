import { describe, expect, it } from "vitest";

import { compareColumns } from "./compare-columns.js";

describe("compareColumns", () => {
	it("should return 0 if the columns are equal", () => {
		expect(compareColumns("A", "A")).toBe(0);
	});

	it("should return -1 if the first column is less than the second", () => {
		expect(compareColumns("A", "B")).toBe(-1);
	});

	it("should return 1 if the first column is greater than the second", () => {
		expect(compareColumns("B", "A")).toBe(1);
	});

	it("should handle columns with different lengths", () => {
		expect(compareColumns("AA", "A")).toBe(1);
		expect(compareColumns("A", "AA")).toBe(-1);
	});

	it("should handle columns with different lengths", () => {
		expect(compareColumns("AA", "A")).toBe(1);
		expect(compareColumns("A", "AA")).toBe(-1);
	});
});
