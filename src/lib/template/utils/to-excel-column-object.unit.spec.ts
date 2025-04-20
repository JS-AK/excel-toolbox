import { describe, expect, it } from "vitest";

import { toExcelColumnObject } from "./to-excel-column-object.js";

describe("toExcelColumnObject", () => {
	it("should convert array to Excel column object", () => {
		const input = ["Value1", 42, true, null, undefined];
		const expected = {
			A: "Value1",
			B: "42",
			C: "true",
			D: "null",
			E: "undefined",
		};
		expect(toExcelColumnObject(input)).toEqual(expected);
	});

	it("should handle empty array", () => {
		expect(toExcelColumnObject([])).toEqual({});
	});

	it("should generate correct Excel column names", () => {
		const input = Array(30).fill("test"); // More than 26 items to test multi-letter columns
		const result = toExcelColumnObject(input);

		// Check first 26 columns (A-Z)
		for (let i = 0; i < 26; i++) {
			const expectedKey = String.fromCharCode(65 + i);
			expect(result[expectedKey]).toBe("test");
		}

		// Check multi-letter columns
		expect(result["AA"]).toBe("test");
		expect(result["AB"]).toBe("test");
		expect(result["AC"]).toBe("test");
		expect(result["AD"]).toBe("test");
	});

	it("should convert all values to strings", () => {
		const input = [
			123,
			true,
			false,
			null,
			undefined,
			{ toString: () => "custom" },
			["array"],
			new Date(0),
		];
		const expected = {
			A: "123",
			B: "true",
			C: "false",
			D: "null",
			E: "undefined",
			F: "custom",
			G: "array",
			H: new Date(0).toString(),
		};
		expect(toExcelColumnObject(input)).toEqual(expected);
	});

	it("should handle very large arrays", () => {
		const largeArray = Array(1000).fill("x");
		const result = toExcelColumnObject(largeArray);

		// Spot check some columns
		expect(result["A"]).toBe("x");
		expect(result["Z"]).toBe("x");
		expect(result["AA"]).toBe("x");
		expect(result["AZ"]).toBe("x");
		expect(result["ZZ"]).toBe("x");
		expect(result["AAA"]).toBe("x");
		expect(result["ALL"]).toBe("x"); // 1000th column
	});

	describe("toExcelColumn helper", () => {
		it("should generate single-letter columns", () => {
			const fn = (i: number) => toExcelColumnObject(Array(i + 1).fill(0));

			expect(Object.keys(fn(0))[0]).toBe("A");
			expect(Object.keys(fn(25))[25]).toBe("Z");
		});

		it("should generate double-letter columns", () => {
			const fn = (i: number) => toExcelColumnObject(Array(i + 1).fill(0));

			expect(Object.keys(fn(26))[26]).toBe("AA");
			expect(Object.keys(fn(27))[27]).toBe("AB");
			expect(Object.keys(fn(51))[51]).toBe("AZ");
		});

		it("should generate triple-letter columns", () => {
			const fn = (i: number) => toExcelColumnObject(Array(i + 1).fill(0));

			expect(Object.keys(fn(26 * 26))[26 * 26]).toBe("ZA");
			expect(Object.keys(fn(26 * 26 + 25))[26 * 26 + 25]).toBe("ZZ");
			expect(Object.keys(fn(26 * 26 + 26))[26 * 26 + 26]).toBe("AAA");
		});
	});
});
