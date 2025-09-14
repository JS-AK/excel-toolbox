import { describe, expect, it } from "vitest";

import { columnIndexToLetter } from "./column-index-to-letter.js";

describe("columnIndexToLetter", () => {
	it("should convert single-letter columns (A-Z)", () => {
		expect(columnIndexToLetter(0)).toBe("A");
		expect(columnIndexToLetter(1)).toBe("B");
		expect(columnIndexToLetter(2)).toBe("C");
		expect(columnIndexToLetter(3)).toBe("D");
		expect(columnIndexToLetter(4)).toBe("E");
		expect(columnIndexToLetter(5)).toBe("F");
		expect(columnIndexToLetter(6)).toBe("G");
		expect(columnIndexToLetter(7)).toBe("H");
		expect(columnIndexToLetter(8)).toBe("I");
		expect(columnIndexToLetter(9)).toBe("J");
		expect(columnIndexToLetter(10)).toBe("K");
		expect(columnIndexToLetter(11)).toBe("L");
		expect(columnIndexToLetter(12)).toBe("M");
		expect(columnIndexToLetter(13)).toBe("N");
		expect(columnIndexToLetter(14)).toBe("O");
		expect(columnIndexToLetter(15)).toBe("P");
		expect(columnIndexToLetter(16)).toBe("Q");
		expect(columnIndexToLetter(17)).toBe("R");
		expect(columnIndexToLetter(18)).toBe("S");
		expect(columnIndexToLetter(19)).toBe("T");
		expect(columnIndexToLetter(20)).toBe("U");
		expect(columnIndexToLetter(21)).toBe("V");
		expect(columnIndexToLetter(22)).toBe("W");
		expect(columnIndexToLetter(23)).toBe("X");
		expect(columnIndexToLetter(24)).toBe("Y");
		expect(columnIndexToLetter(25)).toBe("Z");
	});

	it("should convert double-letter columns (AA-ZZ)", () => {
		expect(columnIndexToLetter(26)).toBe("AA");
		expect(columnIndexToLetter(27)).toBe("AB");
		expect(columnIndexToLetter(51)).toBe("AZ");
		expect(columnIndexToLetter(52)).toBe("BA");
		expect(columnIndexToLetter(77)).toBe("BZ");
		expect(columnIndexToLetter(701)).toBe("ZZ");
	});

	it("should convert triple-letter columns (AAA-XFD)", () => {
		expect(columnIndexToLetter(702)).toBe("AAA");
		expect(columnIndexToLetter(703)).toBe("AAB");
		expect(columnIndexToLetter(728)).toBe("ABA");
		expect(columnIndexToLetter(729)).toBe("ABB");
		expect(columnIndexToLetter(1378)).toBe("BAA");
		expect(columnIndexToLetter(16383)).toBe("XFD"); // Excel's max column
	});

	it("should handle edge cases", () => {
		expect(columnIndexToLetter(26 + 26 + 0)).toBe("BA");
		expect(columnIndexToLetter(26 * 26 + 26 + 0)).toBe("AAA");
	});

	it("should throw for invalid inputs", () => {
		expect(() => columnIndexToLetter(-1)).toThrow();
		expect(() => columnIndexToLetter(-100)).toThrow();
		expect(() => columnIndexToLetter(1.5)).toThrow();
		expect(() => columnIndexToLetter(NaN)).toThrow();
		expect(() => columnIndexToLetter(Infinity)).toThrow();
	});

	it("should handle non-integer inputs by throwing", () => {
		expect(() => columnIndexToLetter(3.14)).toThrow();
	});

	it("should match Excel column numbering", () => {
		// Spot check some Excel-known columns
		expect(columnIndexToLetter(0)).toBe("A");
		expect(columnIndexToLetter(25)).toBe("Z");
		expect(columnIndexToLetter(26)).toBe("AA");
		expect(columnIndexToLetter(51)).toBe("AZ");
		expect(columnIndexToLetter(52)).toBe("BA");
		expect(columnIndexToLetter(77)).toBe("BZ");
		expect(columnIndexToLetter(701)).toBe("ZZ");
		expect(columnIndexToLetter(702)).toBe("AAA");
		expect(columnIndexToLetter(16383)).toBe("XFD"); // Last Excel column
	});

	it("should handle large numbers", () => {
		expect(columnIndexToLetter(1000)).toBe("ALM");
		expect(columnIndexToLetter(5000)).toBe("GJI");
		expect(columnIndexToLetter(10000)).toBe("NTQ");
	});
});
