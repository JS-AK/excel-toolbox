import { describe, expect, it } from "vitest";

import { columnIndexToLetter } from "./column-index-to-letter.js";

describe("columnIndexToLetter", () => {
	it("should convert single-letter columns (A-Z)", () => {
		expect(columnIndexToLetter(0)).toBe("A");
		expect(columnIndexToLetter(1)).toBe("B");
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
