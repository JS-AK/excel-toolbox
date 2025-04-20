import { describe, expect, it } from "vitest";

import { checkRow } from "./check-row.js";

describe("checkRow", () => {
	it("should throw error for invalid column names", () => {
		const invalidRows: Record<string, string>[] = [
			{ "1": "value" },
			{ "A1": "invalid" },
			{ "B-2": "test" },
			{ "column": "data" }, // This will now fail as expected
			{ "": "empty" },
			{ "A B": "space" },
			{ "Д": "cyrillic" },
			{ "AAAA": "too long" }, // Excel columns max at XFD (3 letters)
		];

		invalidRows.forEach(row => {
			expect(() => checkRow(row)).toThrowError(/Invalid cell reference/);
		});
	});

	it("should accept valid Excel column references", () => {
		const validRows: Record<string, string>[] = [
			{ "A": "value" },
			{ "Z": "test" },
			{ "AA": "data" },
			{ "AZ": "valid" },
			{ "ZZ": "test" },
			{ "AAA": "valid" },
			{ "XFD": "valid" }, // Last Excel column
		];

		validRows.forEach(row => {
			expect(() => checkRow(row)).not.toThrow();
		});
	});
});
