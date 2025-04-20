import { describe, expect, it } from "vitest";

import { checkRows } from "./check-rows.js";

describe("checkRows", () => {
	it("should pass when all rows are valid", () => {
		const rows = [{ A: "valid" }];

		expect(() => checkRows(rows)).not.toThrow();
	});

	it("should throw when any row is invalid", () => {
		const rows: Record<string, string>[] = [{ A: "valid" }, { "1": "invalid" }];

		expect(() => checkRows(rows)).toThrowError("Invalid");
	});

	it("should handle empty array", () => {
		expect(() => checkRows([])).not.toThrow();
	});
});
