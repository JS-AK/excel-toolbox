/* eslint-disable @typescript-eslint/no-explicit-any */

import { describe, expect, it } from "vitest";

import { checkStartRow } from "./check-start-row.js";

describe("checkStartRow", () => {
	it("should accept undefined value", () => {
		expect(() => checkStartRow()).not.toThrow();
		expect(() => checkStartRow(undefined)).not.toThrow();
	});

	it("should accept positive integers", () => {
		expect(() => checkStartRow(1)).not.toThrow();
		expect(() => checkStartRow(10)).not.toThrow();
		expect(() => checkStartRow(1000)).not.toThrow();
	});

	it("should reject non-integer numbers", () => {
		expect(() => checkStartRow(1.5)).toThrowError(
			"Invalid startRow \"1.5\". Must be a positive integer.",
		);
		expect(() => checkStartRow(0.99)).toThrowError(
			"Invalid startRow \"0.99\". Must be a positive integer.",
		);
	});

	it("should reject zero and negative numbers", () => {
		expect(() => checkStartRow(0)).toThrowError(
			"Invalid startRow \"0\". Must be a positive integer.",
		);
		expect(() => checkStartRow(-1)).toThrowError(
			"Invalid startRow \"-1\". Must be a positive integer.",
		);
		expect(() => checkStartRow(-100)).toThrowError(
			"Invalid startRow \"-100\". Must be a positive integer.",
		);
	});

	it("should reject non-number values", () => {
		expect(() => checkStartRow(null as any)).toThrowError(
			"Invalid startRow \"null\". Must be a positive integer.",
		);
		expect(() => checkStartRow("1" as any)).toThrowError(
			"Invalid startRow \"1\". Must be a positive integer.",
		);
		expect(() => checkStartRow(true as any)).toThrowError(
			"Invalid startRow \"true\". Must be a positive integer.",
		);
		expect(() => checkStartRow({} as any)).toThrowError(
			"Invalid startRow \"[object Object]\". Must be a positive integer.",
		);
	});

	it("should provide specific error messages", () => {
		try {
			checkStartRow(-5);
		} catch (e) {
			expect(e.message).toBe("Invalid startRow \"-5\". Must be a positive integer.");
		}

		try {
			checkStartRow(3.14);
		} catch (e) {
			expect(e.message).toBe("Invalid startRow \"3.14\". Must be a positive integer.");
		}
	});
});
