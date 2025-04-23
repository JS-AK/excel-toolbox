import { describe, expect, it } from "vitest";

import { prepareRowToCells } from "./prepare-row-to-cells.js";

describe("prepareRowToCells", () => {
	it("should handle basic values", () => {
		const row = ["Text", 42, true];
		const result = prepareRowToCells(row, 1);

		expect(result).toEqual([
			"<c r=\"A1\" t=\"inlineStr\"><is><t>Text</t></is></c>",
			"<c r=\"B1\" t=\"inlineStr\"><is><t>42</t></is></c>",
			"<c r=\"C1\" t=\"inlineStr\"><is><t>true</t></is></c>",
		]);
	});

	it("should handle empty and null values", () => {
		const row = [null, undefined, ""];
		const result = prepareRowToCells(row, 2);

		expect(result).toEqual([
			"<c r=\"A2\" t=\"inlineStr\"><is><t></t></is></c>",
			"<c r=\"B2\" t=\"inlineStr\"><is><t></t></is></c>",
			"<c r=\"C2\" t=\"inlineStr\"><is><t></t></is></c>",
		]);
	});

	it("should escape XML special characters", () => {
		const row = ["<tag>", "&entity", "\"quote\"", "'apos'"];
		const result = prepareRowToCells(row, 3);

		expect(result).toEqual([
			"<c r=\"A3\" t=\"inlineStr\"><is><t>&lt;tag&gt;</t></is></c>",
			"<c r=\"B3\" t=\"inlineStr\"><is><t>&amp;entity</t></is></c>",
			"<c r=\"C3\" t=\"inlineStr\"><is><t>&quot;quote&quot;</t></is></c>",
			"<c r=\"D3\" t=\"inlineStr\"><is><t>&apos;apos&apos;</t></is></c>",
		]);
	});

	it("should handle large row numbers", () => {
		const row = ["Value"];
		const result = prepareRowToCells(row, 1048576);

		expect(result).toEqual([
			"<c r=\"A1048576\" t=\"inlineStr\"><is><t>Value</t></is></c>",
		]);
	});

	it("should handle many columns", () => {
		const row = Array(100).fill("x");
		const result = prepareRowToCells(row, 1);

		expect(result.length).toBe(100);
		expect(result[0]).toContain("A1");
		expect(result[99]).toContain("CV1"); // 100th column
	});

	it("should maintain column order", () => {
		const row = ["A", "B", "C", "D"];
		const result = prepareRowToCells(row, 5);

		expect(result[0]).toContain("A5");
		expect(result[1]).toContain("B5");
		expect(result[2]).toContain("C5");
		expect(result[3]).toContain("D5");
	});

	it("should handle objects with toString()", () => {
		const obj = {
			toString: () => "object content",
		};
		const row = [obj];
		const result = prepareRowToCells(row, 1);

		expect(result[0]).toContain("object content");
	});
});
