import type { WriteStream } from "node:fs";
import fs from "node:fs";

import { Mock, beforeEach, describe, expect, it, vi } from "vitest";

import { writeRowsToStream } from "./write-rows-to-stream.js";

describe("writeRowsToStream", () => {
	let mockStream: Partial<WriteStream>;
	let mockWrite: Mock;
	let mockEnd: Mock;

	beforeEach(() => {
		mockWrite = vi.fn();
		mockEnd = vi.fn();
		mockStream = {
			end: mockEnd,
			write: mockWrite,
		};
	});

	it("should write empty rows correctly", async () => {
		async function* emptyRows() {
			yield [];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, emptyRows(), 1);

		expect(result.rowNumber).toBe(1);
		expect(result.dimension).toEqual({ maxColumn: "A", maxRow: 1, minColumn: "A", minRow: 1 });
	});

	it("should write single row with values", async () => {
		async function* singleRow() {
			yield ["Value1", 42, true];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, singleRow(), 1);
		expect(result.rowNumber).toBe(2);
		expect(result.dimension).toEqual({ maxColumn: "C", maxRow: 1, minColumn: "A", minRow: 1 });
		expect(mockWrite).toHaveBeenCalledWith(
			"<row r=\"1\">" +
			"<c r=\"A1\" t=\"inlineStr\"><is><t>Value1</t></is></c>" +
			"<c r=\"B1\" t=\"inlineStr\"><is><t>42</t></is></c>" +
			"<c r=\"C1\" t=\"inlineStr\"><is><t>true</t></is></c>" +
			"</row>",
		);
	});

	it("should handle multiple rows", async () => {
		async function* multipleRows() {
			yield ["A"];
			yield ["B"];
			yield ["C"];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, multipleRows(), 10);
		expect(result.rowNumber).toBe(13);
		expect(result.dimension).toEqual({ maxColumn: "A", maxRow: 12, minColumn: "A", minRow: 10 });
		expect(mockWrite).toHaveBeenCalledTimes(3);
		expect(mockWrite).toHaveBeenNthCalledWith(1, "<row r=\"10\"><c r=\"A10\" t=\"inlineStr\"><is><t>A</t></is></c></row>");
		expect(mockWrite).toHaveBeenNthCalledWith(2, "<row r=\"11\"><c r=\"A11\" t=\"inlineStr\"><is><t>B</t></is></c></row>");
		expect(mockWrite).toHaveBeenNthCalledWith(3, "<row r=\"12\"><c r=\"A12\" t=\"inlineStr\"><is><t>C</t></is></c></row>");
	});

	it("should handle null/undefined values", async () => {
		async function* rowWithNulls() {
			yield [null, undefined, "valid"];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, rowWithNulls(), 1);
		expect(result.dimension).toEqual({ maxColumn: "C", maxRow: 1, minColumn: "A", minRow: 1 });
		expect(mockWrite).toHaveBeenCalledWith(
			"<row r=\"1\">" +
			"<c r=\"A1\" t=\"inlineStr\"><is><t></t></is></c>" +
			"<c r=\"B1\" t=\"inlineStr\"><is><t></t></is></c>" +
			"<c r=\"C1\" t=\"inlineStr\"><is><t>valid</t></is></c>" +
			"</row>",
		);
	});

	it("should handle special characters in values", async () => {
		async function* rowWithSpecials() {
			yield ["<>&\"'"];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, rowWithSpecials(), 1);
		expect(result.dimension).toEqual({ maxColumn: "A", maxRow: 1, minColumn: "A", minRow: 1 });
		expect(mockWrite).toHaveBeenCalledWith(
			"<row r=\"1\">" +
			"<c r=\"A1\" t=\"inlineStr\"><is><t>&lt;&gt;&amp;&quot;&apos;</t></is></c>" +
			"</row>",
		);
	});

	it("should start from specified row number", async () => {
		async function* singleRow() {
			yield ["Test"];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, singleRow(), 5);
		expect(result.dimension).toEqual({ maxColumn: "A", maxRow: 5, minColumn: "A", minRow: 5 });
		expect(mockWrite).toHaveBeenCalledWith(
			"<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>Test</t></is></c></row>",
		);
	});

	it("should handle large row numbers", async () => {
		async function* singleRow() {
			yield ["Big"];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, singleRow(), 1048576);
		expect(result.dimension).toEqual({ maxColumn: "A", maxRow: 1048576, minColumn: "A", minRow: 1048576 });
		expect(mockWrite).toHaveBeenCalledWith(
			"<row r=\"1048576\"><c r=\"A1048576\" t=\"inlineStr\"><is><t>Big</t></is></c></row>",
		);
	});

	it("should handle many columns", async () => {
		async function* wideRow() {
			yield Array(100).fill("data");
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, wideRow(), 1);
		expect(result.dimension).toEqual({ maxColumn: "CV", maxRow: 1, minColumn: "A", minRow: 1 });
		expect(mockWrite.mock.calls[0][0]).toMatch(/<row r="1">.*<\/row>/);
		expect(mockWrite.mock.calls[0][0].match(/<c /g)?.length).toBe(100);
	});

	it("should handle array of arrays (multi-dimensional)", async () => {
		async function* multiDimensionalRows() {
			yield [
				["A1", "B1", "C1"],
				["A2", "B2", "C2"],
				["A3", "B3", "C3"],
			];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, multiDimensionalRows(), 10);

		expect(result.rowNumber).toBe(13);
		expect(result.dimension).toEqual({ maxColumn: "C", maxRow: 12, minColumn: "A", minRow: 10 });
		expect(mockWrite).toHaveBeenCalledTimes(3);
	});

	it("should handle mixed single and multi-dimensional rows", async () => {
		async function* mixedRows() {
			yield ["Single1", "Row1"];
			yield [
				["Multi1", "Row1"],
				["Multi2", "Row2"],
			];
			yield ["Single2", "Row2"];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, mixedRows(), 1);

		expect(result.rowNumber).toBe(5);
		expect(result.dimension).toEqual({ maxColumn: "B", maxRow: 4, minColumn: "A", minRow: 1 });
		expect(mockWrite).toHaveBeenCalledTimes(4);
	});

	it("should handle empty arrays in multi-dimensional rows", async () => {
		async function* rowsWithEmptyArrays() {
			yield [
				[],
				["A", "B"],
				[],
			];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, rowsWithEmptyArrays(), 5);

		expect(result.rowNumber).toBe(6);
		expect(result.dimension).toEqual({ maxColumn: "B", maxRow: 5, minColumn: "A", minRow: 5 });
		expect(mockWrite).toHaveBeenCalledTimes(1);
	});

	it("should handle array of arrays with special characters", async () => {
		async function* specialCharsRows() {
			yield [
				["<test>", "&amp;"],
				["\"quotes\"", "'apos'"],
			];
		}

		const result = await writeRowsToStream(mockStream as fs.WriteStream, specialCharsRows(), 1);

		expect(result.dimension).toEqual({ maxColumn: "B", maxRow: 2, minColumn: "A", minRow: 1 });
		expect(mockWrite).toHaveBeenCalledTimes(2);
	});
});
