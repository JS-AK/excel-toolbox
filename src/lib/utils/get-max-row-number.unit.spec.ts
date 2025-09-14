import { describe, expect, it } from "vitest";

import { getMaxRowNumber } from "./get-max-row-number.js";

describe("getMaxRowNumber", () => {
	it("должен вернуть 0 для пустого массива", () => {
		expect(getMaxRowNumber([])).toBe(0);
	});

	it("должен вернуть 0 если нет атрибутов r", () => {
		const rows = [
			"<row></row>",
			"<row spans=\"1:3\"></row>",
		];
		expect(getMaxRowNumber(rows)).toBe(0);
	});

	it("должен вернуть правильный номер одной строки", () => {
		const rows = ["<row r=\"5\"></row>"];
		expect(getMaxRowNumber(rows)).toBe(5);
	});

	it("должен вернуть максимальный номер из нескольких строк", () => {
		const rows = [
			"<row r=\"1\"></row>",
			"<row r=\"3\"></row>",
			"<row r=\"10\"></row>",
		];
		expect(getMaxRowNumber(rows)).toBe(10);
	});

	it("игнорирует строки без r", () => {
		const rows = [
			"<row></row>",
			"<row r=\"7\"></row>",
			"<row></row>",
		];
		expect(getMaxRowNumber(rows)).toBe(7);
	});

	it("должен корректно парсить если атрибуты расположены в другом порядке", () => {
		const rows = [
			"<row spans=\"1:3\" r=\"8\"></row>",
			"<row r=\"12\" spans=\"1:5\"></row>",
		];
		expect(getMaxRowNumber(rows)).toBe(12);
	});

	it("игнорирует некорректные r (не числа)", () => {
		const rows = [
			"<row r=\"abc\"></row>",
			"<row r=\"15\"></row>",
		];
		expect(getMaxRowNumber(rows)).toBe(15);
	});

	it("игнорирует строки где r есть но пустое", () => {
		const rows = [
			"<row r=\"\"></row>",
			"<row r=\"20\"></row>",
		];
		expect(getMaxRowNumber(rows)).toBe(20);
	});
});
