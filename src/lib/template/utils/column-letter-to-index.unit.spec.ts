import { describe, expect, it } from "vitest";

import { columnLetterToIndex } from "./column-letter-to-index.js";

describe("columnLetterToIndex", () => {
	it("возвращает корректный индекс для одиночных букв", () => {
		expect(columnLetterToIndex("A")).toBe(1);
		expect(columnLetterToIndex("Z")).toBe(26);
	});

	it("возвращает корректный индекс для двухбуквенных колонок", () => {
		expect(columnLetterToIndex("AA")).toBe(27);
		expect(columnLetterToIndex("AZ")).toBe(52);
		expect(columnLetterToIndex("BA")).toBe(53);
		expect(columnLetterToIndex("ZZ")).toBe(702);
	});

	it("возвращает корректный индекс для трёхбуквенных колонок", () => {
		expect(columnLetterToIndex("AAA")).toBe(703);
		expect(columnLetterToIndex("XFD")).toBe(16384); // максимальная колонка Excel
	});

	it("возвращает -1 для недопустимых символов", () => {
		expect(columnLetterToIndex("A1")).toBe(-1);
		expect(columnLetterToIndex("a")).toBe(-1);
		expect(columnLetterToIndex("A@")).toBe(-1);
		expect(columnLetterToIndex("")).toBe(-1);
	});
});
