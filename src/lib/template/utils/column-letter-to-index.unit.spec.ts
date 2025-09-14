import { describe, expect, it } from "vitest";

import { columnLetterToIndex } from "./column-letter-to-index.js";

describe("columnLetterToIndex", () => {
	it("возвращает корректный индекс для одиночных букв", () => {
		expect(columnLetterToIndex("A")).toBe(1);
		expect(columnLetterToIndex("B")).toBe(2);
		expect(columnLetterToIndex("C")).toBe(3);
		expect(columnLetterToIndex("D")).toBe(4);
		expect(columnLetterToIndex("E")).toBe(5);
		expect(columnLetterToIndex("F")).toBe(6);
		expect(columnLetterToIndex("G")).toBe(7);
		expect(columnLetterToIndex("H")).toBe(8);
		expect(columnLetterToIndex("I")).toBe(9);
		expect(columnLetterToIndex("J")).toBe(10);
		expect(columnLetterToIndex("K")).toBe(11);
		expect(columnLetterToIndex("L")).toBe(12);
		expect(columnLetterToIndex("M")).toBe(13);
		expect(columnLetterToIndex("N")).toBe(14);
		expect(columnLetterToIndex("O")).toBe(15);
		expect(columnLetterToIndex("P")).toBe(16);
		expect(columnLetterToIndex("Q")).toBe(17);
		expect(columnLetterToIndex("R")).toBe(18);
		expect(columnLetterToIndex("S")).toBe(19);
		expect(columnLetterToIndex("T")).toBe(20);
		expect(columnLetterToIndex("U")).toBe(21);
		expect(columnLetterToIndex("V")).toBe(22);
		expect(columnLetterToIndex("W")).toBe(23);
		expect(columnLetterToIndex("X")).toBe(24);
		expect(columnLetterToIndex("Y")).toBe(25);
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
