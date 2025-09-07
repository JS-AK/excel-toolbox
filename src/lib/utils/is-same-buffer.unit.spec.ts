import { describe, expect, it } from "vitest";

import { isSameBuffer } from "./is-same-buffer.js";

describe("isSameBuffer", () => {
	it("возвращает true для одинаковых буферов", () => {
		const buf1 = Buffer.from("hello");
		const buf2 = Buffer.from("hello");
		expect(isSameBuffer(buf1, buf2)).toBe(true);
	});

	it("возвращает false для разных буферов", () => {
		const buf1 = Buffer.from("hello");
		const buf2 = Buffer.from("world");
		expect(isSameBuffer(buf1, buf2)).toBe(false);
	});

	it("возвращает false если длины разные", () => {
		const buf1 = Buffer.from("hello");
		const buf2 = Buffer.from("hello!");
		expect(isSameBuffer(buf1, buf2)).toBe(false);
	});

	it("работает с пустыми буферами", () => {
		const buf1 = Buffer.from([]);
		const buf2 = Buffer.from([]);
		expect(isSameBuffer(buf1, buf2)).toBe(true);
	});

	it("разные по регистру строки дают false", () => {
		const buf1 = Buffer.from("HELLO");
		const buf2 = Buffer.from("hello");
		expect(isSameBuffer(buf1, buf2)).toBe(false);
	});
});
