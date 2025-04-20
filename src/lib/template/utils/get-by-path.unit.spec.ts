import { describe, expect, it } from "vitest";

import { getByPath } from "./get-by-path.js";

describe("getByPath", () => {
	const testObj = {
		a: 1,
		b: {
			c: 2,
			d: {
				e: 3,
				f: null,
				g: undefined,
			},
		},
		h: [10, 20, { i: 30 }],
		j: [],
		k: undefined,
		l: null,
	};

	it("should get top-level properties", () => {
		expect(getByPath(testObj, "a")).toBe(1);
		expect(getByPath(testObj, "k")).toBeUndefined();
		expect(getByPath(testObj, "l")).toBeNull();
	});

	it("should get nested properties", () => {
		expect(getByPath(testObj, "b.c")).toBe(2);
		expect(getByPath(testObj, "b.d.e")).toBe(3);
		expect(getByPath(testObj, "b.d.f")).toBeNull();
		expect(getByPath(testObj, "b.d.g")).toBeUndefined();
	});

	it("should handle array paths", () => {
		expect(getByPath(testObj, "h.0")).toBe(10);
		expect(getByPath(testObj, "h.2.i")).toBe(30);
		expect(getByPath(testObj, "h.3")).toBeUndefined(); // Out of bounds
	});

	it("should return undefined for invalid paths", () => {
		expect(getByPath(testObj, "x")).toBeUndefined();
		expect(getByPath(testObj, "a.x")).toBeUndefined();
		expect(getByPath(testObj, "b.x.y")).toBeUndefined();
		expect(getByPath(testObj, "h.5")).toBeUndefined();
		expect(getByPath(testObj, "j.0")).toBeUndefined(); // Empty array
		expect(getByPath(testObj, "")).toBeUndefined();
	});

	it("should handle primitive values", () => {
		expect(getByPath(42, "toString")).toBeUndefined();
		expect(getByPath("hello", "length")).toBeUndefined();
		expect(getByPath(true, "valueOf")).toBeUndefined();
		expect(getByPath(null, "any")).toBeUndefined();
		expect(getByPath(undefined, "any")).toBeUndefined();
	});
});
