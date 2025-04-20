import { describe, expect, it } from "vitest";

import { applyReplacements } from "./apply-replacements.js";

describe("applyReplacements", () => {
	const testReplacements = {
		count: 5,
		empty: "",
		nullValue: null,
		tags: ["js", "ts"],
		undefinedValue: undefined,
		user: {
			active: true,
			address: { city: "New York", zip: null },
			age: 30,
			name: "John",
		},
	};

	it("should replace simple placeholders", () => {
		const template = "Hello ${user.name}, you are ${user.age} years old.";
		const expected = "Hello John, you are 30 years old.";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle boolean and number values", () => {
		const template = "Active: ${user.active}, Count: ${count}";
		const expected = "Active: true, Count: 5";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle nested objects", () => {
		const template = "City: ${user.address.city}, Zip: ${user.address.zip}";
		const expected = "City: New York, Zip: null";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should keep original placeholder when path not found", () => {
		const template = "Unknown: ${user.unknown}, Missing: ${missing.prop}";
		const expected = "Unknown: ${user.unknown}, Missing: ${missing.prop}";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle empty strings and null/undefined values", () => {
		const template = "Empty: ${empty}, Null: ${nullValue}, Undefined: ${undefinedValue}";
		const expected = "Empty: , Null: null, Undefined: ${undefinedValue}";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle array values", () => {
		const template = "Tags: ${tags.0} and ${tags.1}";
		const expected = "Tags: js and ts";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle multiple occurrences", () => {
		const template = "${user.name}-${user.name}-${user.name}";
		const expected = "John-John-John";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle empty template", () => {
		expect(applyReplacements("", testReplacements)).toBe("");
	});

	it("should handle template without placeholders", () => {
		const template = "Just a regular string";
		expect(applyReplacements(template, testReplacements)).toBe(template);
	});

	it("should handle complex cases", () => {
		const template = "User: ${user.name} (${user.age}), ${user.address.city} ${count} times";
		const expected = "User: John (30), New York 5 times";
		expect(applyReplacements(template, testReplacements)).toBe(expected);
	});

	it("should handle edge cases", () => {
		// Malformed placeholder
		expect(applyReplacements("${user.name", testReplacements)).toBe("${user.name");
		expect(applyReplacements("user.name}", testReplacements)).toBe("user.name}");

		// Empty replacements
		expect(applyReplacements("${any}", {})).toBe("${any}");
	});
});
