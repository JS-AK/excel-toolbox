import { describe, expect, it } from "vitest";

import { foundArraysInReplacements } from "./found-arrays-in-replacements.js";

describe("foundArraysInReplacements", () => {
	it("should return true if the replacements object contains arrays at the top level", () => {
		const replacements = { key: ["value1", "value2"] };

		const result = foundArraysInReplacements(replacements);

		expect(result).toBe(true);
	});

	it("should return true if the replacements object contains arrays at the second level", () => {
		const replacements = { key: { subkey: ["value1", "value2"] } };

		const result = foundArraysInReplacements(replacements);

		expect(result).toBe(true);
	});

	it("should return false if the replacements object does not contain arrays", () => {
		const replacements = { key: "value" };

		const result = foundArraysInReplacements(replacements);

		expect(result).toBe(false);
	});
});
