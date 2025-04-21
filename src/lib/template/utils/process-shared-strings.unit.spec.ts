import { describe, expect, it } from "vitest";

import { processSharedStrings } from "./process-shared-strings.js";

describe("processSharedStrings", () => {
	it("should handle empty XML", () => {
		const result = processSharedStrings("");
		expect(result.sharedStrings).toEqual([]);
		expect(result.sharedIndexMap.size).toBe(0);
		expect(result.sharedStringsHeader).toBeNull();
		expect(result.sheetMergeCells).toEqual([]);
	});

	it("should parse basic shared strings", () => {
		const xml = `
      <?xml version="1.0"?>
      <sst>
        <si><t>Hello</t></si>
        <si><t>World</t></si>
      </sst>
    `;
		const result = processSharedStrings(xml);

		expect(result.sharedStrings).toEqual([
			"<si><t>Hello</t></si>",
			"<si><t>World</t></si>",
		]);
		expect(result.sharedIndexMap.get("<t>Hello</t>")).toBe(0);
		expect(result.sharedIndexMap.get("<t>World</t>")).toBe(1);
		expect(result.sharedStringsHeader).toBe("<?xml version=\"1.0\"?>");
	});

	it("should handle complex shared string items", () => {
		const xml = `
      <sst>
        <si><t>Simple</t></si>
        <si><r><t>Rich</t><t>Text</t></r></si>
        <si><t xml:space="preserve"> With Space </t></si>
      </sst>
    `;
		const result = processSharedStrings(xml);

		expect(result.sharedStrings).toEqual([
			"<si><t>Simple</t></si>",
			"<si><r><t>Rich</t><t>Text</t></r></si>",
			"<si><t xml:space=\"preserve\"> With Space </t></si>",
		]);
		expect(result.sharedIndexMap.size).toBe(3);
	});

	it("should handle empty shared strings", () => {
		const xml = `
      <sst>
        <si><t></t></si>
        <si></si>
      </sst>
    `;
		const result = processSharedStrings(xml);

		expect(result.sharedStrings).toEqual([
			"<si><t></t></si>",
		]);
	});

	it("should handle special characters in shared strings", () => {
		const xml = `
      <sst>
        <si><t>&amp;&lt;&gt;</t></si>
        <si><t>"Quotes"</t></si>
      </sst>
    `;
		const result = processSharedStrings(xml);

		expect(result.sharedStrings[0]).toBe("<si><t>&amp;&lt;&gt;</t></si>");
		expect(result.sharedIndexMap.get("<t>&amp;&lt;&gt;</t>")).toBe(0);
	});

	// it("should maintain correct indexes for duplicate content", () => {
	// 	const xml = `
  //     <sst>
  //       <si><t>Duplicate</t></si>
  //       <si><t>Duplicate</t></si>
  //       <si><t>Unique</t></si>
  //     </sst>
  //   `;
	// 	const result = processSharedStrings(xml);

	// 	expect(result.sharedStrings.length).toBe(3);
	// 	expect(result.sharedIndexMap.get("<t>Duplicate</t>")).toBe(0); // First occurrence
	// 	expect(result.sharedIndexMap.get("<t>Unique</t>")).toBe(2);
	// });

	it("should handle XML with namespace declarations", () => {
		const xml = `
      <?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <si><t>Namespace</t></si>
      </sst>
    `;
		const result = processSharedStrings(xml);

		expect(result.sharedStringsHeader).toBe("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
		expect(result.sharedStrings[0]).toBe("<si><t>Namespace</t></si>");
	});
});
