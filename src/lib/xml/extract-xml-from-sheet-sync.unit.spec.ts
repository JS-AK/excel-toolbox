import { describe, expect, it } from "vitest";
import { deflateRawSync } from "node:zlib";

import { extractXmlFromSheetSync } from "./extract-xml-from-sheet-sync.js";

describe("extractXmlFromSheet", () => {
	it("should handle empty buffer", () => {
		expect(() => extractXmlFromSheetSync(Buffer.alloc(0))).toThrow("Empty buffer provided");
	});

	it("should extract uncompressed XML", async () => {
		const xml = "<?xml version=\"1.0\"?><worksheet><sheetData></sheetData></worksheet>";
		const buffer = Buffer.from(xml);
		expect(extractXmlFromSheetSync(buffer)).toBe(xml);
	});

	it("returns plain XML from uncompressed buffer", async () => {
		const xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet><data>test</data></worksheet>";
		const buffer = Buffer.from(xml, "utf8");
		const result = extractXmlFromSheetSync(buffer);
		expect(result).toBe(xml);
	});

	it("should decompress and extract deflated XML", async () => {
		// This is a deflated version of: <?xml...><worksheet><sheetData></sheetData></worksheet>
		const deflated = Buffer.from([
			0x3c, 0x3f, 0x78, 0x6d, 0x6c, 0x20, 0x76, 0x65, 0x72, 0x73, 0x69, 0x6f,
			0x6e, 0x3d, 0x22, 0x31, 0x2e, 0x30, 0x22, 0x3f, 0x3e, 0x3c, 0x77, 0x6f,
			0x72, 0x6b, 0x73, 0x68, 0x65, 0x65, 0x74, 0x3e, 0x3c, 0x73, 0x68, 0x65,
			0x65, 0x74, 0x44, 0x61, 0x74, 0x61, 0x3e, 0x3c, 0x2f, 0x73, 0x68, 0x65,
			0x65, 0x74, 0x44, 0x61, 0x74, 0x61, 0x3e, 0x3c, 0x2f, 0x77, 0x6f, 0x72,
			0x6b, 0x73, 0x68, 0x65, 0x65, 0x74, 0x3e,
		]);

		const expected = "<?xml version=\"1.0\"?><worksheet><sheetData></sheetData></worksheet>";
		expect(extractXmlFromSheetSync(deflated)).toBe(expected);
	});

	it("should sanitize XML by removing control characters", async () => {
		const xml = "<?xml version=\"1.0\"?><worksheet><sheetData>\x00\x01\x02</sheetData></worksheet>";
		const expected = "<?xml version=\"1.0\"?><worksheet><sheetData></sheetData></worksheet>";
		expect(extractXmlFromSheetSync(Buffer.from(xml))).toBe(expected);
	});

	it("decompresses deflate-encoded XML buffer", async () => {
		const xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet><value>42</value></worksheet>";
		const compressed = deflateRawSync(Buffer.from(xml, "utf8"));
		const result = extractXmlFromSheetSync(compressed);
		expect(result).toBe(xml);
	});

	it("throws on invalid non-XML and non-deflate data", async () => {
		const garbage = Buffer.from([0xde, 0xad, 0xbe, 0xef]);
		expect(() => extractXmlFromSheetSync(garbage)).toThrow(/Failed to decompress sheet XML/);
	});

	it("sanitizes control characters from XML", async () => {
		const xml = "<?xml version=\"1.0\"?><ws>\x01valid</ws>";
		const buffer = Buffer.from(xml, "utf8");
		const result = extractXmlFromSheetSync(buffer);
		expect(result).toBe("<?xml version=\"1.0\"?><ws>valid</ws>");
	});
});
