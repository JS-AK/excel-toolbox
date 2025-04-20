import { describe, expect, it } from "vitest";

import { extractXmlDeclaration } from "./extract-xml-declaration.js";

describe("extractXmlDeclaration", () => {
	it("should extract basic XML declaration", () => {
		const xml = "<?xml version=\"1.0\"?><root></root>";
		expect(extractXmlDeclaration(xml)).toBe("<?xml version=\"1.0\"?>");
	});

	it("should extract declaration with encoding", () => {
		const xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><root></root>";
		expect(extractXmlDeclaration(xml)).toBe("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
	});

	it("should extract declaration with standalone", () => {
		const xml = "<?xml version=\"1.0\" standalone=\"yes\"?><root></root>";
		expect(extractXmlDeclaration(xml)).toBe("<?xml version=\"1.0\" standalone=\"yes\"?>");
	});

	it("should extract declaration with all attributes", () => {
		const xml = "<?xml version=\"1.1\" encoding=\"ISO-8859-1\" standalone=\"no\"?><root></root>";
		expect(extractXmlDeclaration(xml)).toBe(
			"<?xml version=\"1.1\" encoding=\"ISO-8859-1\" standalone=\"no\"?>",
		);
	});

	it("should handle declaration with extra whitespace", () => {
		const xml = "<?xml   version = \"1.0\"   encoding = \"UTF-8\"   ?><root></root>";
		expect(extractXmlDeclaration(xml)).toBe(
			"<?xml   version = \"1.0\"   encoding = \"UTF-8\"   ?>",
		);
	});

	it("should return null for XML without declaration", () => {
		const xml = "<root></root>";
		expect(extractXmlDeclaration(xml)).toBeNull();
	});

	it("should return null for empty string", () => {
		expect(extractXmlDeclaration("")).toBeNull();
	});

	it("should handle malformed declarations", () => {
		const malformed1 = "<?xml version=\"1.0\"? some text><root></root>";
		expect(extractXmlDeclaration(malformed1)).toBeNull();;

		const malformed2 = "<?xml version=\"1.0\"><root></root>";
		expect(extractXmlDeclaration(malformed2)).toBeNull();
	});

	it("should only match declaration at start of string", () => {
		const xml = " <!-- comment --> <?xml version=\"1.0\"?><root></root>";
		expect(extractXmlDeclaration(xml)).toBeNull();
	});

	it("should handle complex XML with comments and processing instructions", () => {
		const xml = `<?xml version="1.0"?>
    <!-- comment -->
    <?processing-instruction?>
    <root></root>`;
		expect(extractXmlDeclaration(xml)).toBe("<?xml version=\"1.0\"?>");
	});
});
