import { extractXmlDeclaration } from "./extract-xml-declaration.js";

/**
 * Processes the shared strings XML by extracting the XML declaration,
 * extracting individual <si> elements, and storing them in an array.
 *
 * The function returns an object with four properties:
 * - sharedIndexMap: A map of shared string content to their corresponding index
 * - sharedStrings: An array of shared strings
 * - sharedStringsHeader: The XML declaration of the shared strings
 * - sheetMergeCells: An empty array, which is only used for type compatibility
 *                    with the return type of processBuild.
 *
 * @param sharedStringsXml - The XML string of the shared strings
 * @returns An object with the four properties above
 */
export function processSharedStrings(sharedStringsXml: string) {
	// Array for storing shared strings
	const sharedStrings: string[] = [];
	const sharedStringsHeader = extractXmlDeclaration(sharedStringsXml);

	// Map for fast lookup of shared string index by content
	const sharedIndexMap = new Map<string, number>();

	// Regular expression for finding <si> elements (shared string items)
	const siRegex = /<si>([\s\S]*?)<\/si>/g;

	// Parse sharedStringsXml and fill sharedStrings and sharedIndexMap
	for (const match of sharedStringsXml.matchAll(siRegex)) {
		const content = match[1];

		if (!content) continue;

		const fullSi = `<si>${content}</si>`;
		sharedIndexMap.set(content, sharedStrings.length);
		sharedStrings.push(fullSi);
	}

	return {
		sharedIndexMap,
		sharedStrings,
		sharedStringsHeader,
	};
};
