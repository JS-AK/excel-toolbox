/**
 * Extracts the XML declaration from a given XML string.
 *
 * The XML declaration is a string that looks like `<?xml ...?>` and is usually
 * present at the beginning of an XML file. It contains information about the
 * XML version, encoding, and standalone status.
 *
 * This function returns `null` if the input string does not have a valid XML
 * declaration.
 *
 * @param xmlString - The XML string to extract the declaration from.
 * @returns The extracted XML declaration string, or `null`.
 */
export function extractXmlDeclaration(xmlString: string): string | null {
	// const declarationRegex = /^<\?xml\s+[^?]+\?>/;
	const declarationRegex = /^<\?xml\s+version\s*=\s*["'][^"']+["'](\s+(encoding|standalone)\s*=\s*["'][^"']+["'])*\s*\?>/;

	const match = xmlString.trim().match(declarationRegex);

	return match ? match[0] : null;
}
