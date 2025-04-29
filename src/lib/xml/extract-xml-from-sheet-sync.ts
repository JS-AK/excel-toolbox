import { inflateRawSync } from "node:zlib";

/**
 * Extracts and parses XML content from an Excel worksheet file (e.g., xl/worksheets/sheet1.xml).
 * Handles both compressed (raw deflate) and uncompressed (plain XML) formats.
 *
 * This function is designed to work with Excel Open XML (.xlsx) worksheet files,
 * which may be stored in either compressed or uncompressed format within the ZIP container.
 *
 * @param {Buffer} buffer - The file content to process, which may be:
 *                         - Raw XML text
 *                         - Deflate-compressed XML data (without zlib headers)
 * @returns {string} - The extracted XML content as a UTF-8 string
 * @throws {Error} - If the buffer is empty or cannot be processed
 */
export function extractXmlFromSheetSync(buffer: Buffer): string {
	if (!buffer || buffer.length === 0) {
		throw new Error("Empty buffer provided");
	}

	let xml: string;

	// Check if the buffer starts with an XML declaration (<?xml)
	const head = buffer.subarray(0, 1024).toString("utf8").replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "").trim();
	const isXml = /^<\?xml[\s\S]+<\w+[\s>]/.test(head);

	if (isXml) {
		// Case 1: Already uncompressed XML - convert directly to string
		xml = buffer.toString("utf8");
	} else {
		// Case 2: Attempt to decompress as raw deflate data
		try {
			xml = inflateRawSync(buffer).toString("utf8");
		} catch (err) {
			throw new Error("Failed to decompress sheet XML: " + (err instanceof Error ? err.message : String(err)));
		}
	}

	// Sanitize XML by removing control characters (except tab, newline, carriage return)
	// This handles potential corruption from binary data or encoding issues
	xml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

	return xml;
}
