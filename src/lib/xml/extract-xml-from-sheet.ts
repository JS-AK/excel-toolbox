import { inflateRaw } from "pako";

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
export function extractXmlFromSheet(buffer: Buffer): string {
	if (!buffer || buffer.length === 0) {
		throw new Error("Empty buffer provided");
	}

	let xml: string | undefined;

	// Check if the buffer starts with an XML declaration (<?xml)
	const startsWithXml = buffer.subarray(0, 5).toString("utf8").trim().startsWith("<?xml");

	if (startsWithXml) {
		// Case 1: Already uncompressed XML - convert directly to string
		xml = buffer.toString("utf8");
	} else {
		// Case 2: Attempt to decompress as raw deflate data
		try {
			const inflated = inflateRaw(buffer, { to: "string" });

			// Validate the decompressed content contains worksheet data
			if (inflated && inflated.includes("<sheetData")) {
				xml = inflated;
			} else {
				throw new Error("Decompressed data does not contain sheetData");
			}
		} catch (e) {
			console.error("Decompression failed:", e);
			// Continue to fallback attempt
		}
	}

	// Fallback: If no XML obtained yet, try direct UTF-8 conversion
	if (!xml) {
		xml = buffer.toString("utf8");
	}

	// Sanitize XML by removing control characters (except tab, newline, carriage return)
	// This handles potential corruption from binary data or encoding issues
	xml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

	return xml;
}
