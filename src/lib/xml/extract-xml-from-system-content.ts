import { inflateRaw } from 'pako';

/**
 * Extracts and decompresses XML content from Excel system files (e.g., workbook.xml, [Content_Types].xml).
 * Handles both compressed (raw DEFLATE) and uncompressed (plain XML) formats with comprehensive error handling.
 *
 * @param {Buffer} buffer - The file content to process, which may be:
 *                         - Raw XML text
 *                         - DEFLATE-compressed XML data (without zlib headers)
 * @param {string} name - The filename being processed (for error reporting)
 * @returns {string} - The extracted XML content as a sanitized UTF-8 string
 * @throws {Error} - With descriptive messages for various failure scenarios:
 *                  - Empty buffer
 *                  - Decompression failures
 *                  - Invalid XML content
 */
export const extractXmlFromSystemContent = (buffer: Buffer, name: string): string => {
	// Validate input buffer
	if (!buffer || buffer.length === 0) {
		throw new Error(`Empty data buffer provided for file ${name}`);
	}

	let xml: string;

	// Check for XML declaration in first 5 bytes (<?xml)
	const startsWithXml = buffer.subarray(0, 5).toString('utf8').trim().startsWith('<?xml');

	if (startsWithXml) {
		// Case 1: Already uncompressed XML - convert directly to string
		xml = buffer.toString('utf8');
	} else {
		// Case 2: Attempt DEFLATE decompression
		try {
			const inflated = inflateRaw(buffer, { to: 'string' });

			// Validate decompressed content contains XML declaration
			if (inflated && inflated.includes('<?xml')) {
				xml = inflated;
			} else {
				throw new Error(`Decompressed data doesn't contain valid XML in ${name}`);
			}
		} catch (error) {
			const message = error instanceof Error ? error.message : 'Unknown error';

			throw new Error(`Failed to decompress ${name}: ${message}`);
		}
	}

	// Sanitize XML by removing illegal control characters (per XML 1.0 spec)
	// Preserves tabs (0x09), newlines (0x0A), and carriage returns (0x0D)
	xml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

	return xml;
};
