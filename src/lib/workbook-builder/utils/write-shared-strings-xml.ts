import fs from "node:fs";
import fsPromises from "node:fs/promises";
import path from "node:path";

import { escapeXml } from "../../utils/index.js";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";

/**
 * Writes the `sharedStrings.xml` content to a file at the given destination.
 *
 * Uses a file write stream with backpressure control to avoid buffering large
 * content in memory.
 *
 * @param destination - Absolute or relative file path to write
 * @param strings - Array of unique strings used in the workbook
 * @returns Promise that resolves when the write stream finishes
 */
export async function writeSharedStringsXml(destination: string, strings: string[] = []): Promise<void> {
	// Ensure destination folder exists
	await fsPromises.mkdir(path.dirname(destination), { recursive: true });

	const stream = fs.createWriteStream(destination, { encoding: "utf-8" });

	try {
		// Document header
		stream.write(XML_DECLARATION + "\n");
		stream.write(`<sst xmlns="${XML_NAMESPACES.SPREADSHEET_ML}" count="${strings.length}" uniqueCount="${strings.length}">\n`);

		// Main string items
		for (const s of strings) {
			const preserve = /^\s|\s$/.test(s) ? " xml:space=\"preserve\"" : "";
			const siXml = `<si><t${preserve}>${escapeXml(s)}</t></si>\n`;

			// Write with backpressure control
			if (!stream.write(siXml)) {
				await new Promise<void>(resolve => stream.once("drain", () => resolve()));
			}
		}

		// Closing tag
		stream.write("</sst>");
	} finally {
		stream.end();
	}

	return new Promise((resolve, reject) => {
		stream.on("error", reject);
		stream.on("finish", resolve);
	});
}
