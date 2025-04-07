import { Buffer } from "node:buffer";
import { deflateRawSync } from "node:zlib";

import { crc32, dosTime, toBytes } from "./utils.js";

import {
	CENTRAL_DIR_HEADER_SIG,
	END_OF_CENTRAL_DIR_SIG,
	LOCAL_FILE_HEADER_SIG,
} from "./constants.js";

/**
 * Creates a ZIP archive from a collection of files.
 *
 * @param {Object.<string, Buffer|string>} files - An object with file paths as keys and either Buffer or string content as values.
 * @returns {Buffer} - The ZIP archive as a Buffer.
 */
export function createSync(files: { [path: string]: Buffer | string }): Buffer {
	const fileEntries: Buffer[] = [];
	const centralDirectory: Buffer[] = [];

	let offset = 0;

	for (const [filename, rawContent] of Object.entries(files).sort(([a], [b]) => a.localeCompare(b))) {
		if (filename.includes("..")) {
			throw new Error(`Invalid filename: ${filename}`);
		}

		const content = Buffer.isBuffer(rawContent) ? rawContent : Buffer.from(rawContent);
		const fileNameBuf = Buffer.from(filename, "utf8");
		const modTime = dosTime(new Date());

		const crc = crc32(content);
		const compressed = deflateRawSync(content);
		const compSize = compressed.length;
		const uncompSize = content.length;

		// Local file header
		const localHeader = Buffer.concat([
			LOCAL_FILE_HEADER_SIG,
			toBytes(20, 2),
			toBytes(0, 2),
			toBytes(8, 2),
			modTime,
			toBytes(crc, 4),
			toBytes(compSize, 4),
			toBytes(uncompSize, 4),
			toBytes(fileNameBuf.length, 2),
			toBytes(0, 2),
		]);

		const localEntry = Buffer.concat([
			localHeader,
			fileNameBuf,
			compressed,
		]);

		fileEntries.push(localEntry);

		const centralEntry = Buffer.concat([
			Buffer.from(CENTRAL_DIR_HEADER_SIG),
			Buffer.from(toBytes(20, 2)), // Version made by
			Buffer.from(toBytes(20, 2)), // Version needed
			Buffer.from(toBytes(0, 2)),  // Flags
			Buffer.from(toBytes(8, 2)),  // Compression
			Buffer.from(modTime),
			Buffer.from(toBytes(crc, 4)),
			Buffer.from(toBytes(compSize, 4)),
			Buffer.from(toBytes(uncompSize, 4)),
			Buffer.from(toBytes(fileNameBuf.length, 2)),
			Buffer.from(toBytes(0, 2)),  // Extra field length
			Buffer.from(toBytes(0, 2)),  // Comment length
			Buffer.from(toBytes(0, 2)),  // Disk start
			Buffer.from(toBytes(0, 2)),  // Internal attrs
			Buffer.from(toBytes(0, 4)),  // External attrs
			Buffer.from(toBytes(offset, 4)),
			fileNameBuf,
		]);

		centralDirectory.push(centralEntry);
		offset += localEntry.length;
	}

	const centralDirSize = centralDirectory.reduce((sum, entry) => sum + entry.length, 0);
	const centralDirOffset = offset;

	const endRecord = Buffer.concat([
		Buffer.from(END_OF_CENTRAL_DIR_SIG),
		Buffer.from(toBytes(0, 2)), // Disk #
		Buffer.from(toBytes(0, 2)), // Start disk #
		Buffer.from(toBytes(centralDirectory.length, 2)),
		Buffer.from(toBytes(centralDirectory.length, 2)),
		Buffer.from(toBytes(centralDirSize, 4)),
		Buffer.from(toBytes(centralDirOffset, 4)),
		Buffer.from(toBytes(0, 2)), // Comment length
	]);

	return Buffer.concat(
		fileEntries.concat(centralDirectory).concat([endRecord]),
	);
}
