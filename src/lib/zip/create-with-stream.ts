import * as path from "node:path";
import { PassThrough, Transform, Writable } from "node:stream";
import { createReadStream } from "node:fs";
import { pipeline } from "node:stream/promises";
import zlib from "node:zlib";

import { crc32Stream, dosTime, toBytes } from "./utils/index.js";

import {
	CENTRAL_DIR_HEADER_SIG,
	END_OF_CENTRAL_DIR_SIG,
	LOCAL_FILE_HEADER_SIG,
} from "./constants.js";

/**
 * Creates a ZIP archive from a collection of files, streaming the output to a provided writable stream.
 *
 * @param fileKeys - An array of file paths (relative to the destination) that will be used to create a new workbook.
 * @param destination - The path where the template files are located.
 * @param output - A Writable stream that the ZIP archive will be written to.
 *
 * @throws {Error} - If a file does not exist in the destination.
 * @throws {Error} - If a file is not readable.
 * @throws {Error} - If the writable stream emits an error.
 */
export async function createWithStream(fileKeys: string[], destination: string, output: Writable): Promise<void> {
	// Stores central directory records
	const centralDirectory: Buffer[] = [];

	// Tracks the current offset in the output stream
	let offset = 0;

	for (const filename of fileKeys.sort((a, b) => a.localeCompare(b))) {
		// Prevent directory traversal
		if (filename.includes("..")) {
			throw new Error(`Invalid filename: ${filename}`);
		}

		// Construct absolute path to the file
		const fullPath = path.join(destination, ...filename.split("/"));

		// Convert filename to UTF-8 buffer
		const fileNameBuf = Buffer.from(filename, "utf8");

		// Get modification time in DOS format
		const modTime = dosTime(new Date());

		// Read file as stream
		const source = createReadStream(fullPath);

		// Create CRC32 transform stream
		const crc32 = crc32Stream();

		// Create raw deflate stream (no zlib headers)
		const deflater = zlib.createDeflateRaw();

		// Uncompressed size counter
		let uncompSize = 0;

		// Compressed size counter
		let compSize = 0;

		// Store compressed output data
		const compressedChunks: Buffer[] = [];

		const sizeCounter = new Transform({
			transform(chunk, _enc, cb) {
				uncompSize += chunk.length;
				cb(null, chunk);
			},
		});

		const collectCompressed = new PassThrough();
		collectCompressed.on("data", chunk => {
			// Count compressed bytes
			compSize += chunk.length;

			// Save compressed chunk
			compressedChunks.push(chunk);
		});

		// Run all transforms in pipeline: read -> count size -> CRC -> deflate -> collect compressed
		await pipeline(
			source,
			sizeCounter,
			crc32,
			deflater,
			collectCompressed,
		);

		// Get final CRC32 value
		const crc = crc32.digest();

		// Concatenate all compressed chunks into a single buffer
		const compressed = Buffer.concat(compressedChunks);

		// Create local file header followed by compressed content
		const localHeader = Buffer.concat([
			LOCAL_FILE_HEADER_SIG,          // Local file header signature
			toBytes(20, 2),                 // Version needed to extract
			toBytes(0, 2),                  // General purpose bit flag
			toBytes(8, 2),                  // Compression method (deflate)
			modTime,                        // File modification time and date
			toBytes(crc, 4),                // CRC-32 checksum
			toBytes(compSize, 4),           // Compressed size
			toBytes(uncompSize, 4),         // Uncompressed size
			toBytes(fileNameBuf.length, 2), // Filename length
			toBytes(0, 2),                  // Extra field length
			fileNameBuf,                    // Filename
			compressed,                     // Compressed file data
		]);

		// Write local file header and data to output
		await new Promise<void>((resolve, reject) => {
			output.write(localHeader, err => err ? reject(err) : resolve());
		});

		// Create central directory entry for this file
		const centralEntry = Buffer.concat([
			CENTRAL_DIR_HEADER_SIG,         // Central directory file header signature
			toBytes(20, 2),                 // Version made by
			toBytes(20, 2),                 // Version needed to extract
			toBytes(0, 2),                  // General purpose bit flag
			toBytes(8, 2),                  // Compression method
			modTime,                        // File modification time and date
			toBytes(crc, 4),                // CRC-32 checksum
			toBytes(compSize, 4),           // Compressed size
			toBytes(uncompSize, 4),         // Uncompressed size
			toBytes(fileNameBuf.length, 2), // Filename length
			toBytes(0, 2),                  // Extra field length
			toBytes(0, 2),                  // File comment length
			toBytes(0, 2),                  // Disk number start
			toBytes(0, 2),                  // Internal file attributes
			toBytes(0, 4),                  // External file attributes
			toBytes(offset, 4),             // Offset of local header
			fileNameBuf,                    // Filename
		]);

		// Store for later
		centralDirectory.push(centralEntry);

		// Update offset after writing this entry
		offset += localHeader.length;
	}

	// Total size of central directory
	const centralDirSize = centralDirectory.reduce((sum, entry) => sum + entry.length, 0);

	// Start of central directory
	const centralDirOffset = offset;

	// Write each central directory entry to output
	for (const entry of centralDirectory) {
		await new Promise<void>((resolve, reject) => {
			output.write(entry, err => err ? reject(err) : resolve());
		});
	}

	// Create and write end of central directory record
	const endRecord = Buffer.concat([
		END_OF_CENTRAL_DIR_SIG,              // End of central directory signature
		toBytes(0, 2),                       // Number of this disk
		toBytes(0, 2),                       // Disk with start of central directory
		toBytes(centralDirectory.length, 2), // Total entries on this disk
		toBytes(centralDirectory.length, 2), // Total entries overall
		toBytes(centralDirSize, 4),          // Size of central directory
		toBytes(centralDirOffset, 4),        // Offset of start of central directory
		toBytes(0, 2),                       // ZIP file comment length
	]);

	await new Promise<void>((resolve, reject) => {
		output.write(endRecord, err => err ? reject(err) : resolve());
	});

	output.end();
}
