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
	const centralDirectory: Buffer[] = [];
	let offset = 0;

	for (const filename of fileKeys.sort((a, b) => a.localeCompare(b))) {
		if (filename.includes("..")) {
			throw new Error(`Invalid filename: ${filename}`);
		}

		const fullPath = path.join(destination, ...filename.split("/"));
		const fileNameBuf = Buffer.from(filename, "utf8");
		const modTime = dosTime(new Date());

		const source = createReadStream(fullPath);
		const crc32 = crc32Stream();
		const deflater = zlib.createDeflateRaw();
		let uncompSize = 0;
		let compSize = 0;
		const compressedChunks: Buffer[] = [];

		const sizeCounter = new Transform({
			transform(chunk, _enc, cb) {
				uncompSize += chunk.length;
				cb(null, chunk);
			},
		});

		const collectCompressed = new Transform({
			transform(chunk, _enc, cb) {
				compressedChunks.push(chunk);
				compSize += chunk.length;
				cb(null, chunk);
			},
		});

		await pipeline(
			source,
			sizeCounter,
			crc32,
			deflater,
			collectCompressed,
			new PassThrough(),
		);

		const crc = crc32.digest();
		const compressed = Buffer.concat(compressedChunks);

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
			fileNameBuf,
			compressed,
		]);

		await new Promise<void>((resolve, reject) => {
			output.write(localHeader, err => err ? reject(err) : resolve());
		});

		const centralEntry = Buffer.concat([
			CENTRAL_DIR_HEADER_SIG,
			toBytes(20, 2),
			toBytes(20, 2),
			toBytes(0, 2),
			toBytes(8, 2),
			modTime,
			toBytes(crc, 4),
			toBytes(compSize, 4),
			toBytes(uncompSize, 4),
			toBytes(fileNameBuf.length, 2),
			toBytes(0, 2),
			toBytes(0, 2),
			toBytes(0, 2),
			toBytes(0, 2),
			toBytes(0, 4),
			toBytes(offset, 4),
			fileNameBuf,
		]);

		centralDirectory.push(centralEntry);
		offset += localHeader.length;
	}

	const centralDirSize = centralDirectory.reduce((sum, entry) => sum + entry.length, 0);
	const centralDirOffset = offset;

	for (const entry of centralDirectory) {
		await new Promise<void>((resolve, reject) => {
			output.write(entry, err => err ? reject(err) : resolve());
		});
	}

	const endRecord = Buffer.concat([
		END_OF_CENTRAL_DIR_SIG,
		toBytes(0, 2),
		toBytes(0, 2),
		toBytes(centralDirectory.length, 2),
		toBytes(centralDirectory.length, 2),
		toBytes(centralDirSize, 4),
		toBytes(centralDirOffset, 4),
		toBytes(0, 2),
	]);

	await new Promise<void>((resolve, reject) => {
		output.write(endRecord, err => err ? reject(err) : resolve());
	});

	output.end();
}
