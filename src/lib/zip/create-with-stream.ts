import * as path from "node:path";
import { Transform, Writable } from "node:stream";
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

		// deflater.on("data", (chunk) => { console.log("deflater data path:", fullPath, "length:", chunk.length); });
		// deflater.on("finish", () => { console.log("deflater finished path:", fullPath, "uncompSize:", uncompSize, "compSize:", compSize); });
		// deflater.on("error", (err) => { console.log("deflater error path:", fullPath, "error:", err); });
		// deflater.on("close", () => { console.log("deflater closed path:", fullPath); });
		// deflater.on("pipe", (src) => { console.log("deflater pipe path:", fullPath); });
		// deflater.on("unpipe", (src) => { console.log("deflater unpipe path:", fullPath); });
		// deflater.on("drain", () => { console.log("deflater drain path:", fullPath); });
		// deflater.on("pause", () => { console.log("deflater pause path:", fullPath); });
		// deflater.on("resume", () => { console.log("deflater resume path:", fullPath); });
		// deflater.on("end", () => console.log("deflater ended, path:", fullPath));

		// source.on("data", (chunk) => { console.log("source data path:", fullPath, "length:", chunk.length); });
		// source.on("finish", () => { console.log("source finished path:", fullPath, "uncompSize:", uncompSize, "compSize:", compSize); });
		// source.on("error", (err) => { console.log("source error path:", fullPath, "error:", err); });
		// source.on("close", () => { console.log("source closed path:", fullPath); });
		// source.on("pipe", (src) => { console.log("source pipe path:", fullPath); });
		// source.on("unpipe", (src) => { console.log("source unpipe path:", fullPath); });
		// source.on("drain", () => { console.log("source drain path:", fullPath); });
		// source.on("pause", () => { console.log("source pause path:", fullPath); });
		// source.on("resume", () => { console.log("source resume path:", fullPath); });
		// source.on("end", () => console.log("source ended, path:", fullPath));

		// sizeCounter.on("data", (chunk) => { console.log("sizeCounter data path:", fullPath, "length:", chunk.length); });
		// sizeCounter.on("finish", () => { console.log("sizeCounter finished path:", fullPath, "uncompSize:", uncompSize, "compSize:", compSize); });
		// sizeCounter.on("error", (err) => { console.log("sizeCounter error path:", fullPath, "error:", err); });
		// sizeCounter.on("close", () => { console.log("sizeCounter closed path:", fullPath); });
		// sizeCounter.on("pipe", (src) => { console.log("sizeCounter pipe path:", fullPath); });
		// sizeCounter.on("unpipe", (src) => { console.log("sizeCounter unpipe path:", fullPath); });
		// sizeCounter.on("drain", () => { console.log("sizeCounter drain path:", fullPath); });
		// sizeCounter.on("pause", () => { console.log("sizeCounter pause path:", fullPath); });
		// sizeCounter.on("resume", () => { console.log("sizeCounter resume path:", fullPath); });
		// sizeCounter.on("end", () => console.log("sizeCounter ended, path:", fullPath));

		// crc32.on("data", (chunk) => { console.log("crc32 data path:", fullPath, "length:", chunk.length); });
		// crc32.on("finish", () => { console.log("crc32 finished path:", fullPath, "uncompSize:", uncompSize, "compSize:", compSize); });
		// crc32.on("error", (err) => { console.log("crc32 error path:", fullPath, "error:", err); });
		// crc32.on("close", () => { console.log("crc32 closed path:", fullPath); });
		// crc32.on("pipe", (src) => { console.log("crc32 pipe path:", fullPath); });
		// crc32.on("unpipe", (src) => { console.log("crc32 unpipe path:", fullPath); });
		// crc32.on("drain", () => { console.log("crc32 drain path:", fullPath); });
		// crc32.on("pause", () => { console.log("crc32 pause path:", fullPath); });
		// crc32.on("resume", () => { console.log("crc32 resume path:", fullPath); });
		// crc32.on("end", () => console.log("crc32 ended, path:", fullPath));

		collectCompressed.on("data", (chunk) => {/*  console.log("collectCompressed data path:", fullPath, "length:", chunk.length); */ });
		// collectCompressed.on("finish", () => { console.log("collectCompressed finished path:", fullPath, "uncompSize:", uncompSize, "compSize:", compSize); });
		// collectCompressed.on("error", (err) => { console.log("collectCompressed error path:", fullPath, "error:", err); });
		// collectCompressed.on("close", () => { console.log("collectCompressed closed path:", fullPath); });
		// collectCompressed.on("pipe", (src) => { console.log("collectCompressed pipe path:", fullPath); });
		// collectCompressed.on("unpipe", (src) => { console.log("collectCompressed unpipe path:", fullPath); });
		// collectCompressed.on("drain", () => { console.log("collectCompressed drain path:", fullPath); });
		// collectCompressed.on("pause", () => { console.log("collectCompressed pause path:", fullPath); });
		// collectCompressed.on("resume", () => { console.log("collectCompressed resume path:", fullPath); });
		// collectCompressed.on("end", () => console.log("collectCompressed ended, path:", fullPath));

		// deflater.on("readable", () => {
		// 	console.log("deflater readable path:", fullPath);
		// });

		await pipeline(
			source,
			sizeCounter,
			crc32,
			deflater,
			collectCompressed,
		);

		// await new Promise<void>((resolve, reject) => {
		// 	source
		// 		.pipe(sizeCounter)
		// 		.pipe(crc32)
		// 		.pipe(deflater)
		// 		.pipe(collectCompressed)
		// 		.on("finish", resolve)
		// 		.on("error", reject);

		// 	source.on("error", reject);
		// 	deflater.on("error", reject);
		// });

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
