import util from "node:util";
import zlib from "node:zlib";

import * as Utils from "./utils/index.js";

const inflateRaw = util.promisify(zlib.inflateRaw);

/**
 * Parses a ZIP archive from a buffer and extracts the files within.
 *
 * @param {Buffer} buffer - The buffer containing the ZIP archive data.
 * @returns {Object.<string, Buffer>} - An object where keys are file names and values are file contents as Buffers.
 * @throws {Error} - Throws an error if an unsupported compression method is encountered or if decompression fails.
 */
export async function read(buffer: Buffer): Promise<Record<string, Buffer>> {
	const files: Record<string, Buffer> = {};
	let offset = 0;

	while (offset + 30 <= buffer.length) {
		const signature = buffer.readUInt32LE(offset);
		if (signature !== 0x04034b50) break; // not a local file header

		const generalPurposeBitFlag = buffer.readUInt16LE(offset + 6);
		const compressionMethod = buffer.readUInt16LE(offset + 8);
		const fileNameLength = buffer.readUInt16LE(offset + 26);
		const extraFieldLength = buffer.readUInt16LE(offset + 28);
		const fileNameStart = offset + 30;
		const fileNameEnd = fileNameStart + fileNameLength;
		const fileName = buffer.subarray(fileNameStart, fileNameEnd).toString();

		const dataStart = fileNameEnd + extraFieldLength;

		const useDataDescriptor = (generalPurposeBitFlag & 0x08) !== 0;

		let compressedData: Buffer;
		let content: Buffer;

		try {
			if (useDataDescriptor) {
				const { compressedSize, offset: ddOffset } = Utils.findDataDescriptor(buffer, dataStart);
				compressedData = buffer.subarray(dataStart, dataStart + compressedSize);

				if (compressionMethod === 0) {
					content = compressedData;
				} else if (compressionMethod === 8) {
					content = await inflateRaw(compressedData);
				} else {
					throw new Error(`Unsupported compression method ${compressionMethod}`);
				}

				offset = ddOffset + 16; // Skip over data descriptor
			} else {
				const compressedSize = buffer.readUInt32LE(offset + 18);
				compressedData = buffer.subarray(dataStart, dataStart + compressedSize);

				if (compressionMethod === 0) {
					content = compressedData;
				} else if (compressionMethod === 8) {
					content = await inflateRaw(compressedData);
				} else {
					throw new Error(`Unsupported compression method ${compressionMethod}`);
				}

				offset = dataStart + compressedSize;
			}
		} catch (error) {
			const message = error instanceof Error ? error.message : "Unknown error";
			throw new Error(`Error unpacking file ${fileName}: ${message}`);
		}

		files[fileName] = content;
	}

	return files;
}
