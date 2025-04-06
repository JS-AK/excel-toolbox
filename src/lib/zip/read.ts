import { inflateRawSync } from 'node:zlib';

/**
 * Parses a ZIP archive from a buffer and extracts the files within.
 *
 * @param {Buffer} buffer - The buffer containing the ZIP archive data.
 * @returns {Object.<string, string>} - An object where keys are file names and values are file contents.
 * @throws {Error} - Throws an error if an unsupported compression method is encountered or if decompression fails.
 */

export function read(buffer: Buffer): { [s: string]: string; } {
	const files: { [s: string]: string; } = {};
	let offset = 0;

	while (offset + 4 <= buffer.length) {
		const signature = buffer.readUInt32LE(offset);
		if (signature !== 0x04034b50) break;

		const compressionMethod = buffer.readUInt16LE(offset + 8);
		const fileNameLength = buffer.readUInt16LE(offset + 26);
		const extraLength = buffer.readUInt16LE(offset + 28);
		const fileNameStart = offset + 30;
		const fileNameEnd = fileNameStart + fileNameLength;
		const fileName = buffer.subarray(fileNameStart, fileNameEnd).toString();
		const dataStart = fileNameEnd + extraLength;

		let nextOffset = dataStart;
		while (nextOffset + 4 <= buffer.length) {
			if (buffer.readUInt32LE(nextOffset) === 0x04034b50) break;
			nextOffset++;
		}
		if (nextOffset + 4 > buffer.length) {
			nextOffset = buffer.length;
		}

		const compressedData = buffer.subarray(dataStart, nextOffset);
		let content = '';

		try {
			if (compressionMethod === 0) {
				content = compressedData.toString();
			} else if (compressionMethod === 8) {
				content = inflateRawSync(new Uint8Array(compressedData)).toString();
			} else {
				throw new Error(`Unsupported compression method ${compressionMethod}`);
			}
		} catch (error) {
			const message = error instanceof Error ? error.message : 'Unknown error';

			throw new Error(`Error unpacking file ${fileName}: ${message}`);
		}

		files[fileName] = content;
		offset = nextOffset;
	}

	return files;
}
