import { Transform } from "node:stream";

import { CRC32_TABLE } from "../constants.js";

/**
 * Computes the CRC-32 checksum for the given byte, using the standard IEEE 802.3 polynomial.
 * This is a low-level function that is used by the crc32Stream() function.
 *
 * @param {number} byte - The byte (0-255) to add to the checksum.
 * @param {number} crc - The current checksum value to update.
 * @returns {number} - The new checksum value.
 */
function crc32(byte: number, crc: number = 0xffffffff): number {
	return CRC32_TABLE[(crc ^ byte) & 0xff] as number ^ (crc >>> 8);
}

/**
 * Creates a Transform stream that computes the CRC-32 checksum of the input data.
 *
 * The `digest()` method can be called on the returned stream to get the final checksum value.
 *
 * @returns {Transform & { digest: () => number }} - The Transform stream.
 */
export function crc32Stream(): Transform & { digest: () => number } {
	let crc = 0xffffffff;

	const transform: Transform & { digest?: () => number } = new Transform({
		final(callback) {
			callback();
		},
		flush(callback) {
			callback();
		},
		transform(chunk, _encoding, callback) {
			for (let i = 0; i < chunk.length; i++) {
				crc = crc32(chunk[i], crc);
			}
			callback(null, chunk);
		},
	});

	transform.digest = () => (crc ^ 0xffffffff) >>> 0;

	return transform as Transform & { digest: () => number };
}
