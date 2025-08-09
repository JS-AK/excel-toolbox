import { Buffer } from "node:buffer";

/**
 * Converts a numeric value into a fixed-length Buffer representation,
 * storing the value in little-endian format with right-padding of zeros.
 *
 * This is particularly useful for binary protocols or file formats that
 * require fixed-width numeric fields.
 *
 * @param {number} value - The numeric value to convert to bytes.
 *        Note: JavaScript numbers are IEEE 754 doubles, but only the
 *        integer portion will be used (up to 53-bit precision).
 * @param {number} len - The desired length of the output Buffer in bytes.
 *        Must be a positive integer.
 * @returns {Buffer} - A new Buffer of exactly `len` bytes containing:
 *        1. The value's bytes in little-endian order (least significant byte first)
 *        2. Zero padding in any remaining higher-order bytes
 *
 * @throws {RangeError} - If the length is not positive or the value is negative.
 */
export function toBytes(value: number, len: number): Buffer {
	if (len <= 0) throw new RangeError("Length must be a positive integer");
	if (value < 0) throw new RangeError("Negative values are not supported");
	if (!Number.isSafeInteger(value)) throw new RangeError("Value must be a safe integer");

	// Use BigInt to correctly handle 53-bit values
	let bigint = BigInt(value);
	const buf = Buffer.alloc(len);

	for (let i = 0; i < len; i++) {
		// Extract the least significant byte and assign to buffer
		buf[i] = Number(bigint & 0xffn);
		bigint >>= 8n;

		// Stop early if all remaining bits are zero
		if (bigint === 0n) break;
	}

	// If bigint is still non-zero, it means we overflowed the buffer length
	if (bigint > 0n) throw new RangeError("Value exceeds the maximum size for the specified length");

	return buf;
}
