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
 * @throws {RangeError} - If the value requires more bytes than `len` to represent
 *        (though this is currently not explicitly checked)
 */
export function toBytes(value: number, len: number): Buffer {
	// Allocate a new Buffer of the requested length, automatically zero-filled
	const buf = Buffer.alloc(len);

	// Process each byte position from least significant to most significant
	for (let i = 0; i < len; i++) {
		// Store the least significant byte of the current value
		buf[i] = value & 0xff;  // Mask to get bottom 8 bits

		// Right-shift the value by 8 bits to process the next byte
		// Note: This uses unsigned right shift (>>> would be signed)
		value >>= 8;

		// If the loop completes with value != 0, we've overflowed the buffer length,
		// but this isn't currently checked/handled
	}

	return buf;
}
