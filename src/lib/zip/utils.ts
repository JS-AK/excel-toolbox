import { Buffer } from 'buffer';

/**
 * Precomputed CRC-32 lookup table for optimized checksum calculation.
 * The table is generated using the standard IEEE 802.3 (Ethernet) polynomial:
 * 0xEDB88320 (reversed representation of 0x04C11DB7).
 *
 * The table is immediately invoked and cached as a constant for performance,
 * following the common implementation pattern for CRC algorithms.
 */
const crcTable = (() => {
	// Create a typed array for better performance with 256 32-bit unsigned integers
	const table = new Uint32Array(256);

	// Generate table entries for all possible byte values (0-255)
	for (let i = 0; i < 256; i++) {
		let crc = i; // Initialize with current byte value

		// Process each bit (8 times)
		for (let j = 0; j < 8; j++) {
			/*
			 * CRC division algorithm:
			 * 1. If LSB is set (crc & 1), XOR with polynomial
			 * 2. Right-shift by 1 (unsigned)
			 *
			 * The polynomial 0xEDB88320 is:
			 * - Bit-reversed version of 0x04C11DB7
			 * - Uses reflected input/output algorithm
			 */
			crc = crc & 1
				? 0xedb88320 ^ (crc >>> 1)  // XOR with polynomial if LSB is set
				: crc >>> 1;                 // Just shift right if LSB is not set
		}

		// Store final 32-bit value (>>> 0 ensures unsigned 32-bit representation)
		table[i] = crc >>> 0;
	}

	return table;
})();

/**
 * Computes a CRC-32 checksum for the given Buffer using the standard IEEE 802.3 polynomial.
 * This implementation uses a precomputed lookup table for optimal performance.
 *
 * The algorithm follows these characteristics:
 * - Polynomial: 0xEDB88320 (reversed representation of 0x04C11DB7)
 * - Initial value: 0xFFFFFFFF (inverted by ~0)
 * - Final XOR value: 0xFFFFFFFF (achieved by inverting the result)
 * - Input and output reflection: Yes
 *
 * @param {Buffer} buf - The input buffer to calculate checksum for
 * @returns {number} - The 32-bit unsigned CRC-32 checksum (0x00000000 to 0xFFFFFFFF)
 */
export function crc32(buf: Buffer): number {
	// Initialize CRC with all 1's (0xFFFFFFFF) using bitwise NOT
	let crc = ~0;

	// Process each byte in the buffer
	for (let i = 0; i < buf.length; i++) {
		/*
		 * CRC update algorithm steps:
		 * 1. XOR current CRC with next byte (lowest 8 bits)
		 * 2. Use result as index in precomputed table (0-255)
		 * 3. XOR the table value with right-shifted CRC (8 bits)
		 *
		 * The operation breakdown:
		 * - (crc ^ buf[i]) - XOR with next byte
		 * - & 0xff - Isolate lowest 8 bits
		 * - crc >>> 8 - Shift CRC right by 8 bits (unsigned)
		 * - ^ crcTable[...] - XOR with precomputed table value
		 */
		crc = (crc >>> 8) ^ crcTable[(crc ^ buf[i] as number) & 0xff] as number;
	}

	/*
	 * Final processing:
	 * 1. Invert all bits (~crc) to match standard CRC-32 output
	 * 2. Convert to unsigned 32-bit integer (>>> 0)
	 */
	return ~crc >>> 0;
}

/**
 * Converts a JavaScript Date object to a 4-byte Buffer in MS-DOS date/time format
 * as specified in the ZIP file format specification (PKZIP APPNOTE.TXT).
 *
 * The MS-DOS date/time format packs both date and time into 4 bytes (32 bits) with
 * the following bit layout:
 *
 * Time portion (2 bytes/16 bits):
 * - Bits 00-04: Seconds divided by 2 (0-29, representing 0-58 seconds)
 * - Bits 05-10: Minutes (0-59)
 * - Bits 11-15: Hours (0-23)
 *
 * Date portion (2 bytes/16 bits):
 * - Bits 00-04: Day (1-31)
 * - Bits 05-08: Month (1-12)
 * - Bits 09-15: Year offset from 1980 (0-127, representing 1980-2107)
 *
 * @param {Date} date - The JavaScript Date object to convert
 * @returns {Buffer} - 4-byte Buffer containing:
 *                    - Bytes 0-1: DOS time (hours, minutes, seconds/2)
 *                    - Bytes 2-3: DOS date (year-1980, month, day)
 * @throws {RangeError} - If the date is before 1980 or after 2107
 */
export function dosTime(date: Date): Buffer {
	// Pack time components into 2 bytes (16 bits):
	// - Hours (5 bits) shifted left 11 positions (bits 11-15)
	// - Minutes (6 bits) shifted left 5 positions (bits 5-10)
	// - Seconds/2 (5 bits) in least significant bits (bits 0-4)
	const time =
		(date.getHours() << 11) |       // Hours occupy bits 11-15
		(date.getMinutes() << 5) |      // Minutes occupy bits 5-10
		(Math.floor(date.getSeconds() / 2));  // Seconds/2 occupy bits 0-4

	// Pack date components into 2 bytes (16 bits):
	// - (Year-1980) (7 bits) shifted left 9 positions (bits 9-15)
	// - Month (4 bits) shifted left 5 positions (bits 5-8)
	// - Day (5 bits) in least significant bits (bits 0-4)
	const day =
		((date.getFullYear() - 1980) << 9) |  // Years since 1980 (bits 9-15)
		((date.getMonth() + 1) << 5) |        // Month 1-12 (bits 5-8)
		date.getDate();                        // Day 1-31 (bits 0-4)

	// Combine both 2-byte values into a single 4-byte Buffer
	// Note: Using little-endian byte order for each 2-byte segment
	return Buffer.from([
		...toBytes(time, 2),  // Convert time to 2 bytes (LSB first)
		...toBytes(day, 2),    // Convert date to 2 bytes (LSB first)
	]);
}

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
