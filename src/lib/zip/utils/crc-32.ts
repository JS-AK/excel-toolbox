import { Buffer } from "node:buffer";

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
