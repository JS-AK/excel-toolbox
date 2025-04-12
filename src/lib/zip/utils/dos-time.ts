import { Buffer } from "node:buffer";

import { toBytes } from "./to-bytes.js";

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
