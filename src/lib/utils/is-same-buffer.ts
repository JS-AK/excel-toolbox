/**
 * Checks if two Buffers are the same
 * @param {Buffer} buf1 - the first Buffer
 * @param {Buffer} buf2 - the second Buffer
 * @returns {boolean} - true if the Buffers are the same, false otherwise
 */
export function isSameBuffer(buf1: Buffer, buf2: Buffer): boolean {
	return buf1.equals(buf2);
}
