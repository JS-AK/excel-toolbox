/**
 * Finds a Data Descriptor in a ZIP archive buffer.
 *
 * The Data Descriptor is an optional 16-byte structure that appears at the end of a file's compressed data.
 * It contains the compressed size of the file, and must be used when the Local File Header does not contain this information.
 *
 * @param buffer - The buffer containing the ZIP archive data.
 * @param start - The starting offset in the buffer to search for the Data Descriptor.
 * @returns - An object with `offset` and `compressedSize` properties.
 * @throws {Error} - If the Data Descriptor is not found.
 */
export function findDataDescriptor(
	buffer: Buffer,
	start: number,
): { offset: number; compressedSize: number } {
	const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
	const DATA_DESCRIPTOR_TOTAL_LENGTH = 16;
	const COMPRESSED_SIZE_OFFSET_FROM_SIGNATURE = 8;

	for (let i = start; i <= buffer.length - DATA_DESCRIPTOR_TOTAL_LENGTH; i++) {
		if (buffer.readUInt32LE(i) === DATA_DESCRIPTOR_SIGNATURE) {
			const compressedSize = buffer.readUInt32LE(i + COMPRESSED_SIZE_OFFSET_FROM_SIGNATURE);
			return {
				compressedSize,
				offset: i,
			};
		}
	}

	throw new Error("Data Descriptor not found");
}
