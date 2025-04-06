import { Buffer } from "node:buffer";

/**
 * ZIP file signature constants in Buffer format.
 * These magic numbers identify different sections of a ZIP file,
 * as specified in PKWARE's APPNOTE.TXT (ZIP File Format Specification).
 */

/**
 * Central Directory Header signature (0x504b0102).
 * Marks an entry in the central directory, which contains metadata
 * about all files in the archive.
 * Format: 'PK\01\02'
 * Found in the central directory that appears at the end of the ZIP file.
 */
export const CENTRAL_DIR_HEADER_SIG = Buffer.from("504b0102", "hex");

/**
 * End of Central Directory Record signature (0x504b0506).
 * Marks the end of the central directory and contains global information
 * about the ZIP archive.
 * Format: 'PK\05\06'
 * This is the last record in a valid ZIP file.
 */
export const END_OF_CENTRAL_DIR_SIG = Buffer.from("504b0506", "hex");

/**
 * Local File Header signature (0x504b0304).
 * Marks the beginning of a file entry within the ZIP archive.
 * Format: 'PK\03\04' (ASCII letters PK followed by version numbers)
 * Appears before each file's compressed data.
 */
export const LOCAL_FILE_HEADER_SIG = Buffer.from("504b0304", "hex");
