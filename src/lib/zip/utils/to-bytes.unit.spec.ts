import { describe, expect, it } from "vitest";
import { toBytes } from "./to-bytes.js";

// Test suite for toBytes function
describe("toBytes", () => {
	it("should convert a small number to a 1-byte buffer", () => {
		expect(toBytes(42, 1)).toEqual(Buffer.from([42]));
	});

	it("should convert a small number to a multi-byte buffer with padding", () => {
		expect(toBytes(42, 4)).toEqual(Buffer.from([42, 0, 0, 0]));
	});

	it("should handle multi-byte numbers correctly", () => {
		expect(toBytes(4660, 4)).toEqual(Buffer.from([0x34, 0x12, 0, 0]));
		expect(toBytes(0x12345678, 4)).toEqual(Buffer.from([0x78, 0x56, 0x34, 0x12]));
	});

	it("should handle maximum safe integer correctly", () => {
		expect(toBytes(Number.MAX_SAFE_INTEGER, 8)).toEqual(Buffer.from([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x1f, 0x00]));
	});

	it("should handle zero correctly", () => {
		expect(toBytes(0, 4)).toEqual(Buffer.from([0, 0, 0, 0]));
	});

	it("should return a zero-filled buffer for zero with a larger length", () => {
		expect(toBytes(0, 8)).toEqual(Buffer.alloc(8));
	});

	it("should throw a RangeError for negative lengths", () => {
		expect(() => toBytes(123, -1)).toThrow(RangeError);
	});

	it("should throw a RangeError for zero length", () => {
		expect(() => toBytes(123, 0)).toThrow(RangeError);
	});

	it("should throw a RangeError for negative values", () => {
		expect(() => toBytes(-123, 4)).toThrow(RangeError);
	});

	it("should throw a RangeError for non-safe integers", () => {
		expect(() => toBytes(Number.MAX_SAFE_INTEGER + 1, 8)).toThrow(RangeError);
	});

	it("should handle maximum value fitting in the given length", () => {
		expect(toBytes(255, 1)).toEqual(Buffer.from([0xff]));
		expect(toBytes(65535, 2)).toEqual(Buffer.from([0xff, 0xff]));
		expect(toBytes(4294967295, 4)).toEqual(Buffer.from([0xff, 0xff, 0xff, 0xff]));
		expect(toBytes(Number.MAX_SAFE_INTEGER, 8)).toEqual(Buffer.from([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x1f, 0x00]));
	});

	it("should throw a RangeError if the value exceeds the buffer length", () => {
		expect(() => toBytes(256, 1)).toThrow(RangeError);
		expect(() => toBytes(65536, 2)).toThrow(RangeError);
		expect(() => toBytes(4294967296, 4)).toThrow(RangeError);
		expect(() => toBytes(Number.MAX_SAFE_INTEGER + 1, 8)).toThrow(RangeError);
	});
});
