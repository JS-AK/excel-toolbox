import { describe, expect, it } from "vitest";

import { dateToExcelSerial } from "./date-to-excel-serial.js";

describe("dateToExcelSerial", () => {
	it("should convert a known date to correct Excel serial number", () => {
		// January 1, 1900 should be serial number 2
		const date = new Date(1900, 0, 1); // Month is 0-indexed
		const serial = dateToExcelSerial(date);
		// Excel incorrectly treats 1900 as a leap year, so January 1, 1900 is serial number 2 in most implementations
		expect(serial).toBeCloseTo(2, 5);
	});

	it("should convert December 30, 1899 to serial number 0", () => {
		// December 30, 1899 should be serial number 0 (Excel day 0)
		const date = new Date(1899, 11, 30); // Month is 0-indexed
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(0, 5);
	});

	it("should convert a modern date correctly", () => {
		// January 1, 2023 should be around serial number 44927
		const date = new Date(2023, 0, 1);
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(44927, 0);
	});

	it("should handle dates with time components", () => {
		// January 1, 1900 at 12:00 PM should be 2.5
		const date = new Date(1900, 0, 1, 12, 0, 0);
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(2.5, 5);
	});

	it("should handle dates with milliseconds", () => {
		// January 1, 1900 at 6:00:00.000 AM should be 2.25
		const date = new Date(1900, 0, 1, 6, 0, 0, 0);
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(2.25, 5);
	});

	it("should handle leap year dates", () => {
		// February 29, 2020 (leap year)
		const date = new Date(2020, 1, 29);
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(43890, 0);
	});

	it("should handle dates before 1900", () => {
		// December 29, 1899 should be negative
		const date = new Date(1899, 11, 29);
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(-1, 5);
	});

	it("should handle dates far in the future", () => {
		// January 1, 2100
		const date = new Date(2100, 0, 1);
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(73051, 0);
	});

	it("should be consistent with multiple calls", () => {
		const date = new Date(2023, 5, 15, 14, 30, 45);
		const serial1 = dateToExcelSerial(date);
		const serial2 = dateToExcelSerial(date);
		expect(serial1).toBe(serial2);
	});

	it("should handle edge case of exactly Excel epoch", () => {
		// December 30, 1899 00:00:00 UTC
		const date = new Date(Date.UTC(1899, 11, 30));
		const serial = dateToExcelSerial(date);
		expect(serial).toBeCloseTo(0, 5);
	});
});
