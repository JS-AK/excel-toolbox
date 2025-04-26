import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import * as Zip from "../../lib/zip/index.js";
import { TemplateFs } from "../../lib/template/template-fs.js";
import { validateWorksheetXml } from "../../lib/template/utils/validate-worksheet-xml.js";

const TEMP_DIR = path.resolve(process.cwd(), "src", "test", "template-fs", "temp");
const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "template-fs", "assets");
const INPUT_FILE = path.resolve(ASSETS_DIR, "input-test-01.xlsx");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-01.xlsx");

describe("TemplateFs integration test", () => {
	it("should copySheet with save", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		// Copy the sheet1 to sheet2
		await template.copySheet("Sheet1", "Sheet2");

		// Save the buffer to a file
		const buffer = await template.save();

		// Check that the buffer is defined
		expect(buffer).toBeDefined();

		// Save the buffer to a file
		await fs.writeFile(OUTPUT_FILE, buffer);

		// Read the original and rebuilt files
		const original = await fs.readFile(INPUT_FILE);
		const rebuilt = await fs.readFile(OUTPUT_FILE);

		// Read the original and rebuilt zip files
		const rebuiltZip = await Zip.read(rebuilt);
		const originalZip = await Zip.read(original);

		// Check that the original and rebuilt zip files have the same keys
		const origKeys = Object.keys(originalZip).sort();
		const rebuiltKeys = Object.keys(rebuiltZip).sort();

		const updatedKeys = [
			...origKeys,
			"xl/worksheets/sheet2.xml", // new sheet added by copy
		];

		// Check that the rebuilt zip file has the same keys as the original zip file
		expect(rebuiltKeys).toEqual(updatedKeys);

		// Check that the sheet1.xml file is the same in the original and rebuilt zip files
		const sheet1Original = originalZip["xl/worksheets/sheet1.xml"].toString();
		const sheet1Rebuilt = rebuiltZip["xl/worksheets/sheet1.xml"].toString();

		// Check that the sheet2.xml file is the same as the sheet1.xml file in the rebuilt zip file
		const sheet2Rebuilt = rebuiltZip["xl/worksheets/sheet2.xml"].toString();

		// Check that the sheet1.xml file is the same in the original and rebuilt zip files
		expect(sheet1Original).toEqual(sheet1Rebuilt);

		// Check that the sheet2.xml file is the same as the sheet1.xml file in the rebuilt zip file
		expect(sheet2Rebuilt).toEqual(sheet1Rebuilt);

		const validationResult1 = validateWorksheetXml(sheet1Rebuilt);
		const validationResult2 = validateWorksheetXml(sheet2Rebuilt);

		expect(validationResult1.isValid).toBe(true);
		expect(validationResult2.isValid).toBe(true);
	});

	it("should copySheet with saveStream", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		// Copy the sheet1 to sheet2
		await template.copySheet("Sheet1", "Sheet2");

		// Save the buffer to a file
		await template.saveStream(fsSync.createWriteStream(OUTPUT_FILE));

		const buffer = await fs.readFile(OUTPUT_FILE);

		// Check that the buffer is defined
		expect(buffer).toBeDefined();

		// Read the original and rebuilt files
		const original = await fs.readFile(INPUT_FILE);
		const rebuilt = await fs.readFile(OUTPUT_FILE);

		// Read the original and rebuilt zip files
		const rebuiltZip = await Zip.read(rebuilt);
		const originalZip = await Zip.read(original);

		// Check that the original and rebuilt zip files have the same keys
		const origKeys = Object.keys(originalZip).sort();
		const rebuiltKeys = Object.keys(rebuiltZip).sort();

		const updatedKeys = [
			...origKeys,
			"xl/worksheets/sheet2.xml", // new sheet added by copy
		];

		// Check that the rebuilt zip file has the same keys as the original zip file
		expect(rebuiltKeys).toEqual(updatedKeys);

		// Check that the sheet1.xml file is the same in the original and rebuilt zip files
		const sheet1Original = originalZip["xl/worksheets/sheet1.xml"].toString();
		const sheet1Rebuilt = rebuiltZip["xl/worksheets/sheet1.xml"].toString();

		// Check that the sheet2.xml file is the same as the sheet1.xml file in the rebuilt zip file
		const sheet2Rebuilt = rebuiltZip["xl/worksheets/sheet2.xml"].toString();

		// Check that the sheet1.xml file is the same in the original and rebuilt zip files
		expect(sheet1Original).toEqual(sheet1Rebuilt);

		// Check that the sheet2.xml file is the same as the sheet1.xml file in the rebuilt zip file
		expect(sheet2Rebuilt).toEqual(sheet1Rebuilt);

		const validationResult1 = validateWorksheetXml(sheet1Rebuilt);
		const validationResult2 = validateWorksheetXml(sheet2Rebuilt);

		expect(validationResult1.isValid).toBe(true);
		expect(validationResult2.isValid).toBe(true);
	});
});
