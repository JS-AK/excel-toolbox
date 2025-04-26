import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import * as Xml from "../../lib/xml/index.js";
import * as Zip from "../../lib/zip/index.js";
import { TemplateFs } from "../../lib/template/template-fs.js";
import { validateWorksheetXml } from "../../lib/template/utils/validate-worksheet-xml.js";

const TEMP_DIR = path.resolve(process.cwd(), "src", "test", "template-fs", "temp");
const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "template-fs", "assets");
const INPUT_FILE = path.resolve(ASSETS_DIR, "input-test-02.xlsx");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-02.xlsx");

describe("TemplateFs integration test", () => {
	it("should insertRows with save", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		await template.insertRows({
			rows: [
				["Name", "Age", "City"],
				["John", 30, "New York"],
				["Jane", 25, "Los Angeles"],
			],
			sheetName: "Sheet1",
			startRowNumber: 2,
		});

		// Save the buffer to a file
		const buffer = await template.save();

		// Save the buffer to a file
		await fs.writeFile(OUTPUT_FILE, buffer);

		expect(buffer).toBeDefined();

		// Read the rebuilt file
		const original = await fs.readFile(INPUT_FILE);
		const rebuilt = await fs.readFile(OUTPUT_FILE);

		// Read the rebuilt zip file
		const rebuiltZip = await Zip.read(rebuilt);
		const originalZip = await Zip.read(original);

		// Check that the rebuilt zip file has the same keys as the original zip file
		const origKeys = Object.keys(originalZip).sort();
		const rebuiltKeys = Object.keys(rebuiltZip).sort();

		expect(rebuiltKeys).toEqual(origKeys);

		// find new rows in the rebuilt zip file
		const sheet1Rebuilt = rebuiltZip["xl/worksheets/sheet1.xml"].toString();

		const rebuiltXml = Xml.extractRowsFromSheet(sheet1Rebuilt);

		expect(rebuiltXml.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>John</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>30</t></is></c><c r=\"C3\" t=\"inlineStr\"><is><t>New York</t></is></c></row>",
			"<row r=\"4\"><c r=\"A4\" t=\"inlineStr\"><is><t>Jane</t></is></c><c r=\"B4\" t=\"inlineStr\"><is><t>25</t></is></c><c r=\"C4\" t=\"inlineStr\"><is><t>Los Angeles</t></is></c></row>",
		]);

		const validationResult = validateWorksheetXml(sheet1Rebuilt);

		expect(validationResult.isValid).toBe(true);
	});

	it("should insertRows with saveStream", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		await template.insertRows({
			rows: [
				["Name", "Age", "City"],
				["John", 30, "New York"],
				["Jane", 25, "Los Angeles"],
			],
			sheetName: "Sheet1",
			startRowNumber: 2,
		});

		// Save the buffer to a file
		await template.saveStream(fsSync.createWriteStream(OUTPUT_FILE));

		const buffer = await fs.readFile(OUTPUT_FILE);
		expect(buffer).toBeDefined();

		// Read the rebuilt file
		const original = await fs.readFile(INPUT_FILE);
		const rebuilt = await fs.readFile(OUTPUT_FILE);

		// Read the rebuilt zip file
		const rebuiltZip = await Zip.read(rebuilt);
		const originalZip = await Zip.read(original);

		// Check that the rebuilt zip file has the same keys as the original zip file
		const origKeys = Object.keys(originalZip).sort();
		const rebuiltKeys = Object.keys(rebuiltZip).sort();

		expect(rebuiltKeys).toEqual(origKeys);

		// find new rows in the rebuilt zip file
		const sheet1Rebuilt = rebuiltZip["xl/worksheets/sheet1.xml"].toString();

		const rebuiltXml = Xml.extractRowsFromSheet(sheet1Rebuilt);

		expect(rebuiltXml.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>John</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>30</t></is></c><c r=\"C3\" t=\"inlineStr\"><is><t>New York</t></is></c></row>",
			"<row r=\"4\"><c r=\"A4\" t=\"inlineStr\"><is><t>Jane</t></is></c><c r=\"B4\" t=\"inlineStr\"><is><t>25</t></is></c><c r=\"C4\" t=\"inlineStr\"><is><t>Los Angeles</t></is></c></row>",
		]);

		const validationResult = validateWorksheetXml(sheet1Rebuilt);

		expect(validationResult.isValid).toBe(true);
	});
});
