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
const INPUT_FILE = path.resolve(ASSETS_DIR, "input-test-03.xlsx");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-03.xlsx");

async function* asyncRowsGenerator(count: number): AsyncIterable<unknown[]> {
	for (let i = 0; i < count; i++) {
		await new Promise((resolve) => setTimeout(resolve, 0));
		yield Array(10).fill(["Name", "Age", "City"]);
	}
}

describe("TemplateFs integration test", () => {
	it("should insertRowsStream with save", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		await template.insertRowsStream({
			rows: asyncRowsGenerator(2),
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

		const sheet1RowsData = Xml.extractRowsFromSheet(sheet1Rebuilt);

		expect(sheet1RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C3\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"4\"><c r=\"A4\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B4\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C4\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B5\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C5\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"6\"><c r=\"A6\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B6\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C6\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"7\"><c r=\"A7\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B7\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C7\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"8\"><c r=\"A8\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B8\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C8\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"9\"><c r=\"A9\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B9\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C9\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"10\"><c r=\"A10\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B10\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C10\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"11\"><c r=\"A11\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B11\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C11\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"12\"><c r=\"A12\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B12\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C12\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"13\"><c r=\"A13\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B13\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C13\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"14\"><c r=\"A14\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B14\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C14\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"15\"><c r=\"A15\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B15\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C15\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"16\"><c r=\"A16\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B16\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C16\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"17\"><c r=\"A17\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B17\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C17\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"18\"><c r=\"A18\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B18\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C18\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"19\"><c r=\"A19\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B19\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C19\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"20\"><c r=\"A20\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B20\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C20\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"21\"><c r=\"A21\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B21\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C21\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
		]);

		const validationResult = validateWorksheetXml(sheet1Rebuilt);

		expect(validationResult.isValid).toBe(true);
	});

	it("should insertRowsStream with saveStream", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		await template.insertRowsStream({
			rows: asyncRowsGenerator(2),
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

		const sheet1RowsData = Xml.extractRowsFromSheet(sheet1Rebuilt);

		expect(sheet1RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C3\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"4\"><c r=\"A4\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B4\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C4\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B5\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C5\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"6\"><c r=\"A6\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B6\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C6\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"7\"><c r=\"A7\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B7\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C7\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"8\"><c r=\"A8\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B8\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C8\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"9\"><c r=\"A9\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B9\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C9\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"10\"><c r=\"A10\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B10\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C10\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"11\"><c r=\"A11\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B11\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C11\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"12\"><c r=\"A12\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B12\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C12\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"13\"><c r=\"A13\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B13\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C13\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"14\"><c r=\"A14\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B14\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C14\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"15\"><c r=\"A15\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B15\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C15\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"16\"><c r=\"A16\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B16\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C16\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"17\"><c r=\"A17\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B17\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C17\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"18\"><c r=\"A18\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B18\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C18\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"19\"><c r=\"A19\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B19\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C19\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"20\"><c r=\"A20\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B20\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C20\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
			"<row r=\"21\"><c r=\"A21\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B21\" t=\"inlineStr\"><is><t>Age</t></is></c><c r=\"C21\" t=\"inlineStr\"><is><t>City</t></is></c></row>",
		]);

		const validationResult = validateWorksheetXml(sheet1Rebuilt);

		expect(validationResult.isValid).toBe(true);
	});
});
