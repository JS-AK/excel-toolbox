import fs from "node:fs/promises";
import path from "node:path";

import { describe, expect, it } from "vitest";

import * as Xml from "../../lib/xml/index.js";
import * as Zip from "../../lib/zip/index.js";
import { TemplateMemory } from "../../lib/template/template-memory.js";
import { validateWorksheetXml } from "../../lib/template/utils/validate-worksheet-xml.js";

const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "template-memory", "assets");
const INPUT_FILE = path.resolve(ASSETS_DIR, "input-test-04.xlsx");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-04.xlsx");

describe("TemplateMemory integration test", () => {
	it("should insert rows into a sheet with save", async () => {
		const template = await TemplateMemory.from({
			source: INPUT_FILE,
		});

		await template.substitute("Sheet1", {
			name: "John",
			name1: "John 1",
			name2: "John 2",
			name3: "John 3",
			name4: "John 4",
			name5: "John 5",
			user: {
				name: {
					name1: "Jack 1",
					name2: "Jack 2",
					name3: "Jack 3",
					name4: "Jack 4",
					name5: "Jack 5",
				},
				name1: "Jane 1",
				name2: "Jane 2",
				name3: "Jane 3",
				name4: "Jane 4",
				name5: "Jane 5",
			},
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

		const sheet1RowsData = await Xml.extractRowsFromSheet(sheet1Rebuilt);

		expect(sheet1RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\" spans=\"1:5\"><c r=\"A2\" t=\"s\"><v>0</v></c><c r=\"B2\" t=\"s\"><v>0</v></c><c r=\"C2\" t=\"s\"><v>0</v></c><c r=\"D2\" t=\"s\"><v>0</v></c><c r=\"E2\" t=\"s\"><v>0</v></c></row>",
			"<row r=\"3\" spans=\"1:5\"><c r=\"A3\" t=\"s\"><v>2</v></c><c r=\"B3\" t=\"s\"><v>12</v></c><c r=\"C3\" t=\"s\"><v>13</v></c><c r=\"D3\" t=\"s\"><v>14</v></c><c r=\"E3\" t=\"s\"><v>15</v></c></row>",
			"<row r=\"4\" spans=\"1:5\"><c r=\"A4\" t=\"s\"><v>3</v></c><c r=\"B4\" t=\"s\"><v>4</v></c><c r=\"C4\" t=\"s\"><v>5</v></c><c r=\"D4\" t=\"s\"><v>1</v></c><c r=\"E4\" t=\"s\"><v>6</v></c></row>",
			"<row r=\"5\" spans=\"1:5\"><c r=\"A5\" t=\"s\"><v>7</v></c><c r=\"B5\" t=\"s\"><v>8</v></c><c r=\"C5\" t=\"s\"><v>9</v></c><c r=\"D5\" t=\"s\"><v>11</v></c><c r=\"E5\" t=\"s\"><v>10</v></c></row>",
		]);

		const validationResult = validateWorksheetXml(sheet1Rebuilt);

		expect(validationResult.isValid).toBe(true);
	});
});
