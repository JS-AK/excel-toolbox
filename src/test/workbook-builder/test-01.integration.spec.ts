import fs from "node:fs/promises";
import path from "node:path";

import { beforeEach, describe, expect, it } from "vitest";

import * as Zip from "../../lib/zip/index.js";
import { WorkbookBuilder } from "../../lib/workbook-builder/workbook-builder.js";

const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "workbook-builder", "assets");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-01.xlsx");

describe("WorkbookBuilder integration test", () => {
	beforeEach(async () => {
		if (await fs.stat(OUTPUT_FILE).then(() => true).catch(() => false)) {
			await fs.unlink(OUTPUT_FILE);
		}
	});

	it("should create basic workbook with single sheet", async () => {
		const wb = new WorkbookBuilder();

		// Add data to the default sheet
		const sheet = wb.getSheet("Sheet1");

		if (!sheet) {
			throw new Error("Sheet 'Sheet1' not found");
		}

		sheet.setCell(1, "A", { type: "s", value: "Name" });
		// shared-string-ref 0 -> Name

		sheet.setCell(1, "B", { type: "s", value: "Age" });
		// shared-string-ref 1 -> Age

		sheet.setCell(1, "C", { type: "s", value: "City" });
		// shared-string-ref 2 -> City

		sheet.setCell(2, "A", { type: "s", value: "John" });
		// shared-string-ref 3 -> John

		sheet.setCell(2, "B", { type: "n", value: 30 });

		sheet.setCell(2, "C", { type: "s", value: "New York" });
		// shared-string-ref 4 -> New York

		sheet.setCell(3, "A", { type: "s", value: "Jane" });
		// shared-string-ref 5 -> Jane

		sheet.setCell(3, "B", { type: "n", value: 25 });

		sheet.setCell(3, "C", { type: "s", value: "Los Angeles" });
		// shared-string-ref 6 -> Los Angeles

		// Save the workbook
		await wb.saveToFile(OUTPUT_FILE);

		// Check that the file was created
		const stats = await fs.stat(OUTPUT_FILE);
		expect(stats.isFile()).toBe(true);
		expect(stats.size).toBeGreaterThan(0);

		// Read the generated file
		const buffer = await fs.readFile(OUTPUT_FILE);

		// Read the zip file
		const zipContent = await Zip.read(buffer);

		// Check that all required files are present
		expect(zipContent["[Content_Types].xml"]).toBeDefined();
		expect(zipContent["_rels/.rels"]).toBeDefined();
		expect(zipContent["docProps/app.xml"]).toBeDefined();
		expect(zipContent["docProps/core.xml"]).toBeDefined();
		expect(zipContent["xl/workbook.xml"]).toBeDefined();
		expect(zipContent["xl/_rels/workbook.xml.rels"]).toBeDefined();
		expect(zipContent["xl/worksheets/sheet1.xml"]).toBeDefined();
		expect(zipContent["xl/styles.xml"]).toBeDefined();
		expect(zipContent["xl/sharedStrings.xml"]).toBeDefined();
		expect(zipContent["xl/theme/theme1.xml"]).toBeDefined();

		// Check that the worksheet contains our data
		const worksheetXml = zipContent["xl/worksheets/sheet1.xml"].toString();

		expect(worksheetXml).toContain("<v>0</v>"); // type="s"
		expect(worksheetXml).toContain("<v>1</v>"); // type="s"
		expect(worksheetXml).toContain("<v>2</v>"); // type="s"
		expect(worksheetXml).toContain("<v>3</v>"); // type="s"
		expect(worksheetXml).toContain("<v>30</v>"); // type="n"
		expect(worksheetXml).toContain("<v>4</v>"); // type="s"
		expect(worksheetXml).toContain("<v>5</v>"); // type="s"
		expect(worksheetXml).toContain("<v>25</v>"); // type="n"
		expect(worksheetXml).toContain("<v>6</v>"); // type="s"

		// Check that the shared strings file contains our data
		const sharedStringsXml = zipContent["xl/sharedStrings.xml"].toString();

		expect(sharedStringsXml).toContain("<t>Name</t>");
		expect(sharedStringsXml).toContain("<t>Age</t>");
		expect(sharedStringsXml).toContain("<t>City</t>");
		expect(sharedStringsXml).toContain("<t>John</t>");
		expect(sharedStringsXml).toContain("<t>Jane</t>");
		expect(sharedStringsXml).toContain("<t>New York</t>");
		expect(sharedStringsXml).toContain("<t>Los Angeles</t>");

		// Clean up
		await fs.unlink(OUTPUT_FILE);
	});
});
