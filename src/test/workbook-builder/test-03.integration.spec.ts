import fs from "node:fs/promises";
import path from "node:path";

import { describe, expect, it } from "vitest";

import * as Zip from "../../lib/zip/index.js";
import { WorkbookBuilder } from "../../lib/workbook-builder/workbook-builder.js";

const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "workbook-builder", "assets");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-03.xlsx");

describe("WorkbookBuilder integration test", () => {
	it("should create workbook with merge cells", async () => {
		const wb = new WorkbookBuilder();

		// Add data to the sheet
		const sheet = wb.getSheet("Sheet1");

		if (!sheet) {
			throw new Error("Sheet 'Sheet1' not found");
		}

		if (sheet) {
			sheet.setCell(1, "A", { type: "s", value: "Report Title" });
			// shared-string-ref 0 -> Report Title

			sheet.setCell(1, "B", { type: "s", value: "" });
			// shared-string-ref 1 -> ""

			sheet.setCell(1, "C", { type: "s", value: "" });
			// shared-string-ref 1 -> ""

			sheet.setCell(1, "D", { type: "s", value: "" });
			// shared-string-ref 1 -> ""

			sheet.setCell(2, "A", { type: "s", value: "Q1" });
			// shared-string-ref 2 -> Q1

			sheet.setCell(2, "B", { type: "s", value: "Q2" });
			// shared-string-ref 3 -> Q2

			sheet.setCell(2, "C", { type: "s", value: "Q3" });
			// shared-string-ref 4 -> Q3

			sheet.setCell(2, "D", { type: "s", value: "Q4" });
			// shared-string-ref 5 -> Q4

			sheet.setCell(3, "A", { type: "s", value: "Sales" });
			// shared-string-ref 6 -> Sales

			sheet.setCell(3, "B", { type: "n", value: 1000 });
			sheet.setCell(3, "C", { type: "n", value: 1200 });
			sheet.setCell(3, "D", { type: "n", value: 1100 });

			sheet.setCell(4, "A", { type: "s", value: "Expenses" });
			// shared-string-ref 7 -> Expenses

			sheet.setCell(4, "B", { type: "n", value: 800 });
			sheet.setCell(4, "C", { type: "n", value: 900 });
			sheet.setCell(4, "D", { type: "n", value: 850 });
		}

		sheet.addMerge({ endCol: 3, endRow: 1, startCol: 0, startRow: 1 }); // Report Title
		sheet.addMerge({ endCol: 0, endRow: 4, startCol: 0, startRow: 3 }); // Sales/Expenses label

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

		// Check that the worksheet contains merge cells
		const worksheetXml = zipContent["xl/worksheets/sheet1.xml"].toString();
		expect(worksheetXml).toContain("<mergeCells");
		expect(worksheetXml).toContain("<mergeCell ref=\"A1:D1\"/>");
		expect(worksheetXml).toContain("<mergeCell ref=\"A3:A4\"/>");
		expect(worksheetXml).toContain("</mergeCells>");

		// Check that the data is present
		expect(worksheetXml).toContain("<v>0</v>"); // type="s"
		expect(worksheetXml).toContain("<v>1</v>"); // type="s"
		expect(worksheetXml).toContain("<v>2</v>"); // type="s"
		expect(worksheetXml).toContain("<v>3</v>"); // type="s"
		expect(worksheetXml).toContain("<v>4</v>"); // type="s"
		expect(worksheetXml).toContain("<v>5</v>"); // type="s"
		expect(worksheetXml).toContain("<v>6</v>"); // type="s"
		expect(worksheetXml).toContain("1000");
		expect(worksheetXml).toContain("1100");
		expect(worksheetXml).toContain("1200");
		expect(worksheetXml).toContain("<v>7</v>"); // type="s"

		// Check that the shared strings file contains our data
		const sharedStringsXml = zipContent["xl/sharedStrings.xml"].toString();

		expect(sharedStringsXml).toContain("<t>Report Title</t>");
		expect(sharedStringsXml).toContain("<t></t>");
		expect(sharedStringsXml).toContain("<t>Q1</t>");
		expect(sharedStringsXml).toContain("<t>Q2</t>");
		expect(sharedStringsXml).toContain("<t>Q3</t>");
		expect(sharedStringsXml).toContain("<t>Q4</t>");
		expect(sharedStringsXml).toContain("<t>Sales</t>");
		expect(sharedStringsXml).toContain("<t>Expenses</t>");

		// Clean up
		await fs.unlink(OUTPUT_FILE);
	});
});
