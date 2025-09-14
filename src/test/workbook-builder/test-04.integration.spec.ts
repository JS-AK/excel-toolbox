import fs from "node:fs/promises";
import path from "node:path";

import { beforeEach, describe, expect, it } from "vitest";

import * as Zip from "../../lib/zip/index.js";
import { WorkbookBuilder } from "../../lib/workbook-builder/workbook-builder.js";

const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "workbook-builder", "assets");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-04.xlsx");

describe("WorkbookBuilder integration test", () => {
	beforeEach(async () => {
		if (await fs.stat(OUTPUT_FILE).then(() => true).catch(() => false)) {
			await fs.unlink(OUTPUT_FILE);
		}
	});

	it("should create workbook with shared strings", async () => {
		const wb = new WorkbookBuilder();

		// Add data to the sheet using shared strings
		const sheet = wb.getSheet("Sheet1");
		if (sheet) {
			// Headers
			sheet.setCell(1, "A", { type: "s", value: "Product Name" });
			sheet.setCell(1, "B", { type: "s", value: "Category" });
			sheet.setCell(1, "C", { type: "s", value: "Price" });

			// Data rows
			sheet.setCell(2, "A", { type: "s", value: "Laptop" });
			sheet.setCell(2, "B", { type: "s", value: "Electronics" });
			sheet.setCell(2, "C", { type: "n", value: 999.99 });

			sheet.setCell(3, "A", { type: "s", value: "T-Shirt" });
			sheet.setCell(3, "B", { type: "s", value: "Clothing" });
			sheet.setCell(3, "C", { type: "n", value: 19.99 });

			sheet.setCell(4, "A", { type: "s", value: "Programming Guide" });
			sheet.setCell(4, "B", { type: "s", value: "Books" });
			sheet.setCell(4, "C", { type: "n", value: 49.99 });
		}

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

		// Check that shared strings file is present
		expect(zipContent["xl/sharedStrings.xml"]).toBeDefined();

		// Check shared strings content
		const sharedStringsXml = zipContent["xl/sharedStrings.xml"].toString();
		expect(sharedStringsXml).toContain("<sst");
		expect(sharedStringsXml).toContain("Product Name");
		expect(sharedStringsXml).toContain("Electronics");
		expect(sharedStringsXml).toContain("Laptop");
		expect(sharedStringsXml).toContain("T-Shirt");
		expect(sharedStringsXml).toContain("Programming Guide");

		// Check that the worksheet references shared strings
		const worksheetXml = zipContent["xl/worksheets/sheet1.xml"].toString();
		expect(worksheetXml).toContain("t=\"s\""); // Shared string type
		expect(worksheetXml).toContain("<v>0</v>"); // Reference to "Product Name"
		expect(worksheetXml).toContain("<v>1</v>"); // Reference to "Category"

		// Clean up
		await fs.unlink(OUTPUT_FILE);
	});
});
