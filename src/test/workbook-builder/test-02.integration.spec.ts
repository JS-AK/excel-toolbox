import fs from "node:fs/promises";
import path from "node:path";

import { beforeEach, describe, expect, it } from "vitest";

import * as Zip from "../../lib/zip/index.js";
import { WorkbookBuilder } from "../../lib/workbook-builder/workbook-builder.js";

const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "workbook-builder", "assets");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-02.xlsx");

describe("WorkbookBuilder integration test", () => {
	beforeEach(async () => {
		if (await fs.stat(OUTPUT_FILE).then(() => true).catch(() => false)) {
			await fs.unlink(OUTPUT_FILE);
		}
	});

	it("should create workbook with multiple sheets", async () => {
		const wb = new WorkbookBuilder();

		// Add multiple sheets
		wb.addSheet("Sales");
		wb.addSheet("Products");
		wb.addSheet("Customers");

		// Add data to Sales sheet
		const salesSheet = wb.getSheet("Sales");

		if (!salesSheet) {
			throw new Error("Sheet 'Sales' not found");
		}

		salesSheet.setCell(1, "A", { type: "s", value: "Date" });
		// shared-string-ref 0 -> Date

		salesSheet.setCell(1, "B", { type: "s", value: "Product" });
		// shared-string-ref 1 -> Product

		salesSheet.setCell(1, "C", { type: "s", value: "Quantity" });
		// shared-string-ref 2 -> Quantity

		salesSheet.setCell(1, "D", { type: "s", value: "Price" });
		// shared-string-ref 3 -> Price

		salesSheet.setCell(2, "A", { type: "s", value: "2024-01-01" });
		// shared-string-ref 4 -> 2024-01-01

		salesSheet.setCell(2, "B", { type: "s", value: "Product A" });
		// shared-string-ref 5 -> Product A

		salesSheet.setCell(2, "C", { type: "n", value: 10 });
		salesSheet.setCell(2, "D", { type: "n", value: 25.5 });

		// Add data to Products sheet
		const productsSheet = wb.getSheet("Products");

		if (!productsSheet) {
			throw new Error("Sheet 'Products' not found");
		}

		productsSheet.setCell(1, "A", { type: "s", value: "Product ID" });
		// shared-string-ref 6 -> Product ID

		productsSheet.setCell(1, "B", { type: "s", value: "Name" });
		// shared-string-ref 7 -> Name

		productsSheet.setCell(1, "C", { type: "s", value: "Category" });
		// shared-string-ref 8 -> Category

		productsSheet.setCell(2, "A", { type: "s", value: "A001" });
		// shared-string-ref 9 -> A001

		productsSheet.setCell(2, "B", { type: "s", value: "Product A" });
		// shared-string-ref Product A exists -> 5

		productsSheet.setCell(2, "C", { type: "s", value: "Electronics" });
		// shared-string-ref 10 -> Electronics

		// Add data to Customers sheet
		const customersSheet = wb.getSheet("Customers");

		if (!customersSheet) {
			throw new Error("Sheet 'Customers' not found");
		}

		customersSheet.setCell(1, "A", { type: "s", value: "Customer ID" });
		// shared-string-ref 11 -> Customer ID

		customersSheet.setCell(1, "B", { type: "s", value: "Name" });
		// shared-string-ref Product A exists -> 7

		customersSheet.setCell(1, "C", { type: "s", value: "Email" });
		// shared-string-ref 12 -> Email

		customersSheet.setCell(2, "A", { type: "s", value: "C001" });
		// shared-string-ref 13 -> C001

		customersSheet.setCell(2, "B", { type: "s", value: "John Doe" });
		// shared-string-ref 14 -> John Doe

		customersSheet.setCell(2, "C", { type: "s", value: "john@example.com" });
		// shared-string-ref 15 -> john@example.com

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
		expect(zipContent["xl/worksheets/sheet1.xml"]).toBeDefined(); // Sheet1
		expect(zipContent["xl/worksheets/sheet2.xml"]).toBeDefined(); // Sales
		expect(zipContent["xl/worksheets/sheet3.xml"]).toBeDefined(); // Products
		expect(zipContent["xl/worksheets/sheet4.xml"]).toBeDefined(); // Customers

		// Check workbook.xml contains all sheets
		const workbookXml = zipContent["xl/workbook.xml"].toString();

		expect(workbookXml).toContain("Sheet1");
		expect(workbookXml).toContain("Sales");
		expect(workbookXml).toContain("Products");
		expect(workbookXml).toContain("Customers");

		// Check that each sheet contains its data

		const salesXml = zipContent["xl/worksheets/sheet2.xml"].toString();

		expect(salesXml).toContain("<v>0</v>"); // type="s"
		expect(salesXml).toContain("<v>1</v>"); // type="s"
		expect(salesXml).toContain("<v>2</v>"); // type="s"
		expect(salesXml).toContain("<v>3</v>"); // type="s"
		expect(salesXml).toContain("<v>4</v>"); // type="s"
		expect(salesXml).toContain("<v>5</v>"); // type="s"
		expect(salesXml).toContain("<v>10</v>"); // type="n"
		expect(salesXml).toContain("<v>25.5</v>"); // type="n"

		const productsXml = zipContent["xl/worksheets/sheet3.xml"].toString();

		expect(productsXml).toContain("<v>6</v>"); // type="s"
		expect(productsXml).toContain("<v>7</v>"); // type="s"
		expect(productsXml).toContain("<v>8</v>"); // type="s"
		expect(productsXml).toContain("<v>9</v>"); // type="s"
		expect(productsXml).toContain("<v>5</v>"); // type="s"
		expect(productsXml).toContain("<v>10</v>"); // type="s"

		const customersXml = zipContent["xl/worksheets/sheet4.xml"].toString();

		expect(customersXml).toContain("<v>11</v>"); // type="s"
		expect(customersXml).toContain("<v>7</v>"); // type="s"
		expect(customersXml).toContain("<v>12</v>"); // type="s"
		expect(customersXml).toContain("<v>13</v>"); // type="s"
		expect(customersXml).toContain("<v>14</v>"); // type="s"
		expect(customersXml).toContain("<v>15</v>"); // type="s"

		// Check that the shared strings file contains our data
		const sharedStringsXml = zipContent["xl/sharedStrings.xml"].toString();

		expect(sharedStringsXml).toContain("<t>Date</t>");
		expect(sharedStringsXml).toContain("<t>Product</t>");
		expect(sharedStringsXml).toContain("<t>Quantity</t>");
		expect(sharedStringsXml).toContain("<t>Price</t>");
		expect(sharedStringsXml).toContain("<t>2024-01-01</t>");
		expect(sharedStringsXml).toContain("<t>Product A</t>");
		expect(sharedStringsXml).toContain("<t>Product ID</t>");
		expect(sharedStringsXml).toContain("<t>Name</t>");
		expect(sharedStringsXml).toContain("<t>Category</t>");
		expect(sharedStringsXml).toContain("<t>A001</t>");
		expect(sharedStringsXml).toContain("<t>Electronics</t>");
		expect(sharedStringsXml).toContain("<t>Customer ID</t>");
		expect(sharedStringsXml).toContain("<t>Email</t>");
		expect(sharedStringsXml).toContain("<t>C001</t>");
		expect(sharedStringsXml).toContain("<t>John Doe</t>");
		expect(sharedStringsXml).toContain("<t>john@example.com</t>");

		// Clean up
		await fs.unlink(OUTPUT_FILE);
	});
});
