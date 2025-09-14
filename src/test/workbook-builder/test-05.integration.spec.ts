import fs from "node:fs/promises";
import path from "node:path";

import { beforeEach, describe, expect, it } from "vitest";

import * as Zip from "../../lib/zip/index.js";
import { WorkbookBuilder } from "../../lib/workbook-builder/workbook-builder.js";

const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "workbook-builder", "assets");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-05.xlsx");

describe("WorkbookBuilder integration test", () => {
	beforeEach(async () => {
		if (await fs.stat(OUTPUT_FILE).then(() => true).catch(() => false)) {
			await fs.unlink(OUTPUT_FILE);
		}
	});

	it("should create workbook with different cell types", async () => {
		const wb = new WorkbookBuilder();

		// Add data with different cell types
		const sheet = wb.getSheet("Sheet1");

		if (!sheet) {
			throw new Error("Sheet 'Sheet1' not found");
		}

		// Headers
		sheet.setCell(1, "A", { type: "s", value: "String" });
		sheet.setCell(1, "B", { type: "s", value: "Number" });
		sheet.setCell(1, "C", { type: "s", value: "Boolean" });
		sheet.setCell(1, "D", { type: "s", value: "Date" });
		sheet.setCell(1, "E", { type: "s", value: "Inline String" });
		sheet.setCell(1, "F", { type: "s", value: "Formula" });
		sheet.setCell(1, "G", { type: "s", value: "Error" });

		// Data rows with different types
		sheet.setCell(2, "A", { type: "s", value: "Hello World" });
		sheet.setCell(2, "B", { type: "n", value: 123.45 });
		sheet.setCell(2, "C", { type: "b", value: true });
		sheet.setCell(2, "D", { style: { numberFormat: "yyyy-mm-dd" }, value: new Date(2024, 0, 15) });
		sheet.setCell(2, "E", { type: "inlineStr", value: "Inline Text" });
		sheet.setCell(2, "F", { isFormula: true, value: "=SUM(A2:A10)" });
		sheet.setCell(2, "G", { type: "e", value: "#DIV/0!" });

		// Add more rows with different types
		sheet.setCell(3, "A", { type: "s", value: "Another String" });
		sheet.setCell(3, "B", { type: "n", value: -67.89 });
		sheet.setCell(3, "C", { type: "b", value: false });
		sheet.setCell(3, "D", { value: new Date(2024, 11, 31).toISOString() });
		sheet.setCell(3, "E", { type: "inlineStr", value: "More Inline" });
		sheet.setCell(3, "F", { isFormula: true, value: "=AVERAGE(B2:B10)" });
		sheet.setCell(3, "G", { type: "e", value: "#N/A" });

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

		// Check that shared strings file is present (for string type)
		expect(zipContent["xl/sharedStrings.xml"]).toBeDefined();

		// Check shared strings content
		const sharedStringsXml = zipContent["xl/sharedStrings.xml"].toString();

		expect(sharedStringsXml).toContain("Hello World");
		expect(sharedStringsXml).toContain("Another String");

		// Check worksheet content
		const worksheetXml = zipContent["xl/worksheets/sheet1.xml"].toString();

		expect(worksheetXml).toContain(`
    <row r="1">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1" t="s">
        <v>1</v>
      </c>
      <c r="C1" t="s">
        <v>2</v>
      </c>
      <c r="D1" t="s">
        <v>3</v>
      </c>
      <c r="E1" t="s">
        <v>4</v>
      </c>
      <c r="F1" t="s">
        <v>5</v>
      </c>
      <c r="G1" t="s">
        <v>6</v>
      </c>
    </row>
    <row r="2">
      <c r="A2" t="s">
        <v>7</v>
      </c>
      <c r="B2">
        <v>123.45</v>
      </c>
      <c r="C2" t="b">
        <v>1</v>
      </c>
      <c r="D2" s="1">
        <v>45306</v>
      </c>
      <c r="E2" t="inlineStr">
        <is>
          <t>Inline Text</t>
        </is>
      </c>
      <c r="F2">
        <f>=SUM(A2:A10)</f>
      </c>
      <c r="G2" t="e">
        <v>#DIV/0!</v>
      </c>
    </row>
    <row r="3">
      <c r="A3" t="s">
        <v>8</v>
      </c>
      <c r="B3">
        <v>-67.89</v>
      </c>
      <c r="C3" t="b">
        <v>0</v>
      </c>
      <c r="D3" t="inlineStr">
        <is>
          <t>2024-12-31T00:00:00.000Z</t>
        </is>
      </c>
      <c r="E3" t="inlineStr">
        <is>
          <t>More Inline</t>
        </is>
      </c>
      <c r="F3">
        <f>=AVERAGE(B2:B10)</f>
      </c>
      <c r="G3" t="e">
        <v>#N/A</v>
      </c>
    </row>`);

		// Clean up
		await fs.unlink(OUTPUT_FILE);
	});
});
