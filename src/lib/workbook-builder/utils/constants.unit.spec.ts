import { describe, expect, it } from "vitest";

import {
	CONTENT_TYPES,
	FILE_PATHS,
	RELATIONSHIP_TYPES,
	XML_DECLARATION,
	XML_NAMESPACES,
} from "./constants.js"; // путь замени на актуальный

import { initializeFiles } from "./initialize-files.js";

describe("initializeFiles", () => {
	it("должна возвращать объект с обязательными ключами", () => {
		const files = initializeFiles();

		const files2 = {
			"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>
  <Default ContentType="application/xml" Extension="xml"/>
  <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
  <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>
  <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
  <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet1.xml"/>
</Types>`,
			"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
</Relationships>`,
			"xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
  <Relationship Id="rId2" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
  <Relationship Id="rId3" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
</Relationships>`,
			"xl/sharedStrings.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst count="0" uniqueCount="0" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`,
			"xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
  </cellXfs>
</styleSheet>`,
			"xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" r:id="rId1" sheetId="1"/>
  </sheets>
</workbook>`,
			"xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:A1"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
</worksheet>`,
		};

		for (const key of Object.keys(files2)) {
			expect(files2[key]).toEqual(files[key]);
		}

		for (const key of Object.keys(files)) {
			expect(files2[key]).toEqual(files[key]);
		}

		expect(Object.keys(files)).toEqual(
			expect.arrayContaining([
				FILE_PATHS.CONTENT_TYPES,
				FILE_PATHS.RELS,
				FILE_PATHS.WORKBOOK_RELS,
				FILE_PATHS.STYLES,
				FILE_PATHS.SHARED_STRINGS,
				FILE_PATHS.WORKBOOK,
			]),
		);
	});

	it("каждый XML должен начинаться с декларации", () => {
		const files = initializeFiles();

		for (const [, content] of Object.entries(files)) {
			const xml = content.toString();
			expect(xml.startsWith(XML_DECLARATION)).toBe(true);
		}
	});

	it("[Content_Types].xml содержит правильные Override и Default", () => {
		const content = initializeFiles()[FILE_PATHS.CONTENT_TYPES].toString();

		expect(content).toContain(`<Default ContentType="${CONTENT_TYPES.RELATIONSHIPS}" Extension="rels"/>`);
		expect(content).toContain(`<Default ContentType="${CONTENT_TYPES.XML}" Extension="xml"/>`);
		expect(content).toContain(`<Override ContentType="${CONTENT_TYPES.WORKBOOK}" PartName="/xl/workbook.xml"/>`);
		expect(content).toContain(`<Override ContentType="${CONTENT_TYPES.STYLES}" PartName="/xl/styles.xml"/>`);
		expect(content).toContain(`<Override ContentType="${CONTENT_TYPES.SHARED_STRINGS}" PartName="/xl/sharedStrings.xml"/>`);
	});

	it("_rels/.rels содержит ссылку на workbook.xml", () => {
		const content = initializeFiles()[FILE_PATHS.RELS].toString();

		expect(content).toContain(`<Relationship Id="rId1" Target="xl/workbook.xml" Type="${RELATIONSHIP_TYPES.OFFICE_DOCUMENT}"/>`);
	});

	it("workbook.xml содержит тег <sheets>", () => {
		const content = initializeFiles()[FILE_PATHS.WORKBOOK].toString();

		expect(content).toContain("<sheets>");
		expect(content).toContain("<sheet name=\"Sheet1\" r:id=\"rId1\" sheetId=\"1\"/>");
		expect(content).toContain(`xmlns="${XML_NAMESPACES.SPREADSHEET_ML}"`);
		expect(content).toContain(`xmlns:r="${XML_NAMESPACES.OFFICE_DOCUMENT}"`);
	});

	it("styles.xml содержит пустой styleSheet с нужным xmlns", () => {
		const content = initializeFiles()[FILE_PATHS.STYLES].toString();

		expect(content).toContain(`<styleSheet xmlns="${XML_NAMESPACES.SPREADSHEET_ML}">`);
	});

	it("sharedStrings.xml содержит sst с count=0 и uniqueCount=0", () => {
		const content = initializeFiles()[FILE_PATHS.SHARED_STRINGS].toString();

		expect(content).toContain(`<sst count="0" uniqueCount="0" xmlns="${XML_NAMESPACES.SPREADSHEET_ML}"/>`);
	});

	it("xl/_rels/workbook.xml.rels изначально пустой список Relationships", () => {
		const content = initializeFiles()[FILE_PATHS.WORKBOOK_RELS].toString();

		expect(content).toContain(`<Relationships xmlns="${XML_NAMESPACES.PACKAGE_RELATIONSHIPS}">`);
	});
});
