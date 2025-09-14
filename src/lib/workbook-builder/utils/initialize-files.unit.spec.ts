import { describe, expect, it } from "vitest";

import {
	CONTENT_TYPES,
	FILE_PATHS,
	RELATIONSHIP_TYPES,
	XML_DECLARATION,
	XML_NAMESPACES,
} from "./constants.js";

import { trimAndJoinMultiline } from "../../utils/trim-and-join-multiline.js";

import { initializeFiles } from "./initialize-files.js";

import { sheetName } from "../default/index.js";

describe("initializeFiles", () => {
	it("should return an object with required keys", () => {
		const files = initializeFiles(sheetName());

		const files2 = {
			"[Content_Types].xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
					<Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>
					<Default ContentType="application/xml" Extension="xml"/>
					<Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml" PartName="/xl/theme/theme1.xml"/>
					<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
					<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>
					<Override ContentType="application/vnd.openxmlformats-package.core-properties+xml" PartName="/docProps/core.xml"/>
					<Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" PartName="/docProps/app.xml"/>
					<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
					<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet1.xml"/>
				</Types>`,
			"_rels/.rels": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
					<Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
					<Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
					<Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>
				</Relationships>`,
			"docProps/app.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
					<Application>Microsoft Excel</Application>
					<DocSecurity>0</DocSecurity>
					<ScaleCrop>false</ScaleCrop>
					<HeadingPairs>
						<vt:vector baseType="variant" size="2">
							<vt:variant>
								<vt:lpstr>Worksheets</vt:lpstr>
							</vt:variant>
							<vt:variant>
								<vt:i4>1</vt:i4>
							</vt:variant>
						</vt:vector>
					</HeadingPairs>
					<TitlesOfParts>
						<vt:vector baseType="lpstr" size="1">
							<vt:lpstr>Sheet1</vt:lpstr>
						</vt:vector>
					</TitlesOfParts>
					<Company></Company>
					<LinksUpToDate>false</LinksUpToDate>
					<SharedDoc>false</SharedDoc>
					<HyperlinksChanged>false</HyperlinksChanged>
					<AppVersion>16.0300</AppVersion>
				</Properties>`,
			"docProps/core.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
					<dc:creator>Excel Generator</dc:creator>
					<cp:lastModifiedBy>Excel Generator</cp:lastModifiedBy>
					<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
					<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>
				</cp:coreProperties>`,
			"xl/_rels/workbook.xml.rels": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
					<Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
					<Relationship Id="rId2" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
					<Relationship Id="rId3" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>
					<Relationship Id="rId4" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
				</Relationships>`,
			"xl/sharedStrings.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<sst count="0" uniqueCount="0" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`,
			"xl/styles.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
					<fonts count="1">
						<font>
							<sz val="11"/>
							<color theme="1"/>
							<name val="Calibri"/>
							<family val="2"/>
							<scheme val="minor"/>
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
			"xl/theme/theme1.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<a:theme name="Office Theme" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
					<a:themeElements>
						<a:clrScheme name="Office">
							<a:dk1>
								<a:sysClr lastClr="000000" val="windowText"/>
							</a:dk1>
							<a:lt1>
								<a:sysClr lastClr="FFFFFF" val="window"/>
							</a:lt1>
							<a:dk2>
								<a:srgbClr val="1F497D"/>
							</a:dk2>
							<a:lt2>
								<a:srgbClr val="EEECE1"/>
							</a:lt2>
							<a:accent1>
								<a:srgbClr val="4F81BD"/>
							</a:accent1>
							<a:accent2>
								<a:srgbClr val="C0504D"/>
							</a:accent2>
							<a:accent3>
								<a:srgbClr val="9BBB59"/>
							</a:accent3>
							<a:accent4>
								<a:srgbClr val="8064A2"/>
							</a:accent4>
							<a:accent5>
								<a:srgbClr val="4BACC6"/>
							</a:accent5>
							<a:accent6>
								<a:srgbClr val="F79646"/>
							</a:accent6>
							<a:hlink>
								<a:srgbClr val="0000FF"/>
							</a:hlink>
							<a:folHlink>
								<a:srgbClr val="800080"/>
							</a:folHlink>
						</a:clrScheme>
						<a:fontScheme name="Office">
							<a:majorFont>
								<a:latin typeface="Calibri Light"/>
								<a:ea typeface=""/>
								<a:cs typeface=""/>
								<a:font script="Jpan" typeface="游ゴシック Light"/>
								<a:font script="Hang" typeface="맑은 고딕"/>
								<a:font script="Hans" typeface="等线 Light"/>
								<a:font script="Hant" typeface="新細明體"/>
								<a:font script="Arab" typeface="Times New Roman"/>
								<a:font script="Hebr" typeface="Times New Roman"/>
								<a:font script="Thai" typeface="Tahoma"/>
								<a:font script="Ethi" typeface="Nyala"/>
								<a:font script="Beng" typeface="Vrinda"/>
								<a:font script="Gujr" typeface="Shruti"/>
								<a:font script="Khmr" typeface="MoolBoran"/>
								<a:font script="Knda" typeface="Tunga"/>
								<a:font script="Guru" typeface="Raavi"/>
								<a:font script="Cans" typeface="Euphemia"/>
								<a:font script="Cher" typeface="Plantagenet Cherokee"/>
								<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
								<a:font script="Tibt" typeface="Microsoft Himalaya"/>
								<a:font script="Thaa" typeface="MV Boli"/>
								<a:font script="Deva" typeface="Mangal"/>
								<a:font script="Telu" typeface="Gautami"/>
								<a:font script="Taml" typeface="Latha"/>
								<a:font script="Syrc" typeface="Estrangelo Edessa"/>
								<a:font script="Orya" typeface="Kalinga"/>
								<a:font script="Mlym" typeface="Kartika"/>
								<a:font script="Laoo" typeface="DokChampa"/>
								<a:font script="Sinh" typeface="Iskoola Pota"/>
								<a:font script="Mong" typeface="Mongolian Baiti"/>
								<a:font script="Viet" typeface="Times New Roman"/>
								<a:font script="Uigh" typeface="Microsoft Uighur"/>
								<a:font script="Geor" typeface="Sylfaen"/>
							</a:majorFont>
							<a:minorFont>
								<a:latin typeface="Calibri"/>
								<a:ea typeface=""/>
								<a:cs typeface=""/>
								<a:font script="Jpan" typeface="游ゴシック"/>
								<a:font script="Hang" typeface="맑은 고딕"/>
								<a:font script="Hans" typeface="等线"/>
								<a:font script="Hant" typeface="新細明體"/>
								<a:font script="Arab" typeface="Arial"/>
								<a:font script="Hebr" typeface="Arial"/>
								<a:font script="Thai" typeface="Tahoma"/>
								<a:font script="Ethi" typeface="Nyala"/>
								<a:font script="Beng" typeface="Vrinda"/>
								<a:font script="Gujr" typeface="Shruti"/>
								<a:font script="Khmr" typeface="DaunPenh"/>
								<a:font script="Knda" typeface="Tunga"/>
								<a:font script="Guru" typeface="Raavi"/>
								<a:font script="Cans" typeface="Euphemia"/>
								<a:font script="Cher" typeface="Plantagenet Cherokee"/>
								<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
								<a:font script="Tibt" typeface="Microsoft Himalaya"/>
								<a:font script="Thaa" typeface="MV Boli"/>
								<a:font script="Deva" typeface="Mangal"/>
								<a:font script="Telu" typeface="Gautami"/>
								<a:font script="Taml" typeface="Latha"/>
								<a:font script="Syrc" typeface="Estrangelo Edessa"/>
								<a:font script="Orya" typeface="Kalinga"/>
								<a:font script="Mlym" typeface="Kartika"/>
								<a:font script="Laoo" typeface="DokChampa"/>
								<a:font script="Sinh" typeface="Iskoola Pota"/>
								<a:font script="Mong" typeface="Mongolian Baiti"/>
								<a:font script="Viet" typeface="Arial"/>
								<a:font script="Uigh" typeface="Microsoft Uighur"/>
								<a:font script="Geor" typeface="Sylfaen"/>
							</a:minorFont>
						</a:fontScheme>
						<a:fmtScheme name="Office">
							<a:fillStyleLst>
								<a:solidFill>
									<a:schemeClr val="phClr"/>
								</a:solidFill>
								<a:gradFill rotWithShape="1">
									<a:gsLst>
										<a:gs pos="0">
											<a:schemeClr val="phClr">
												<a:tint val="50000"/>
												<a:satMod val="300000"/>
											</a:schemeClr>
										</a:gs>
										<a:gs pos="35000">
											<a:schemeClr val="phClr">
												<a:tint val="37000"/>
												<a:satMod val="300000"/>
											</a:schemeClr>
										</a:gs>
										<a:gs pos="100000">
											<a:schemeClr val="phClr">
												<a:tint val="15000"/>
												<a:satMod val="350000"/>
											</a:schemeClr>
										</a:gs>
									</a:gsLst>
									<a:lin ang="16200000" scaled="1"/>
								</a:gradFill>
								<a:gradFill rotWithShape="1">
									<a:gsLst>
										<a:gs pos="0">
											<a:schemeClr val="phClr">
												<a:shade val="51000"/>
												<a:satMod val="130000"/>
											</a:schemeClr>
										</a:gs>
										<a:gs pos="80000">
											<a:schemeClr val="phClr">
												<a:shade val="93000"/>
												<a:satMod val="130000"/>
											</a:schemeClr>
										</a:gs>
										<a:gs pos="100000">
											<a:schemeClr val="phClr">
												<a:shade val="94000"/>
												<a:satMod val="135000"/>
											</a:schemeClr>
										</a:gs>
									</a:gsLst>
									<a:lin ang="16200000" scaled="0"/>
								</a:gradFill>
							</a:fillStyleLst>
							<a:lnStyleLst>
								<a:ln algn="ctr" cap="flat" cmpd="sng" w="9525">
									<a:solidFill>
										<a:schemeClr val="phClr">
											<a:shade val="95000"/>
											<a:satMod val="105000"/>
										</a:schemeClr>
									</a:solidFill>
									<a:prstDash val="solid"/>
								</a:ln>
								<a:ln algn="ctr" cap="flat" cmpd="sng"  w="25400">
									<a:solidFill>
										<a:schemeClr val="phClr"/>
									</a:solidFill>
									<a:prstDash val="solid"/>
								</a:ln>
								<a:ln algn="ctr" cap="flat" cmpd="sng" w="38100">
									<a:solidFill>
										<a:schemeClr val="phClr"/>
									</a:solidFill>
									<a:prstDash val="solid"/>
								</a:ln>
							</a:lnStyleLst>
							<a:effectStyleLst>
								<a:effectStyle>
									<a:effectLst>
										<a:outerShdw blurRad="40000" dir="5400000" dist="20000" rotWithShape="0">
											<a:srgbClr val="000000">
												<a:alpha val="38000"/>
											</a:srgbClr>
										</a:outerShdw>
									</a:effectLst>
								</a:effectStyle>
								<a:effectStyle>
									<a:effectLst>
										<a:outerShdw blurRad="40000" dir="5400000" dist="23000" rotWithShape="0">
											<a:srgbClr val="000000">
												<a:alpha val="35000"/>
											</a:srgbClr>
										</a:outerShdw>
									</a:effectLst>
								</a:effectStyle>
								<a:effectStyle>
									<a:effectLst>
										<a:outerShdw blurRad="40000" dir="5400000" dist="23000" rotWithShape="0">
											<a:srgbClr val="000000">
												<a:alpha val="35000"/>
											</a:srgbClr>
										</a:outerShdw>
									</a:effectLst>
									<a:scene3d>
										<a:camera prst="orthographicFront">
											<a:rot lat="0" lon="0" rev="0"/>
										</a:camera>
										<a:lightRig dir="t" rig="threePt">
											<a:rot lat="0" lon="0" rev="0"/>
										</a:lightRig>
									</a:scene3d>
									<a:sp3d>
										<a:bevelT h="25400" w="63500"/>
									</a:sp3d>
								</a:effectStyle>
							</a:effectStyleLst>
							<a:bgFillStyleLst>
								<a:solidFill>
									<a:schemeClr val="phClr"/>
								</a:solidFill>
								<a:gradFill rotWithShape="1">
									<a:gsLst>
										<a:gs pos="0">
											<a:schemeClr val="phClr">
												<a:tint val="40000"/>
												<a:satMod val="350000"/>
											</a:schemeClr>
										</a:gs>
										<a:gs pos="40000">
										<a:schemeClr val="phClr">
										<a:tint val="45000"/>
										<a:shade val="99000"/>
										<a:satMod val="350000"/>
										</a:schemeClr>
										</a:gs>
										<a:gs pos="100000">
										<a:schemeClr val="phClr">
										<a:shade val="20000"/>
										<a:satMod val="255000"/>
										</a:schemeClr>
										</a:gs>
									</a:gsLst>
									<a:path path="circle">
										<a:fillToRect b="180000" l="50000" r="50000" t="-80000"/>
									</a:path>
								</a:gradFill>
								<a:gradFill rotWithShape="1">
									<a:gsLst>
										<a:gs pos="0">
											<a:schemeClr val="phClr">
												<a:tint val="80000"/>
												<a:satMod val="300000"/>
											</a:schemeClr>
										</a:gs>
										<a:gs pos="100000">
											<a:schemeClr val="phClr">
												<a:shade val="30000"/>
												<a:satMod val="200000"/>
											</a:schemeClr>
										</a:gs>
									</a:gsLst>
									<a:path path="circle">
										<a:fillToRect b="50000" l="50000" r="50000" t="50000"/>
									</a:path>
								</a:gradFill>
							</a:bgFillStyleLst>
						</a:fmtScheme>
					</a:themeElements>
					<a:objectDefaults/>
					<a:extraClrSchemeLst/>
				</a:theme>`,
			"xl/workbook.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
					<sheets>
						<sheet name="Sheet1" r:id="rId1" sheetId="1"/>
					</sheets>
				</workbook>`,
			"xl/worksheets/sheet1.xml": `
				<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
			const files2String = trimAndJoinMultiline({ inputString: files2[key], separator: "" });
			const filesString = trimAndJoinMultiline({ inputString: files[key] as string, separator: "" });

			expect(files2String).toEqual(filesString);
		}

		for (const key of Object.keys(files)) {
			const files2String = trimAndJoinMultiline({ inputString: files2[key], separator: "" });
			const filesString = trimAndJoinMultiline({ inputString: files[key] as string, separator: "" });

			expect(files2String).toEqual(filesString);
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

	it("each XML should start with a declaration", () => {
		const files = initializeFiles(sheetName());

		for (const [, content] of Object.entries(files)) {
			const xml = content.toString();
			expect(xml.startsWith(XML_DECLARATION)).toBe(true);
		}
	});

	it("[Content_Types].xml contains correct Override and Default entries", () => {
		const content = initializeFiles(sheetName())[FILE_PATHS.CONTENT_TYPES].toString();

		expect(content).toContain(`<Default ContentType="${CONTENT_TYPES.RELATIONSHIPS}" Extension="rels"/>`);
		expect(content).toContain(`<Default ContentType="${CONTENT_TYPES.XML}" Extension="xml"/>`);
		expect(content).toContain(`<Override ContentType="${CONTENT_TYPES.WORKBOOK}" PartName="/xl/workbook.xml"/>`);
		expect(content).toContain(`<Override ContentType="${CONTENT_TYPES.STYLES}" PartName="/xl/styles.xml"/>`);
		expect(content).toContain(`<Override ContentType="${CONTENT_TYPES.SHARED_STRINGS}" PartName="/xl/sharedStrings.xml"/>`);
	});

	it("_rels/.rels contains a relationship to workbook.xml", () => {
		const content = initializeFiles(sheetName())[FILE_PATHS.RELS].toString();

		expect(content).toContain(`<Relationship Id="rId1" Target="xl/workbook.xml" Type="${RELATIONSHIP_TYPES.OFFICE_DOCUMENT}"/>`);
	});

	it("workbook.xml contains <sheets> tag", () => {
		const content = initializeFiles(sheetName())[FILE_PATHS.WORKBOOK].toString();

		expect(content).toContain("<sheets>");
		expect(content).toContain("<sheet name=\"Sheet1\" r:id=\"rId1\" sheetId=\"1\"/>");
		expect(content).toContain(`xmlns="${XML_NAMESPACES.SPREADSHEET_ML}"`);
		expect(content).toContain(`xmlns:r="${XML_NAMESPACES.OFFICE_DOCUMENT}"`);
	});

	it("styles.xml contains a base styleSheet with required xmlns", () => {
		const content = initializeFiles(sheetName())[FILE_PATHS.STYLES].toString();

		expect(content).toContain(`<styleSheet xmlns="${XML_NAMESPACES.SPREADSHEET_ML}">`);
	});

	it("sharedStrings.xml contains sst with count=0 and uniqueCount=0", () => {
		const content = initializeFiles(sheetName())[FILE_PATHS.SHARED_STRINGS].toString();

		expect(content).toContain(`<sst count="0" uniqueCount="0" xmlns="${XML_NAMESPACES.SPREADSHEET_ML}"/>`);
	});

	it("xl/_rels/workbook.xml.rels initially contains an empty Relationships list", () => {
		const content = initializeFiles(sheetName())[FILE_PATHS.WORKBOOK_RELS].toString();

		expect(content).toContain(`<Relationships xmlns="${XML_NAMESPACES.PACKAGE_RELATIONSHIPS}">`);
	});
});
