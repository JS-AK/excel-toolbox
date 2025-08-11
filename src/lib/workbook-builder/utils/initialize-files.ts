import { RowData } from "./sheet.js";
import { buildContentTypesXml } from "./build-content-types-xml.js";
import { buildSharedStringsXml } from "./build-shared-strings-xml.js";
import { buildWorkbookRels } from "./build-workbook-rels-xml.js";
import { buildWorkbookXml } from "./build-workbook-xml.js";
import { buildWorksheetXml } from "./build-worksheet-xml.js";
import { buildXml } from "./build-xml.js";

import {
	FILE_PATHS,
	RELATIONSHIP_TYPES,
	XML_DECLARATION,
	XML_NAMESPACES,
} from "./constants.js";

export type ExcelFileContent = Buffer | string;

export type ExcelFiles = {
	[FILE_PATHS.CONTENT_TYPES]: ExcelFileContent;
	[FILE_PATHS.RELS]: ExcelFileContent;
	[FILE_PATHS.WORKBOOK_RELS]: ExcelFileContent;
	[FILE_PATHS.STYLES]: ExcelFileContent;
	[FILE_PATHS.SHARED_STRINGS]: ExcelFileContent;
	[FILE_PATHS.WORKBOOK]: ExcelFileContent;
	[FILE_PATHS.WORKSHEET]: ExcelFileContent;
	[key: string]: ExcelFileContent;
}

export const initializeFiles = (): ExcelFiles => {
	const declaration = XML_DECLARATION;
	const sheetsCount = 1;

	const contentTypesXml = buildContentTypesXml(sheetsCount);

	const relsXml = [
		declaration,
		buildXml({
			attrs: { xmlns: XML_NAMESPACES.PACKAGE_RELATIONSHIPS },
			children: [{ attrs: { Id: "rId1", Target: "xl/workbook.xml", Type: RELATIONSHIP_TYPES.OFFICE_DOCUMENT }, tag: "Relationship" }],
			tag: "Relationships",
		}),
	].join("\n");

	const workbookRelsXml = buildWorkbookRels(sheetsCount);

	const stylesXml = [
		declaration,
		buildXml({
			attrs: { xmlns: XML_NAMESPACES.SPREADSHEET_ML },
			children: [
				{
					attrs: { count: "1" },
					children: [
						{
							children: [
								{ attrs: { val: "11" }, tag: "sz" },
								{ attrs: { theme: "1" }, tag: "color" },
								{ attrs: { val: "Calibri" }, tag: "name" },
							],
							tag: "font",
						},
					],
					tag: "fonts",
				},
				{
					attrs: { count: "1" },
					children: [
						{
							children: [
								{ attrs: { patternType: "none" }, tag: "patternFill" },
							],
							tag: "fill",
						},
					],
					tag: "fills",
				},
				{
					attrs: { count: "1" },
					children: [
						{
							children: [
								{ tag: "left" },
								{ tag: "right" },
								{ tag: "top" },
								{ tag: "bottom" },
							],
							tag: "border",
						},
					],
					tag: "borders",
				},
				{
					attrs: { count: "1" },
					children: [
						{ attrs: { borderId: "0", fillId: "0", fontId: "0", numFmtId: "0" }, tag: "xf" },
					],
					tag: "cellStyleXfs",
				},
				{
					attrs: { count: "1" },
					children: [
						{ attrs: { borderId: "0", fillId: "0", fontId: "0", numFmtId: "0", xfId: "0" }, tag: "xf" },
					],
					tag: "cellXfs",
				},
			],
			tag: "styleSheet",
		}),
	].join("\n");

	const sharedStringsXml = buildSharedStringsXml([]);

	const workbookXml = buildWorkbookXml([{ name: "Sheet1" }]);

	const worksheet = buildWorksheetXml(new Map<number, RowData>());

	return {
		[FILE_PATHS.CONTENT_TYPES]: contentTypesXml,
		[FILE_PATHS.RELS]: relsXml,
		[FILE_PATHS.WORKBOOK_RELS]: workbookRelsXml,
		[FILE_PATHS.STYLES]: stylesXml,
		[FILE_PATHS.SHARED_STRINGS]: sharedStringsXml,
		[FILE_PATHS.WORKBOOK]: workbookXml,
		[FILE_PATHS.WORKSHEET]: worksheet,
	};
};
