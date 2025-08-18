import { buildContentTypesXml } from "./build-content-types-xml.js";
import { buildSharedStringsXml } from "./build-shared-strings-xml.js";
import { buildStylesXml } from "./build-styles-xml.js";
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

export const initializeFiles = (sheetName: string): ExcelFiles => {
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
	const stylesXml = buildStylesXml();
	const sharedStringsXml = buildSharedStringsXml();
	const workbookXml = buildWorkbookXml([{ name: sheetName }]);
	const worksheet = buildWorksheetXml();

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
