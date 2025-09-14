import { buildAppXml } from "./build-app-xml.js";
import { buildContentTypesXml } from "./build-content-types-xml.js";
import { buildCoreXml } from "./build-core-xml.js";
import { buildRelsXml } from "./build-rels-xml.js";
import { buildSharedStringsXml } from "./build-shared-strings-xml.js";
import { buildStylesXml } from "./build-styles-xml.js";
import { buildThemeXml } from "./build-theme-xml.js";
import { buildWorkbookRels } from "./build-workbook-rels-xml.js";
import { buildWorkbookXml } from "./build-workbook-xml.js";
import { buildWorksheetXml } from "./build-worksheet-xml.js";

import { FILE_PATHS } from "./constants.js";

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
	const sheetsCount = 1;

	const contentTypesXml = buildContentTypesXml(sheetsCount);
	const relsXml = buildRelsXml();
	const appXml = buildAppXml({ sheetNames: [sheetName] });
	const coreXml = buildCoreXml();
	const workbookRelsXml = buildWorkbookRels(sheetsCount);
	const stylesXml = buildStylesXml();
	const sharedStringsXml = buildSharedStringsXml();
	const themeXml = buildThemeXml();
	const workbookXml = buildWorkbookXml([{ name: sheetName }]);
	const worksheetXml = buildWorksheetXml();

	return {
		[FILE_PATHS.CONTENT_TYPES]: contentTypesXml,
		[FILE_PATHS.RELS]: relsXml,
		[FILE_PATHS.APP]: appXml,
		[FILE_PATHS.CORE]: coreXml,
		[FILE_PATHS.WORKBOOK_RELS]: workbookRelsXml,
		[FILE_PATHS.STYLES]: stylesXml,
		[FILE_PATHS.SHARED_STRINGS]: sharedStringsXml,
		[FILE_PATHS.THEME]: themeXml,
		[FILE_PATHS.WORKBOOK]: workbookXml,
		[FILE_PATHS.WORKSHEET]: worksheetXml,
	};
};
