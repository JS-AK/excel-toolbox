import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

export function buildWorkbookXml(sheets: { name: string }[]): string {
	return [
		XML_DECLARATION,
		buildXml({
			attrs: {
				xmlns: XML_NAMESPACES.SPREADSHEET_ML,
				"xmlns:r": XML_NAMESPACES.OFFICE_DOCUMENT,
			},
			children: [
				{
					children: sheets.map((sheet, i) => ({
						attrs: { name: sheet.name, "r:id": `rId${i + 1}`, sheetId: (i + 1).toString() },
						tag: "sheet",
					})),
					tag: "sheets",
				},
			],
			tag: "workbook",
		}),
	].join("\n");
}
