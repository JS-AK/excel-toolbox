import { RowData, SheetData, XmlNode } from "../types/index.js";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildCellChildren } from "./build-cell-children.js";
import { buildXml } from "./build-xml.js";

export function buildWorksheetXml(
	rows: SheetData["rows"] = new Map<number, RowData>(),
	merges: string[] = [],
): string {
	const typeSetSkipping = new Set<string>(["n"]);

	const children: XmlNode["children"] = [
		{
			attrs: { ref: "A1:A1" },
			tag: "dimension",
		},
		{
			children: [{ attrs: { workbookViewId: "0" }, tag: "sheetView" }],
			tag: "sheetViews",
		},
		{
			attrs: { defaultRowHeight: "15" },
			tag: "sheetFormatPr",
		},
		{
			children: Array.from(rows.entries()).map(([rowNumber, row]) => ({
				attrs: { r: rowNumber.toString() },
				children: Array.from(row.cells.entries()).map(([colNumber, cell]) => {
					const cellRef = `${colNumber}${rowNumber}`;

					const attrT = (cell.type && typeSetSkipping.has(cell.type))
						? undefined
						: cell.type;

					return {
						attrs: {
							r: cellRef,
							s: cell.style?.index,
							t: attrT,
						},
						children: buildCellChildren(cell),
						tag: "c",
					};
				}),
				tag: "row",
			})),
			tag: "sheetData",
		},
	];

	if (merges.length > 0) {
		children.push({
			attrs: { count: merges.length.toString() },
			children: merges.map((ref) => ({
				attrs: { ref },
				tag: "mergeCell",
			})),
			tag: "mergeCells",
		});
	}

	return [
		XML_DECLARATION,
		buildXml({
			attrs: {
				xmlns: XML_NAMESPACES.SPREADSHEET_ML,
				"xmlns:r": XML_NAMESPACES.OFFICE_DOCUMENT,
			},
			children,
			tag: "worksheet",
		}),
	].join("\n");
}
