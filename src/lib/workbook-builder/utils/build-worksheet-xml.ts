import { CellData, RowData, SheetData } from "./sheet.js";
import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { XmlNode, buildXml } from "./build-xml.js";

export function buildWorksheetXml(
	rows: SheetData["rows"] = new Map<number, RowData>(),
	merges: string[] = [],
): string {
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

					return {
						attrs: { r: cellRef, s: cell.style?.index, t: cell.type },
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

function buildCellChildren(cell: CellData) {
	if (cell.value === undefined) return [];

	switch (cell.type) {
		case "b": {
			return [{
				children: [cell.value ? "1" : "0"],
				tag: "v",
			}];
		}

		case "inlineStr": {
			// для inlineStr вложение <is><t>значение</t></is>
			return [
				{
					children: [
						{
							children: [String(cell.value)],
							tag: "t",
						},
					],
					tag: "is",
				},
			];
		}

		case "s":
		case "n":
		case "str":
		case "e":
		default: {
			return [{
				children: [String(cell.value)],
				tag: "v",
			}];
		}
	}
}
