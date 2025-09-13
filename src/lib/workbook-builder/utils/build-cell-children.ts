import type { CellData } from "./sheet.js";
import type { XmlNode } from "./build-xml.js";

/**
 * Builds XML children nodes for a cell based on its type and value.
 *
 * @param cell - The cell data containing value and type information
 *
 * @returns Array of XML nodes representing the cell's content structure
 */
export function buildCellChildren(cell: CellData): XmlNode[] {
	if (cell.value === undefined) {
		return [];
	}

	if (cell.isFormula) {
		return [
			{
				children: [String(cell.value)],
				tag: "f",
			},
		];
	}

	switch (cell.type) {
		case "b": {
			return [
				{
					children: [cell.value ? "1" : "0"],
					tag: "v",
				},
			];
		}

		case "inlineStr": {
			// For inlineStr, wrap value in <is><t>value</t></is>
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
			return [
				{
					children: [String(cell.value)],
					tag: "v",
				},
			];
		}
	}
}
