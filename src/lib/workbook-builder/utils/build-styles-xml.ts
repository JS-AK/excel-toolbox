import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { XmlNode, buildXml } from "./build-xml.js";

type Alignment = {
	horizontal?: "left" | "center" | "right" | "justify";
	vertical?: "top" | "center" | "bottom";
	wrapText?: boolean;
	indent?: number;
};

type CellXf = {
	numFmtId: number;
	borderId: number;
	fillId: number;
	fontId: number;
	alignment?: Alignment;
};

export type CellXfs = CellXf[];

export function buildStylesXml(data?: {
	borders: XmlNode["children"];
	cellXfs: CellXfs;
	fills: XmlNode["children"];
	fonts: XmlNode["children"];
	numFmts: { formatCode: string; id: number }[];
}): string {
	const {
		borders = [],
		cellXfs = [],
		fills = [],
		fonts = [],
		numFmts = [],
	} = data || {};

	const children: XmlNode["children"] = [];

	if (numFmts.length) {
		children.push({
			attrs: { count: String(numFmts.length) },
			children: numFmts.map(nf => ({
				attrs: {
					formatCode: nf.formatCode,
					numFmtId: String(nf.id),
				},
				tag: "numFmt",
			})),
			tag: "numFmts",
		});
	}

	if (fonts.length) {
		children.push({
			attrs: { count: String(fonts.length) },
			children: fonts,
			tag: "fonts",
		});
	} else {
		children.push({
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
		});
	}

	if (fills.length) {
		children.push({
			attrs: { count: String(fills.length) },
			children: fills,
			tag: "fills",
		});
	} else {
		children.push({
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
		});
	}

	if (borders.length) {
		children.push({
			attrs: { count: String(borders.length) },
			children: borders,
			tag: "borders",
		});
	} else {
		children.push({
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
		});
	}

	children.push({
		attrs: { count: "1" },
		children: [{
			attrs: { borderId: "0", fillId: "0", fontId: "0", numFmtId: "0" },
			tag: "xf",
		}],
		tag: "cellStyleXfs",
	});

	if (cellXfs.length) {
		children.push({
			attrs: { count: String(cellXfs.length) },
			children: cellXfs.map((xf, i) => {
				const isBaseXf = i === 0;
				const hasAlignment = !!xf.alignment;

				const xfChildren: XmlNode["children"] = [];

				if (hasAlignment) {
					xfChildren.push({
						attrs: {
							...(xf.alignment?.horizontal ? { horizontal: xf.alignment.horizontal } : {}),
							...(xf.alignment?.vertical ? { vertical: xf.alignment.vertical } : {}),
							...(xf.alignment?.wrapText ? { wrapText: "1" } : {}),
							...(xf.alignment?.indent !== undefined ? { indent: String(xf.alignment.indent) } : {}),
						},
						tag: "alignment",
					});
				}

				return {
					attrs: {
						...(isBaseXf
							? {}
							: {
								applyBorder: "1",
								applyFill: "1",
								applyFont: "1",
								applyNumberFormat: xf.numFmtId ? "1" : "0",
							}),
						...(hasAlignment ? { applyAlignment: "1" } : {}),
						borderId: String(xf.borderId),
						fillId: String(xf.fillId),
						fontId: String(xf.fontId),
						numFmtId: String(xf.numFmtId ?? 0),
						xfId: "0",
					},
					children: xfChildren,
					tag: "xf",
				};
			}),
			tag: "cellXfs",
		});
	} else {
		//<!--базовый стиль без заливки-- >
		children.push({
			attrs: { count: "1" },
			children: [
				{ attrs: { borderId: "0", fillId: "0", fontId: "0", numFmtId: "0", xfId: "0" }, tag: "xf" },
			],
			tag: "cellXfs",
		});
	}

	return [
		XML_DECLARATION,
		buildXml({
			attrs: { xmlns: XML_NAMESPACES.SPREADSHEET_ML },
			children,
			tag: "styleSheet",
		}),
	].join("\n");
}
