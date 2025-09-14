import fs from "node:fs";
import fsPromises from "node:fs/promises";
import path from "node:path";

import * as Default from "../default/index.js";
import { CellXf, XmlNode } from "../types/index.js";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { writeXml } from "./write-xml.js";

export async function writeStylesXml(
	destination: string,
	data?: {
		borders: XmlNode["children"];
		cellXfs: CellXf[];
		fills: XmlNode["children"];
		fonts: XmlNode["children"];
		numFmts: { formatCode: string; id: number }[];
	},
): Promise<void> {
	// create with folder
	await fsPromises.mkdir(path.dirname(destination), { recursive: true });

	const stream = fs.createWriteStream(destination, { encoding: "utf-8" });

	try {

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
				children: [Default.font()],
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
				children: [Default.fill()],
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
				children: [Default.border()],
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

		// XML declaration
		stream.write(XML_DECLARATION + "\n");

		await writeXml(
			{
				attrs: { xmlns: XML_NAMESPACES.SPREADSHEET_ML },
				children,
				tag: "styleSheet",
			},
			stream,
		);

	} finally {
		stream.end();
	}

	return new Promise((resolve, reject) => {
		stream.on("error", reject);
		stream.on("finish", resolve);
	});
}
