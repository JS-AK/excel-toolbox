import { escapeXml } from "../../utils/index.js";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

/**
 * Builds the `sharedStrings.xml` content from the provided strings.
 *
 * Note: According to the Excel specification, leading/trailing spaces in text
 * nodes must set the attribute `xml:space="preserve"` on the corresponding <t>.
 *
 * @param strings - Array of unique strings used in the workbook
 * @returns XML string for the shared strings part
 */
export function buildSharedStringsXml(strings: string[] = []): string {
	return [
		XML_DECLARATION,
		buildXml({
			attrs: {
				count: String(strings.length),
				uniqueCount: String(strings.length),
				xmlns: XML_NAMESPACES.SPREADSHEET_ML,
			},
			children: strings.map((s) => ({
				children: [
					{
						// Excel spec: leading/trailing spaces require xml:space="preserve"
						attrs: /^\s|\s$/.test(s) ? { "xml:space": "preserve" } : undefined,
						children: [escapeXml(s)],
						tag: "t",
					},
				],
				tag: "si",
			})),
			tag: "sst",
		}),
	].join("\n");
}
