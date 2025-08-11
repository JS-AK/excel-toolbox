import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

/**
 * Создаёт содержимое файла `sharedStrings.xml` на основе переданных строк.
 *
 * @param strings - Массив уникальных строк, используемых в книге
 */
export function buildSharedStringsXml(strings: string[]): string {
	const escapeXml = (str: string) =>
		str
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			.replace(/'/g, "&apos;");

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
						// По спецификации Excel: пробелы в начале/конце требуют xml:space="preserve"
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
