import { XmlNode } from "../types/index.js";

import { CONTENT_TYPES, XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

export function buildContentTypesXml(sheetsCount: number): string {
	const defaults: XmlNode[] = [
		{ ContentType: CONTENT_TYPES.RELATIONSHIPS, Extension: "rels" },
		{ ContentType: CONTENT_TYPES.XML, Extension: "xml" },
	].map(({ ContentType, Extension }) => ({
		attrs: { ContentType, Extension },
		tag: "Default",
	}));

	const overrides: XmlNode[] = [
		{ ContentType: CONTENT_TYPES.THEME, PartName: "/xl/theme/theme1.xml" },
		{ ContentType: CONTENT_TYPES.WORKBOOK, PartName: "/xl/workbook.xml" },
		{ ContentType: CONTENT_TYPES.STYLES, PartName: "/xl/styles.xml" },
		{ ContentType: CONTENT_TYPES.CORE, PartName: "/docProps/core.xml" },
		{ ContentType: CONTENT_TYPES.APP, PartName: "/docProps/app.xml" },
		{ ContentType: CONTENT_TYPES.SHARED_STRINGS, PartName: "/xl/sharedStrings.xml" },
	].map(({ ContentType, PartName }) => ({
		attrs: { ContentType, PartName },
		tag: "Override",
	}));

	for (let i = 1; i <= sheetsCount; i++) {
		overrides.push({
			attrs: {
				ContentType: CONTENT_TYPES.WORKSHEET,
				PartName: `/xl/worksheets/sheet${i}.xml`,
			},
			tag: "Override",
		});
	}

	return [
		XML_DECLARATION,
		buildXml({
			attrs: { xmlns: XML_NAMESPACES.CONTENT_TYPES },
			children: [...defaults, ...overrides],
			tag: "Types",
		}),
	].join("\n");
}
