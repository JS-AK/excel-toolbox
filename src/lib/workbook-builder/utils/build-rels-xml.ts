import { RELATIONSHIP_TYPES, XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

export function buildRelsXml(): string {
	return [
		XML_DECLARATION,
		buildXml({
			attrs: { xmlns: XML_NAMESPACES.PACKAGE_RELATIONSHIPS },
			children: [
				{ attrs: { Id: "rId1", Target: "xl/workbook.xml", Type: RELATIONSHIP_TYPES.OFFICE_DOCUMENT }, tag: "Relationship" },
				{ attrs: { Id: "rId2", Target: "docProps/core.xml", Type: RELATIONSHIP_TYPES.CORE }, tag: "Relationship" },
				{ attrs: { Id: "rId3", Target: "docProps/app.xml", Type: RELATIONSHIP_TYPES.APP }, tag: "Relationship" },
			],
			tag: "Relationships",
		}),
	].join("\n");
}
