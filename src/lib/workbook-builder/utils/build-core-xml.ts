import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { XmlNode, buildXml } from "./build-xml.js";

export function buildCoreXml(
	creator = "Excel Generator",
	lastModifiedBy = "Excel Generator",
	created = "2025-01-01T00:00:00Z",
	modified = "2025-01-01T00:00:00Z",
): string {
	const coreProps: XmlNode = {
		attrs: {
			"xmlns:cp": XML_NAMESPACES.CORE_PROPERTIES,
			"xmlns:dc": XML_NAMESPACES.DC,
			"xmlns:dcmitype": XML_NAMESPACES.DCMITYPE,
			"xmlns:dcterms": XML_NAMESPACES.DCTERMS,
			"xmlns:xsi": XML_NAMESPACES.XSI,
		},
		children: [
			{ children: [creator], tag: "dc:creator" },
			{ children: [lastModifiedBy], tag: "cp:lastModifiedBy" },
			{ attrs: { "xsi:type": "dcterms:W3CDTF" }, children: [created], tag: "dcterms:created" },
			{ attrs: { "xsi:type": "dcterms:W3CDTF" }, children: [modified], tag: "dcterms:modified" },
		],
		tag: "cp:coreProperties",
	};

	return [
		XML_DECLARATION,
		buildXml(coreProps),
	].join("\n");
}
