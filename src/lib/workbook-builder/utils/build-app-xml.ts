import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { XmlNode, buildXml } from "./build-xml.js";

export interface AppXmlOptions {
	appVersion?: string;
	application?: string;
	company?: string;
	hyperlinksChanged?: boolean;
	linksUpToDate?: boolean;
	sharedDoc?: boolean;
	sheetNames?: string[];
}

export function buildAppXml({
	appVersion = "16.0300",
	application = "Microsoft Excel",
	company = "",
	hyperlinksChanged = false,
	linksUpToDate = false,
	sharedDoc = false,
	sheetNames = ["Sheet1"],
}: AppXmlOptions = {}): string {
	const app: XmlNode = {
		attrs: {
			xmlns: XML_NAMESPACES.EXTENDED_PROPERTIES,
			"xmlns:vt": XML_NAMESPACES.DOC_PROPS_VTYPES,
		},
		children: [
			{ children: [application], tag: "Application" },
			{ children: ["0"], tag: "DocSecurity" },
			{ children: ["false"], tag: "ScaleCrop" },
			{
				children: [
					{
						attrs: { baseType: "variant", size: "2" },
						children: [
							{ children: [{ children: ["Worksheets"], tag: "vt:lpstr" }], tag: "vt:variant" },
							{ children: [{ children: [String(sheetNames.length)], tag: "vt:i4" }], tag: "vt:variant" },
						],
						tag: "vt:vector",
					},
				],
				tag: "HeadingPairs",
			},
			{
				children: [
					{
						attrs: { baseType: "lpstr", size: String(sheetNames.length) },
						children: sheetNames.map((name) => ({ children: [name], tag: "vt:lpstr" })),
						tag: "vt:vector",
					},
				],
				tag: "TitlesOfParts",
			},
			{ children: [company], tag: "Company" },
			{ children: [String(linksUpToDate)], tag: "LinksUpToDate" },
			{ children: [String(sharedDoc)], tag: "SharedDoc" },
			{ children: [String(hyperlinksChanged)], tag: "HyperlinksChanged" },
			{ children: [appVersion], tag: "AppVersion" },
		],
		tag: "Properties",
	};

	return [
		XML_DECLARATION,
		buildXml(app),
	].join("\n");
}
