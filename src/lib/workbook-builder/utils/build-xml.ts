/**
 * Represents an XML node structure for building XML documents.
 */
export type XmlNode = {
	/** The XML tag name */
	tag: string;
	/** Optional attributes for the XML element */
	attrs?: Record<string, string | number | undefined | null>;
	/** Child elements or text content */
	children?: (string | XmlNode)[];
};

/**
 * Builds XML string from an XmlNode structure.
 *
 * @param node - The XML node to convert to string
 * @param level - The indentation level for formatting (default: 0)
 * 
 * @returns The formatted XML string
 */
export function buildXml(node: XmlNode, level = 0): string {
	const { attrs = {}, children = [], tag } = node;

	const attrStr = Object.entries(attrs)
		.filter(attr => ((attr[1] !== undefined) && (attr[1] !== null)))
		.map(([k, v]) => ` ${k}="${v}"`)
		.join("");

	const gap = "  ".repeat(level);

	// No children → self-closing tag
	if (!children.length) {
		return `${gap}<${tag}${attrStr}/>`;
	}

	// Single text child → inline formatting
	if (children.length === 1 && typeof children[0] === "string" && !children[0].includes("<")) {
		return `${gap}<${tag}${attrStr}>${children[0].trimEnd()}</${tag}>`;
	}

	// Has children → recursive rendering
	const inner = children
		.map(c => typeof c === "string" ? `${"  ".repeat(level + 1)}${c.trimEnd()}` : buildXml(c, level + 1))
		.join("\n");

	return `${gap}<${tag}${attrStr}>\n${inner}\n${gap}</${tag}>`;
}

/**
 * Cell XML Structure Documentation
 *
 * Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
 *
 * Parent Elements:
 * - row
 *
 * Child Elements:
 * - extLst (Future Feature Data Storage Area)
 * - f (Formula)
 * - is (Rich Text Inline)
 * - v (Cell Value)
 *
 * Attributes:
 * - cm (Cell Metadata Index): The zero-based index of the cell metadata record associated with this cell.
 *   Metadata information is found in the Metadata Part. Cell metadata is extra information stored at the
 *   cell level, and is attached to the cell (travels through moves, copy / paste, clear, etc).
 *   Cell metadata is not accessible via formula reference.
 *   The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype.
 *
 * - ph (Show Phonetic): A Boolean value indicating if the spreadsheet application should show phonetic
 *   information. Phonetic information is displayed in the same cell across the top of the cell and serves
 *   as a 'hint' which indicates how the text should be pronounced. This should only be used for East Asian languages.
 *   The possible values for this attribute are defined by the W3C XML Schema boolean datatype.
 *
 * - r (Reference): An A1 style reference to the location of this cell
 *   The possible values for this attribute are defined by the ST_CellRef simple type.
 *
 * - s (Style Index): The index of this cell's style. Style records are stored in the Styles Part.
 *   The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype.
 *
 * - t (Cell Data Type): An enumeration representing the cell's data type.
 *   The possible values for this attribute are defined by the ST_CellType simple type.
 *
 * - vm (Value Metadata Index): The zero-based index of the value metadata record associated with this cell's value.
 *   Metadata records are stored in the Metadata Part. Value metadata is extra information stored at the cell level,
 *   but associated with the value rather than the cell itself. Value metadata is accessible via formula reference.
 *   The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype.
 */
