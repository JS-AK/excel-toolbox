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
