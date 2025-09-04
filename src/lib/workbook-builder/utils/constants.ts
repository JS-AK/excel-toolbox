// Content Types
export const CONTENT_TYPES = {
	APP: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
	CORE: "application/vnd.openxmlformats-package.core-properties+xml",
	RELATIONSHIPS: "application/vnd.openxmlformats-package.relationships+xml",
	SHARED_STRINGS: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
	STYLES: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
	THEME: "application/vnd.openxmlformats-officedocument.theme+xml",
	WORKBOOK: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
	WORKSHEET: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	XML: "application/xml",
} as const;

// Default file paths
export const FILE_PATHS = {
	APP: "docProps/app.xml",
	CONTENT_TYPES: "[Content_Types].xml",
	CORE: "docProps/core.xml",
	RELS: "_rels/.rels",
	SHARED_STRINGS: "xl/sharedStrings.xml",
	STYLES: "xl/styles.xml",
	THEME: "xl/theme/theme1.xml",
	WORKBOOK: "xl/workbook.xml",
	WORKBOOK_RELS: "xl/_rels/workbook.xml.rels",
	WORKSHEET: "xl/worksheets/sheet1.xml",
} as const;

// Relationship Types
export const RELATIONSHIP_TYPES = {
	APP: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
	CORE: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
	OFFICE_DOCUMENT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	SHARED_STRINGS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	STYLES: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	THEME: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
	WORKSHEET: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
} as const;

// XML Declarations
export const XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

// XML Namespaces
export const XML_NAMESPACES = {
	CONTENT_TYPES: "http://schemas.openxmlformats.org/package/2006/content-types",
	CORE_PROPERTIES: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
	DC: "http://purl.org/dc/elements/1.1/",
	DCMITYPE: "http://purl.org/dc/dcmitype/",
	DCTERMS: "http://purl.org/dc/terms/",
	DOC_PROPS_VTYPES: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
	DRAWINGML: "http://schemas.openxmlformats.org/drawingml/2006/main",
	EXTENDED_PROPERTIES: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
	OFFICE_DOCUMENT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	PACKAGE_RELATIONSHIPS: "http://schemas.openxmlformats.org/package/2006/relationships",
	SPREADSHEET_ML: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
	XSI: "http://www.w3.org/2001/XMLSchema-instance",
} as const;
