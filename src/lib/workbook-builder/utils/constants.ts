// Content Types
export const CONTENT_TYPES = {
	RELATIONSHIPS: "application/vnd.openxmlformats-package.relationships+xml",
	SHARED_STRINGS: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
	STYLES: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
	WORKBOOK: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
	WORKSHEET: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	XML: "application/xml",
} as const;

// Default file paths
export const FILE_PATHS = {
	CONTENT_TYPES: "[Content_Types].xml",
	RELS: "_rels/.rels",
	SHARED_STRINGS: "xl/sharedStrings.xml",
	STYLES: "xl/styles.xml",
	WORKBOOK: "xl/workbook.xml",
	WORKBOOK_RELS: "xl/_rels/workbook.xml.rels",
	WORKSHEET: "xl/worksheets/sheet1.xml",
} as const;

// Relationship Types
export const RELATIONSHIP_TYPES = {
	OFFICE_DOCUMENT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	SHARED_STRINGS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	STYLES: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	WORKSHEET: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
} as const;

// XML Declarations
export const XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

// XML Namespaces
export const XML_NAMESPACES = {
	CONTENT_TYPES: "http://schemas.openxmlformats.org/package/2006/content-types",
	OFFICE_DOCUMENT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	PACKAGE_RELATIONSHIPS: "http://schemas.openxmlformats.org/package/2006/relationships",
	SPREADSHEET_ML: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
} as const;
