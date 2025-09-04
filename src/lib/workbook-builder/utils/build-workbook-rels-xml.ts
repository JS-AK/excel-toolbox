import { RELATIONSHIP_TYPES, XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

export function buildWorkbookRels(sheetsCount: number): string {
	// Создаем relationship для листов
	const sheetRels = Array.from({ length: sheetsCount }, (_, i) => ({
		attrs: {
			Id: `rId${i + 1}`,
			Target: `worksheets/sheet${i + 1}.xml`,
			Type: RELATIONSHIP_TYPES.WORKSHEET,
		},
		tag: "Relationship",
	}));

	// Id для стилей и sharedStrings идут после листов
	const stylesRel = {
		attrs: {
			Id: `rId${sheetsCount + 1}`,
			Target: "styles.xml",
			Type: RELATIONSHIP_TYPES.STYLES,
		},
		tag: "Relationship",
	};

	const themeRel = {
		attrs: {
			Id: `rId${sheetsCount + 2}`,
			Target: "theme/theme1.xml",
			Type: RELATIONSHIP_TYPES.THEME,
		},
		tag: "Relationship",
	};

	const sharedStringsRel = {
		attrs: {
			Id: `rId${sheetsCount + 3}`,
			Target: "sharedStrings.xml",
			Type: RELATIONSHIP_TYPES.SHARED_STRINGS,
		},
		tag: "Relationship",
	};

	const allRels = [...sheetRels, stylesRel, themeRel, sharedStringsRel];

	return [
		XML_DECLARATION,
		buildXml({
			attrs: { xmlns: XML_NAMESPACES.PACKAGE_RELATIONSHIPS },
			children: allRels,
			tag: "Relationships",
		}),
	].join("\n");
}
