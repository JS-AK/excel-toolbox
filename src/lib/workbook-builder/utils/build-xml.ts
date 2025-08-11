export type XmlNode = {
	tag: string;
	attrs?: Record<string, string>;
	children?: (string | XmlNode)[];
};

export function buildXml(node: XmlNode, level = 0): string {
	const { attrs = {}, children = [], tag } = node;

	const attrStr = Object.entries(attrs)
		.map(([k, v]) => ` ${k}="${v}"`)
		.join("");

	const gap = "  ".repeat(level);

	// нет детей → самозакрывающийся
	if (!children.length) {
		return `${gap}<${tag}${attrStr}/>`;
	}

	// единственный текст → инлайн
	if (children.length === 1 && typeof children[0] === "string" && !children[0].includes("<")) {
		return `${gap}<${tag}${attrStr}>${children[0]}</${tag}>`;
	}

	// есть дети → рекурсивно рендерим
	const inner = children
		.map(c => typeof c === "string" ? `${"  ".repeat(level + 1)}${c}` : buildXml(c, level + 1))
		.join("\n");

	return `${gap}<${tag}${attrStr}>\n${inner}\n${gap}</${tag}>`;
}
