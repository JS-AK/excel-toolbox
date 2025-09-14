import { Writable } from "node:stream";

type XmlNode = {
	/** The XML tag name */
	tag: string;
	/** Optional attributes for the XML element */
	attrs?: Record<string, string | number | undefined | null>;
	/** Child elements or text content */
	children?: (string | XmlNode)[];
};

export async function writeXml(
	node: XmlNode,
	stream: Writable,
	level = 0,
	chunkSize = 100,
): Promise<void> {
	const { attrs = {}, children = [], tag } = node;

	const attrStr = Object.entries(attrs)
		.filter(([, v]) => v !== undefined && v !== null)
		.map(([k, v]) => ` ${k}="${v}"`)
		.join("");

	const gap = "  ".repeat(level);

	// No children → self-closing tag
	if (!children.length) {
		stream.write(`${gap}<${tag}${attrStr}/>\n`);

		return;
	}

	// Single text child → inline formatting
	if (children.length === 1 && typeof children[0] === "string" && !children[0].includes("<")) {
		stream.write(`${gap}<${tag}${attrStr}>${children[0]}</${tag}>\n`);

		return;
	}

	// Has children → recursive streaming
	stream.write(`${gap}<${tag}${attrStr}>\n`);

	let processed = 0;

	for (const c of children) {
		if (typeof c === "string") {
			stream.write(`${"  ".repeat(level + 1)}${c}\n`);
		} else {
			await writeXml(c, stream, level + 1, chunkSize);
		}

		processed++;
		if (processed % chunkSize === 0) {
			await new Promise(resolve => setImmediate(resolve));
		}
	}

	stream.write(`${gap}</${tag}>\n`);
}

/**
 * Builds XML and writes it directly to a writable stream.
 * This is memory-efficient for large XML documents as it streams the output
 * instead of building the entire string in memory.
 *
 * @param node - The XML node to convert to string
 * @param stream - The writable stream to write XML to
 * @param level - The indentation level for formatting (default: 0)
 * @param chunkSize - Number of children to process before yielding control (default: 1000)
 *
 * @returns Promise that resolves when XML is fully written to stream
 */
export async function writeXml2(
	node: XmlNode,
	stream: Writable,
	level = 0,
	chunkSize = 1000,
): Promise<void> {
	const { attrs = {}, children = [], tag } = node;

	const attrStr = Object.entries(attrs)
		.filter(([, v]) => v !== undefined && v !== null)
		.map(([k, v]) => ` ${k}="${v}"`)
		.join("");

	const gap = "  ".repeat(level);

	// No children → self-closing tag
	if (!children.length) {
		stream.write(`${gap}<${tag}${attrStr}/>\n`);
		return;
	}

	// Single text child → inline formatting
	if (children.length === 1 && typeof children[0] === "string" && !children[0].includes("<")) {
		stream.write(`${gap}<${tag}${attrStr}>${children[0].trimEnd()}</${tag}>\n`);
		return;
	}

	// Has children → recursive streaming
	stream.write(`${gap}<${tag}${attrStr}>\n`);

	for (let i = 0; i < children.length; i += chunkSize) {
		const chunk = children.slice(i, i + chunkSize);

		for (const c of chunk) {
			if (typeof c === "string") {
				stream.write(`${"  ".repeat(level + 1)}${c.trimEnd()}\n`);
			} else {
				await writeXml(c, stream, level + 1, chunkSize);
			}
		}

		// Yield control to event loop every chunkSize children
		if (i + chunkSize < children.length) {
			await new Promise((resolve) => setImmediate(resolve));
		}
	}

	stream.write(`${gap}</${tag}>\n`);
}
