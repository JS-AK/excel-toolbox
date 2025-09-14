import { Writable } from "node:stream";

import { XmlNode } from "../types/index.js";

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
