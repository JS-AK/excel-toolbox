import { describe, expect, it } from "vitest";

import type { XmlNode } from "./build-xml.js";
import { buildXml } from "./build-xml.js";

describe("buildXml", () => {
	it("should build simple self-closing tag", () => {
		const node: XmlNode = {
			tag: "br",
		};

		const result = buildXml(node);
		expect(result).toBe("<br/>");
	});

	it("should build self-closing tag with attributes", () => {
		const node: XmlNode = {
			attrs: {
				name: "username",
				required: "true",
				type: "text",
			},
			tag: "input",
		};

		const result = buildXml(node);
		expect(result).toBe("<input name=\"username\" required=\"true\" type=\"text\"/>");
	});

	it("should build tag with single text child inline", () => {
		const node: XmlNode = {
			children: ["Hello World"],
			tag: "title",
		};

		const result = buildXml(node);
		expect(result).toBe("<title>Hello World</title>");
	});

	it("should build tag with single text child and attributes inline", () => {
		const node: XmlNode = {
			attrs: {
				class: "header",
				id: "main-title",
			},
			children: ["Welcome"],
			tag: "h1",
		};

		const result = buildXml(node);
		expect(result).toBe("<h1 class=\"header\" id=\"main-title\">Welcome</h1>");
	});

	it("should build nested structure with proper indentation", () => {
		const node: XmlNode = {
			attrs: { version: "1.0" },
			children: [
				{
					children: [
						{
							children: ["My App"],
							tag: "title",
						},
						{
							children: ["Home", "About"],
							tag: "nav",
						},
					],
					tag: "header",
				},
				{
					children: [
						{
							children: ["Content here"],
							tag: "p",
						},
					],
					tag: "main",
				},
			],
			tag: "root",
		};

		const result = buildXml(node);
		const expected = `<root version="1.0">
  <header>
    <title>My App</title>
    <nav>
      Home
      About
    </nav>
  </header>
  <main>
    <p>Content here</p>
  </main>
</root>`;

		expect(result).toBe(expected);
	});

	it("should handle mixed content (text and elements)", () => {
		const node: XmlNode = {
			children: [
				"Text before ",
				{
					children: ["bold text"],
					tag: "strong",
				},
				" and after",
			],
			tag: "div",
		};

		const result = buildXml(node);
		const expected = `<div>
  Text before
  <strong>bold text</strong>
   and after
</div>`;

		expect(result).toBe(expected);
	});

	it("should handle empty children array", () => {
		const node: XmlNode = {
			tag: "div",

			children: [],
		};

		const result = buildXml(node);
		expect(result).toBe("<div/>");
	});

	it("should filter out undefined and null attributes", () => {
		const node: XmlNode = {
			tag: "input",

			attrs: {
				name: "test",
				placeholder: null,
				required: "true",
				type: "text",
				value: undefined,
			},
		};

		const result = buildXml(node);
		expect(result).toBe("<input name=\"test\" required=\"true\" type=\"text\"/>");
	});

	it("should handle numeric attributes", () => {
		const node: XmlNode = {
			tag: "div",

			attrs: {
				height: 200,
				id: 123,
				width: 100,
			},
		};

		const result = buildXml(node);
		expect(result).toBe("<div height=\"200\" id=\"123\" width=\"100\"/>");
	});

	it("should handle custom indentation level", () => {
		const node: XmlNode = {
			children: [
				{
					tag: "span",

					children: ["test"],
				},
			],
			tag: "div",
		};

		const result = buildXml(node, 2);
		const expected = `    <div>
      <span>test</span>
    </div>`;

		expect(result).toBe(expected);
	});

	it("should handle text content with HTML-like characters", () => {
		const node: XmlNode = {
			tag: "div",

			children: ["Text with < and > characters"],
		};

		const result = buildXml(node);
		const expected = `<div>
  Text with < and > characters
</div>`;

		expect(result).toBe(expected);
	});

	it("should handle deeply nested structure", () => {
		const node: XmlNode = {
			tag: "level1",

			children: [
				{
					tag: "level2",

					children: [
						{
							tag: "level3",

							children: [
								{
									tag: "level4",

									children: ["Deep content"],
								},
							],
						},
					],
				},
			],
		};

		const result = buildXml(node);
		const expected = `<level1>
  <level2>
    <level3>
      <level4>Deep content</level4>
    </level3>
  </level2>
</level1>`;

		expect(result).toBe(expected);
	});

	it("should handle complex real-world example", () => {
		const node: XmlNode = {
			tag: "worksheet",

			attrs: { "xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main" },
			children: [
				{
					tag: "sheetData",

					children: [
						{
							tag: "row",

							attrs: { r: 1 },
							children: [
								{
									tag: "c",

									attrs: { r: "A1", t: "s" },
									children: [
										{
											tag: "v",

											children: ["0"],
										},
									],
								},
								{
									tag: "c",

									attrs: { r: "B1", t: "n" },
									children: [
										{
											tag: "v",

											children: ["42"],
										},
									],
								},
							],
						},
					],
				},
			],
		};

		const result = buildXml(node);
		expect(result).toContain("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
		expect(result).toContain("<row r=\"1\">");
		expect(result).toContain("<c r=\"A1\" t=\"s\">");
		expect(result).toContain("<v>0</v>");
		expect(result).toContain("<c r=\"B1\" t=\"n\">");
		expect(result).toContain("<v>42</v>");
	});
});
