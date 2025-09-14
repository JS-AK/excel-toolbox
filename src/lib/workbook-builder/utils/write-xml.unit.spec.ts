import { PassThrough } from "node:stream";
import { createWriteStream } from "node:fs";

import { describe, expect, it } from "vitest";

import type { XmlNode } from "./build-xml.js";
import { writeXml } from "./write-xml.js";

describe("writeXml", () => {
	it("should write simple self-closing tag to stream", async () => {
		const node: XmlNode = {
			tag: "br",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<br/>\n");
	});

	it("should write self-closing tag with attributes to stream", async () => {
		const node: XmlNode = {
			tag: "input",

			attrs: {
				name: "username",
				required: "true",
				type: "text",
			},
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<input name=\"username\" required=\"true\" type=\"text\"/>\n");
	});

	it("should write tag with single text child inline", async () => {
		const node: XmlNode = {
			children: ["Hello World"],
			tag: "title",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<title>Hello World</title>\n");
	});

	it("should write tag with single text child and attributes inline", async () => {
		const node: XmlNode = {
			attrs: {
				class: "header",
				id: "main-title",
			},
			children: ["Welcome"],
			tag: "h1",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<h1 class=\"header\" id=\"main-title\">Welcome</h1>\n");
	});

	it("should write nested structure with proper indentation", async () => {
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

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
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

		expect(result).toBe(expected + "\n");
	});

	it("should handle mixed content (text and elements)", async () => {
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

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		const expected = `<div>
  Text before 
  <strong>bold text</strong>
   and after
</div>`;

		expect(result).toBe(expected + "\n");
	});

	it("should handle empty children array", async () => {
		const node: XmlNode = {
			children: [],
			tag: "div",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<div/>\n");
	});

	it("should filter out undefined and null attributes", async () => {
		const node: XmlNode = {
			attrs: {
				name: "test",
				placeholder: null,
				required: "true",
				type: "text",
				value: undefined,
			},
			tag: "input",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<input name=\"test\" required=\"true\" type=\"text\"/>\n");
	});

	it("should handle numeric attributes", async () => {
		const node: XmlNode = {
			attrs: {
				height: 200,
				id: 123,
				width: 100,
			},
			tag: "div",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toBe("<div height=\"200\" id=\"123\" width=\"100\"/>\n");
	});

	it("should handle custom indentation level", async () => {
		const node: XmlNode = {
			children: [
				{
					children: ["test"],
					tag: "span",
				},
			],
			tag: "div",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream, 2);

		const result = Buffer.concat(chunks).toString();
		const expected = `    <div>
      <span>test</span>
    </div>`;

		expect(result).toBe(expected + "\n");
	});

	it("should handle text content with HTML-like characters", async () => {
		const node: XmlNode = {
			children: ["Text with < and > characters"],
			tag: "div",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		const expected = `<div>
  Text with < and > characters
</div>`;

		expect(result).toBe(expected + "\n");
	});

	it("should handle deeply nested structure", async () => {
		const node: XmlNode = {
			children: [
				{
					children: [
						{
							children: [
								{
									children: ["Deep content"],
									tag: "level4",
								},
							],
							tag: "level3",
						},
					],
					tag: "level2",
				},
			],
			tag: "level1",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		const expected = `<level1>
  <level2>
    <level3>
      <level4>Deep content</level4>
    </level3>
  </level2>
</level1>`;

		expect(result).toBe(expected + "\n");
	});

	it("should handle large dataset with chunking", async () => {
		const children: XmlNode[] = [];
		for (let i = 1; i <= 2000; i++) {
			children.push({
				attrs: { id: i },
				children: [`Item ${i}`],
				tag: "item",
			});
		}

		const node: XmlNode = {
			children,
			tag: "root",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		const startTime = performance.now();
		await writeXml(node, stream, 0, 100); // Small chunk size for testing
		const endTime = performance.now();

		const result = Buffer.concat(chunks).toString();
		expect(result).toContain("<root>");
		expect(result).toContain("</root>");
		expect(result).toContain("<item id=\"1\">Item 1</item>");
		expect(result).toContain("<item id=\"2000\">Item 2000</item>");

		// Should complete in reasonable time
		expect(endTime - startTime).toBeLessThan(5000);
	});

	it("should work with file write stream", async () => {
		const node: XmlNode = {
			children: [
				{ children: ["Test Document"], tag: "title" },
				{ children: ["This is test content"], tag: "content" },
			],
			tag: "document",
		};

		// Create a temporary file stream
		const tempFile = "temp-test-write.xml";
		const writeStream = createWriteStream(tempFile);

		await writeXml(node, writeStream);
		writeStream.end();

		// Wait for stream to finish
		await new Promise<void>((resolve, reject) => {
			writeStream.on("finish", () => resolve());
			writeStream.on("error", reject);
		});
		const fs = await import("node:fs/promises");
		const content = await fs.readFile(tempFile, "utf-8");
		expect(content).toContain("<document>");
		expect(content).toContain("<title>Test Document</title>");
		expect(content).toContain("<content>This is test content</content>");
		expect(content).toContain("</document>");

		// Clean up
		await fs.unlink(tempFile);
	});

	it("should handle complex real-world example", async () => {
		const node: XmlNode = {
			attrs: { xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main" },
			children: [
				{
					children: [
						{
							attrs: { r: 1 },
							children: [
								{
									attrs: { r: "A1", t: "s" },
									children: [
										{
											children: ["0"],
											tag: "v",
										},
									],
									tag: "c",
								},
								{
									attrs: { r: "B1", t: "n" },
									children: [
										{
											children: ["42"],
											tag: "v",
										},
									],
									tag: "c",
								},
							],
							tag: "row",
						},
					],
					tag: "sheetData",
				},
			],
			tag: "worksheet",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		await writeXml(node, stream);

		const result = Buffer.concat(chunks).toString();
		expect(result).toContain("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
		expect(result).toContain("<sheetData>");
		expect(result).toContain("<row r=\"1\">");
		expect(result).toContain("<c r=\"A1\" t=\"s\">");
		expect(result).toContain("<v>0</v>");
		expect(result).toContain("<c r=\"B1\" t=\"n\">");
		expect(result).toContain("<v>42</v>");
	});

	it("should handle custom chunk size", async () => {
		const children: XmlNode[] = [];
		for (let i = 1; i <= 10; i++) {
			children.push({
				children: [`Item ${i}`],
				tag: "item",
			});
		}

		const node: XmlNode = {
			children,
			tag: "root",
		};

		const stream = new PassThrough();
		const chunks: Buffer[] = [];

		stream.on("data", (chunk) => {
			chunks.push(chunk);
		});

		// Use chunk size of 3 to test chunking behavior
		await writeXml(node, stream, 0, 3);

		const result = Buffer.concat(chunks).toString();
		expect(result).toContain("<root>");
		expect(result).toContain("</root>");
		expect(result).toContain("<item>Item 1</item>");
		expect(result).toContain("<item>Item 10</item>");
	});

	it("should handle stream errors gracefully", async () => {
		const node: XmlNode = {
			children: ["test"],
			tag: "root",
		};

		const stream = new PassThrough();
		let errorThrown = false;

		// Simulate stream error
		stream.on("error", () => {
			errorThrown = true;
		});

		// Write some data then destroy the stream
		stream.write("some data");
		stream.destroy(new Error("Test error"));

		try {
			await writeXml(node, stream);
		} catch (error) {
			// Expected to throw due to destroyed stream
			expect(error).toBeDefined();
		}

		// The errorThrown flag is set by the stream's "error" event, but in Node.js,
		// destroying a stream with an error does not always guarantee the "error" event
		// will be emitted synchronously before the catch block runs. To ensure the event
		// loop processes the "error" event, wait for the next tick.
		await new Promise(process.nextTick);

		expect(errorThrown).toBe(true);
	});
});
