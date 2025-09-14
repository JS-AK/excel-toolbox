import fs from "node:fs";
import fsPromises from "node:fs/promises";
import path from "node:path";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";

/**
 * Пишет содержимое `sharedStrings.xml` в поток.
 *
 * @param stream - Writable поток (например, fs.WriteStream или Zip entry stream)
 * @param strings - Массив уникальных строк, используемых в книге
 */
export async function writeSharedStringsXml(destination: string, strings: string[] = []): Promise<void> {
	checkRam("writeSharedStringsXml 1");

	// create with folder
	await fsPromises.mkdir(path.dirname(destination), { recursive: true });

	checkRam("writeSharedStringsXml 2");

	const stream = fs.createWriteStream(destination, { encoding: "utf-8" });

	checkRam("writeSharedStringsXml 3");

	try {
		const escapeXml = (str: string) =>
			str
				.replace(/&/g, "&amp;")
				.replace(/</g, "&lt;")
				.replace(/>/g, "&gt;")
				.replace(/"/g, "&quot;")
				.replace(/'/g, "&apos;");

		// Заголовок документа
		stream.write(XML_DECLARATION + "\n");
		stream.write(`<sst xmlns="${XML_NAMESPACES.SPREADSHEET_ML}" count="${strings.length}" uniqueCount="${strings.length}">\n`);

		checkRam("writeWorksheetXml 4");

		// Основные строки
		for (const s of strings) {
			const preserve = /^\s|\s$/.test(s) ? " xml:space=\"preserve\"" : "";
			const siXml = `<si><t${preserve}>${escapeXml(s)}</t></si>\n`;

			// пишем с контролем backpressure
			if (!stream.write(siXml)) {
				await new Promise<void>(resolve => stream.once("drain", () => resolve()));
			}
		}

		checkRam("writeWorksheetXml 5");

		// Закрывающий тег
		stream.write("</sst>");
	} finally {
		stream.end();
	}

	return new Promise((resolve, reject) => {
		stream.on("error", reject);
		stream.on("finish", resolve);
	});
}

function checkRam(message: string) {
	const used = process.memoryUsage().heapTotal / 1024 / 1024;
	const timeStamp = new Date().toISOString().slice(0, 22);

	console.log(timeStamp, message, `Used: ${used} MB`);
}
