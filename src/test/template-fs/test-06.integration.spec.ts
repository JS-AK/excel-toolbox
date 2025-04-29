import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import * as Xml from "../../lib/xml/index.js";
import * as Zip from "../../lib/zip/index.js";
import { TemplateFs } from "../../lib/template/template-fs.js";
import { validateWorksheetXml } from "../../lib/template/utils/validate-worksheet-xml.js";

const TEMP_DIR = path.resolve(process.cwd(), "src", "test", "template-fs", "temp");
const ASSETS_DIR = path.resolve(process.cwd(), "src", "test", "template-fs", "assets");
const INPUT_FILE = path.resolve(ASSETS_DIR, "input-test-06.xlsx");
const OUTPUT_FILE = path.resolve(ASSETS_DIR, "output-test-06.xlsx");

describe("TemplateFs integration test", () => {
	it("should copySheet and substitute with save", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		await template.copySheet("Sheet1", "Sheet2");

		await template.substitute("Sheet1", {
			users: [
				{ age: 30, city: "New York", counter: 1, name: "John", surname: "Doe" },
				{ age: 31, city: "Los Angeles", counter: 2, name: "Jane", surname: "Smith" },
				{ age: 32, city: "Chicago", counter: 3, name: "Jim", surname: "Beam" },
				{ age: 33, city: "San Francisco", counter: 4, name: "Jill", surname: "Dow" },
				{ age: 34, city: "Miami", counter: 5, name: "Jack", surname: "Eck" },
			],
		});

		await template.substitute("Sheet2", {
			users: [
				{ age: 10, city: "Moscow", counter: 1, name: "Vladimir", surname: "Efimov" },
				{ age: 20, city: "Moscow", counter: 2, name: "Ivan", surname: "Ivanov" },
				{ age: 30, city: "Moscow", counter: 3, name: "Petr", surname: "Petrov" },
				{ age: 40, city: "Moscow", counter: 4, name: "Sergey", surname: "Sidorov" },
				{ age: 50, city: "Moscow", counter: 5, name: "Dmitry", surname: "Fedorov" },
			],
		});

		// Save the buffer to a file
		const buffer = await template.save();

		// Save the buffer to a file
		await fs.writeFile(OUTPUT_FILE, buffer);

		expect(buffer).toBeDefined();

		// Read the rebuilt file
		const original = await fs.readFile(INPUT_FILE);
		const rebuilt = await fs.readFile(OUTPUT_FILE);

		// Read the rebuilt zip file
		const rebuiltZip = await Zip.read(rebuilt);
		const originalZip = await Zip.read(original);

		// Check that the rebuilt zip file has the same keys as the original zip file
		const origKeys = Object.keys(originalZip).sort();
		const rebuiltKeys = Object.keys(rebuiltZip).sort();

		const updatedKeys = [
			...origKeys,
			"xl/worksheets/sheet2.xml", // new sheet added by copy
		];

		// Check that the rebuilt zip file has the same keys as the original zip file
		expect(rebuiltKeys).toEqual(updatedKeys);

		// find new rows in the rebuilt zip file
		const sheet1Rebuilt = rebuiltZip["xl/worksheets/sheet1.xml"].toString();
		const sheet2Rebuilt = rebuiltZip["xl/worksheets/sheet2.xml"].toString();

		const sheet1RowsData = await Xml.extractRowsFromSheet(sheet1Rebuilt);
		const sheet2RowsData = await Xml.extractRowsFromSheet(sheet2Rebuilt);

		expect(sheet1RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\" spans=\"1:5\"><c r=\"A2\" t=\"s\"><v>8</v></c><c r=\"B2\" t=\"s\"><v>7</v></c><c r=\"C2\" t=\"s\"><v>6</v></c><c r=\"D2\" t=\"s\"><v>5</v></c><c r=\"E2\" t=\"s\"><v>4</v></c></row>",
			"<row r=\"3\" spans=\"1:5\"><c r=\"A3\" t=\"s\"><v>10</v></c><c r=\"B3\" t=\"s\"><v>11</v></c><c r=\"C3\" t=\"s\"><v>12</v></c><c r=\"D3\" t=\"s\"><v>13</v></c><c r=\"E3\" t=\"s\"><v>14</v></c></row>",
			"<row r=\"4\" spans=\"1:5\"><c r=\"A4\" t=\"s\"><v>15</v></c><c r=\"B4\" t=\"s\"><v>16</v></c><c r=\"C4\" t=\"s\"><v>17</v></c><c r=\"D4\" t=\"s\"><v>18</v></c><c r=\"E4\" t=\"s\"><v>19</v></c></row>",
			"<row r=\"5\" spans=\"1:5\"><c r=\"A5\" t=\"s\"><v>20</v></c><c r=\"B5\" t=\"s\"><v>21</v></c><c r=\"C5\" t=\"s\"><v>22</v></c><c r=\"D5\" t=\"s\"><v>23</v></c><c r=\"E5\" t=\"s\"><v>24</v></c></row>",
			"<row r=\"6\" spans=\"1:5\"><c r=\"A6\" t=\"s\"><v>25</v></c><c r=\"B6\" t=\"s\"><v>26</v></c><c r=\"C6\" t=\"s\"><v>27</v></c><c r=\"D6\" t=\"s\"><v>28</v></c><c r=\"E6\" t=\"s\"><v>29</v></c></row>",
			"<row r=\"7\" spans=\"1:5\"><c r=\"A7\" t=\"s\"><v>30</v></c><c r=\"B7\" t=\"s\"><v>31</v></c><c r=\"C7\" t=\"s\"><v>32</v></c><c r=\"D7\" t=\"s\"><v>33</v></c><c r=\"E7\" t=\"s\"><v>34</v></c></row>",
			"<row r=\"9\" spans=\"1:5\"><c r=\"A9\"><v>5</v></c><c r=\"B9\"><v>4</v></c><c r=\"C9\"><v>3</v></c><c r=\"D9\"><v>2</v></c><c r=\"E9\"><v>1</v></c></row>",
		]);

		expect(sheet2RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\" spans=\"1:5\"><c r=\"A2\" t=\"s\"><v>8</v></c><c r=\"B2\" t=\"s\"><v>7</v></c><c r=\"C2\" t=\"s\"><v>6</v></c><c r=\"D2\" t=\"s\"><v>5</v></c><c r=\"E2\" t=\"s\"><v>4</v></c></row>",
			"<row r=\"3\" spans=\"1:5\"><c r=\"A3\" t=\"s\"><v>35</v></c><c r=\"B3\" t=\"s\"><v>36</v></c><c r=\"C3\" t=\"s\"><v>37</v></c><c r=\"D3\" t=\"s\"><v>13</v></c><c r=\"E3\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"4\" spans=\"1:5\"><c r=\"A4\" t=\"s\"><v>39</v></c><c r=\"B4\" t=\"s\"><v>40</v></c><c r=\"C4\" t=\"s\"><v>41</v></c><c r=\"D4\" t=\"s\"><v>18</v></c><c r=\"E4\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"5\" spans=\"1:5\"><c r=\"A5\" t=\"s\"><v>42</v></c><c r=\"B5\" t=\"s\"><v>43</v></c><c r=\"C5\" t=\"s\"><v>12</v></c><c r=\"D5\" t=\"s\"><v>23</v></c><c r=\"E5\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"6\" spans=\"1:5\"><c r=\"A6\" t=\"s\"><v>44</v></c><c r=\"B6\" t=\"s\"><v>45</v></c><c r=\"C6\" t=\"s\"><v>46</v></c><c r=\"D6\" t=\"s\"><v>28</v></c><c r=\"E6\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"7\" spans=\"1:5\"><c r=\"A7\" t=\"s\"><v>47</v></c><c r=\"B7\" t=\"s\"><v>48</v></c><c r=\"C7\" t=\"s\"><v>49</v></c><c r=\"D7\" t=\"s\"><v>33</v></c><c r=\"E7\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"9\" spans=\"1:5\"><c r=\"A9\"><v>5</v></c><c r=\"B9\"><v>4</v></c><c r=\"C9\"><v>3</v></c><c r=\"D9\"><v>2</v></c><c r=\"E9\"><v>1</v></c></row>",
		]);

		const validationResult1 = validateWorksheetXml(sheet1Rebuilt);
		const validationResult2 = validateWorksheetXml(sheet2Rebuilt);

		expect(validationResult1.isValid).toBe(true);
		expect(validationResult2.isValid).toBe(true);
	});

	it("should copySheet and substitute with saveStream", async () => {
		const template = await TemplateFs.from({
			destination: TEMP_DIR,
			source: INPUT_FILE,
		});

		await template.copySheet("Sheet1", "Sheet2");

		await template.substitute("Sheet1", {
			users: [
				{ age: 30, city: "New York", counter: 1, name: "John", surname: "Doe" },
				{ age: 31, city: "Los Angeles", counter: 2, name: "Jane", surname: "Smith" },
				{ age: 32, city: "Chicago", counter: 3, name: "Jim", surname: "Beam" },
				{ age: 33, city: "San Francisco", counter: 4, name: "Jill", surname: "Dow" },
				{ age: 34, city: "Miami", counter: 5, name: "Jack", surname: "Eck" },
			],
		});

		await template.substitute("Sheet2", {
			users: [
				{ age: 10, city: "Moscow", counter: 1, name: "Vladimir", surname: "Efimov" },
				{ age: 20, city: "Moscow", counter: 2, name: "Ivan", surname: "Ivanov" },
				{ age: 30, city: "Moscow", counter: 3, name: "Petr", surname: "Petrov" },
				{ age: 40, city: "Moscow", counter: 4, name: "Sergey", surname: "Sidorov" },
				{ age: 50, city: "Moscow", counter: 5, name: "Dmitry", surname: "Fedorov" },
			],
		});

		// Save the buffer to a file
		await template.saveStream(fsSync.createWriteStream(OUTPUT_FILE));

		// Save the buffer to a file
		const buffer = await fs.readFile(OUTPUT_FILE);

		expect(buffer).toBeDefined();

		// Read the rebuilt file
		const original = await fs.readFile(INPUT_FILE);
		const rebuilt = await fs.readFile(OUTPUT_FILE);

		// Read the rebuilt zip file
		const rebuiltZip = await Zip.read(rebuilt);
		const originalZip = await Zip.read(original);

		// Check that the rebuilt zip file has the same keys as the original zip file
		const origKeys = Object.keys(originalZip).sort();
		const rebuiltKeys = Object.keys(rebuiltZip).sort();

		const updatedKeys = [
			...origKeys,
			"xl/worksheets/sheet2.xml", // new sheet added by copy
		];

		// Check that the rebuilt zip file has the same keys as the original zip file
		expect(rebuiltKeys).toEqual(updatedKeys);

		// find new rows in the rebuilt zip file
		const sheet1Rebuilt = rebuiltZip["xl/worksheets/sheet1.xml"].toString();
		const sheet2Rebuilt = rebuiltZip["xl/worksheets/sheet2.xml"].toString();

		const sheet1RowsData = await Xml.extractRowsFromSheet(sheet1Rebuilt);
		const sheet2RowsData = await Xml.extractRowsFromSheet(sheet2Rebuilt);

		expect(sheet1RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\" spans=\"1:5\"><c r=\"A2\" t=\"s\"><v>8</v></c><c r=\"B2\" t=\"s\"><v>7</v></c><c r=\"C2\" t=\"s\"><v>6</v></c><c r=\"D2\" t=\"s\"><v>5</v></c><c r=\"E2\" t=\"s\"><v>4</v></c></row>",
			"<row r=\"3\" spans=\"1:5\"><c r=\"A3\" t=\"s\"><v>10</v></c><c r=\"B3\" t=\"s\"><v>11</v></c><c r=\"C3\" t=\"s\"><v>12</v></c><c r=\"D3\" t=\"s\"><v>13</v></c><c r=\"E3\" t=\"s\"><v>14</v></c></row>",
			"<row r=\"4\" spans=\"1:5\"><c r=\"A4\" t=\"s\"><v>15</v></c><c r=\"B4\" t=\"s\"><v>16</v></c><c r=\"C4\" t=\"s\"><v>17</v></c><c r=\"D4\" t=\"s\"><v>18</v></c><c r=\"E4\" t=\"s\"><v>19</v></c></row>",
			"<row r=\"5\" spans=\"1:5\"><c r=\"A5\" t=\"s\"><v>20</v></c><c r=\"B5\" t=\"s\"><v>21</v></c><c r=\"C5\" t=\"s\"><v>22</v></c><c r=\"D5\" t=\"s\"><v>23</v></c><c r=\"E5\" t=\"s\"><v>24</v></c></row>",
			"<row r=\"6\" spans=\"1:5\"><c r=\"A6\" t=\"s\"><v>25</v></c><c r=\"B6\" t=\"s\"><v>26</v></c><c r=\"C6\" t=\"s\"><v>27</v></c><c r=\"D6\" t=\"s\"><v>28</v></c><c r=\"E6\" t=\"s\"><v>29</v></c></row>",
			"<row r=\"7\" spans=\"1:5\"><c r=\"A7\" t=\"s\"><v>30</v></c><c r=\"B7\" t=\"s\"><v>31</v></c><c r=\"C7\" t=\"s\"><v>32</v></c><c r=\"D7\" t=\"s\"><v>33</v></c><c r=\"E7\" t=\"s\"><v>34</v></c></row>",
			"<row r=\"9\" spans=\"1:5\"><c r=\"A9\"><v>5</v></c><c r=\"B9\"><v>4</v></c><c r=\"C9\"><v>3</v></c><c r=\"D9\"><v>2</v></c><c r=\"E9\"><v>1</v></c></row>",
		]);

		expect(sheet2RowsData.rows).toEqual([
			"<row r=\"1\" spans=\"1:5\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"D1\"><v>4</v></c><c r=\"E1\"><v>5</v></c></row>",
			"<row r=\"2\" spans=\"1:5\"><c r=\"A2\" t=\"s\"><v>8</v></c><c r=\"B2\" t=\"s\"><v>7</v></c><c r=\"C2\" t=\"s\"><v>6</v></c><c r=\"D2\" t=\"s\"><v>5</v></c><c r=\"E2\" t=\"s\"><v>4</v></c></row>",
			"<row r=\"3\" spans=\"1:5\"><c r=\"A3\" t=\"s\"><v>35</v></c><c r=\"B3\" t=\"s\"><v>36</v></c><c r=\"C3\" t=\"s\"><v>37</v></c><c r=\"D3\" t=\"s\"><v>13</v></c><c r=\"E3\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"4\" spans=\"1:5\"><c r=\"A4\" t=\"s\"><v>39</v></c><c r=\"B4\" t=\"s\"><v>40</v></c><c r=\"C4\" t=\"s\"><v>41</v></c><c r=\"D4\" t=\"s\"><v>18</v></c><c r=\"E4\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"5\" spans=\"1:5\"><c r=\"A5\" t=\"s\"><v>42</v></c><c r=\"B5\" t=\"s\"><v>43</v></c><c r=\"C5\" t=\"s\"><v>12</v></c><c r=\"D5\" t=\"s\"><v>23</v></c><c r=\"E5\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"6\" spans=\"1:5\"><c r=\"A6\" t=\"s\"><v>44</v></c><c r=\"B6\" t=\"s\"><v>45</v></c><c r=\"C6\" t=\"s\"><v>46</v></c><c r=\"D6\" t=\"s\"><v>28</v></c><c r=\"E6\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"7\" spans=\"1:5\"><c r=\"A7\" t=\"s\"><v>47</v></c><c r=\"B7\" t=\"s\"><v>48</v></c><c r=\"C7\" t=\"s\"><v>49</v></c><c r=\"D7\" t=\"s\"><v>33</v></c><c r=\"E7\" t=\"s\"><v>38</v></c></row>",
			"<row r=\"9\" spans=\"1:5\"><c r=\"A9\"><v>5</v></c><c r=\"B9\"><v>4</v></c><c r=\"C9\"><v>3</v></c><c r=\"D9\"><v>2</v></c><c r=\"E9\"><v>1</v></c></row>",
		]);

		const validationResult1 = validateWorksheetXml(sheet1Rebuilt);
		const validationResult2 = validateWorksheetXml(sheet2Rebuilt);

		expect(validationResult1.isValid).toBe(true);
		expect(validationResult2.isValid).toBe(true);
	});
});
