import { describe, expect, it } from "vitest";
import { relationshipMatch, sheetMatch } from "./regexp";

describe("regexp utilities", () => {
	describe("relationshipMatch", () => {
		it("should match a relationship element with the given ID and capture the Target", () => {
			const regex = relationshipMatch("rId1");
			const xml = "<Relationship Id=\"rId1\" Target=\"worksheets/sheet1.xml\"/>";
			const match = xml.match(regex);

			expect(match).not.toBeNull();
			expect(match![1]).toBe("worksheets/sheet1.xml");
		});

		it("should not match a relationship element with a different ID", () => {
			const regex = relationshipMatch("rId1");
			const xml = "<Relationship Id=\"rId2\" Target=\"worksheets/sheet1.xml\"/>";
			const match = xml.match(regex);

			expect(match).toBeNull();
		});

		it("should handle additional attributes in the relationship element", () => {
			const regex = relationshipMatch("rId1");
			const xml = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>";
			const match = xml.match(regex);

			expect(match).not.toBeNull();
			expect(match![1]).toBe("worksheets/sheet1.xml");
		});
	});

	describe("sheetMatch", () => {
		it("should match a sheet element with the given name and capture the r:id", () => {
			const regex = sheetMatch("Sheet1");
			const xml = "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>";
			const match = xml.match(regex);

			expect(match).not.toBeNull();
			expect(match![1]).toBe("rId1");
		});

		it("should not match a sheet element with a different name", () => {
			const regex = sheetMatch("Sheet1");
			const xml = "<sheet name=\"Sheet2\" sheetId=\"1\" r:id=\"rId1\"/>";
			const match = xml.match(regex);

			expect(match).toBeNull();
		});

		it("should handle additional attributes in the sheet element", () => {
			const regex = sheetMatch("Sheet1");
			const xml = "<sheet name=\"Sheet1\" sheetId=\"1\" state=\"visible\" r:id=\"rId1\"/>";
			const match = xml.match(regex);

			expect(match).not.toBeNull();
			expect(match![1]).toBe("rId1");
		});
	});
});
