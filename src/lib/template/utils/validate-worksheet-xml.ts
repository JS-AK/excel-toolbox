interface ValidationResult {
	isValid: boolean;
	error?: {
		message: string;
		details?: string;
	};
}

/**
 * Validates an Excel worksheet XML against the expected structure and rules.
 *
 * Checks the following:
 * 1. XML starts with <?xml declaration
 * 2. Root element is worksheet
 * 3. Required elements are present
 * 4. row numbers are in ascending order
 * 5. No duplicate row numbers
 * 6. No overlapping merge ranges
 * 7. All cells are within the specified dimension
 * 8. All mergeCell tags refer to existing cells
 *
 * @param xml The raw XML content of the worksheet
 * @returns A ValidationResult object indicating if the XML is valid, and an error message if it's not
 */
export function validateWorksheetXml(xml: string): ValidationResult {
	const createError = (message: string, details?: string): ValidationResult => ({
		error: { details, message },
		isValid: false,
	});

	// 1. Check for XML declaration
	if (!xml.startsWith("<?xml")) {
		return createError("XML must start with <?xml> declaration");
	}

	if (!xml.includes("<worksheet") || !xml.includes("</worksheet>")) {
		return createError("Root element worksheet not found");
	}

	// 2. Check for required elements
	const requiredElements = [
		{ name: "sheetViews", tag: "<sheetViews>" },
		{ name: "sheetFormatPr", tag: "<sheetFormatPr" },
		{ name: "sheetData", tag: "<sheetData" },
	];

	for (const { name, tag } of requiredElements) {
		if (!xml.includes(tag)) {
			return createError(`Missing required element ${name}`);
		}
	}

	// 3. Extract and validate sheetData
	// 3. Extract and validate sheetData
	// Поддержка пустого <sheetData/> или обычного открывающего-закрывающего тега
	const sheetDataMatch = xml.match(/<sheetData([^>]*)>/);
	if (!sheetDataMatch) {
		return createError("Missing sheetData element");
	}

	const isSelfClosing = sheetDataMatch[1]?.includes("/"); // <sheetData/> самозакрывающийся
	let sheetDataContent = "";

	if (!isSelfClosing) {
		const sheetDataStart = sheetDataMatch.index! + sheetDataMatch[0].length;
		const sheetDataEnd = xml.indexOf("</sheetData>", sheetDataStart);
		if (sheetDataEnd === -1) {
			return createError("Invalid sheetData structure: missing closing tag");
		}
		sheetDataContent = xml.substring(sheetDataStart, sheetDataEnd);
	}

	// Разбиваем на строки, если есть содержимое
	const rows = sheetDataContent
		? sheetDataContent.split("</row>").map(r => r.trim()).filter(r => r.length)
		: [];

	// Collect information about all rows and cells
	const allRows: number[] = [];
	const allCells: { row: number; col: string }[] = [];
	let prevRowNum = 0;

	for (const row of rows.slice(0, -1)) {
		if (!row.includes("<row ")) {
			return createError("Row tag not found", `Fragment: ${row.substring(0, 50)}...`);
		}

		if (!row.includes("<c ")) {
			return createError("Row does not contain any cells", `Row: ${row.substring(0, 50)}...`);
		}

		// Extract row number
		const rowNumMatch = row.match(/<row\s+r="(\d+)"/);
		if (!rowNumMatch) {
			return createError("Row number (attribute r) not specified", `Row: ${row.substring(0, 50)}...`);
		}
		const rowNum = parseInt(rowNumMatch[1]!);

		// Check for duplicate row numbers
		if (allRows.includes(rowNum)) {
			return createError("Duplicate row number found", `Row number: ${rowNum}`);
		}
		allRows.push(rowNum);

		// Check row number order (should be in ascending order)
		if (rowNum <= prevRowNum) {
			return createError(
				"Row order is broken",
				`Current row: ${rowNum}, previous: ${prevRowNum}`,
			);
		}
		prevRowNum = rowNum;

		// Extract all cells in the row
		const cells = row.match(/<c\s+r="([A-Z]+)(\d+)"/g) || [];
		for (const cell of cells) {
			const match = cell.match(/<c\s+r="([A-Z]+)(\d+)"/);
			if (!match) {
				return createError("Invalid cell format", `Cell: ${cell}`);
			}

			const col = match[1]!;
			const cellRowNum = parseInt(match[2]!);

			// Check row number match for each cell
			if (cellRowNum !== rowNum) {
				return createError(
					"Row number mismatch in cell",
					`Expected: ${rowNum}, found: ${cellRowNum} in cell ${col}${cellRowNum}`,
				);
			}

			allCells.push({
				col,
				row: rowNum,
			});
		}
	}

	// 4. Check mergeCells
	if (xml.includes("<mergeCells")) {

		const mergeCellsStart = xml.indexOf("<mergeCells");
		const mergeCellsEnd = xml.indexOf("</mergeCells>");
		if (mergeCellsStart === -1 || mergeCellsEnd === -1) {
			return createError("Invalid mergeCells structure");
		}

		const mergeCellsContent = xml.substring(mergeCellsStart, mergeCellsEnd);
		const countMatch = mergeCellsContent.match(/count="(\d+)"/);
		if (!countMatch) {
			return createError("Count attribute not specified for mergeCells");
		}

		const mergeCellTags = mergeCellsContent.match(/<mergeCell\s+ref="([A-Z]+\d+:[A-Z]+\d+)"\s*\/>/g);
		if (!mergeCellTags) {
			return createError("No merged cells found");
		}

		// Check if the number of mergeCells matches the count attribute
		if (mergeCellTags.length !== parseInt(countMatch[1]!)) {
			return createError(
				"Mismatch in the number of merged cells",
				`Expected: ${countMatch[1]}, found: ${mergeCellTags.length}`,
			);
		}

		// Check for duplicates of mergeCell
		const mergeRefs = new Set<string>();
		const duplicates = new Set<string>();

		for (const mergeTag of mergeCellTags) {
			const refMatch = mergeTag.match(/ref="([A-Z]+\d+:[A-Z]+\d+)"/);
			if (!refMatch) {
				return createError("Invalid merge cell format", `Tag: ${mergeTag}`);
			}

			const ref = refMatch[1];
			if (mergeRefs.has(ref!)) {
				duplicates.add(ref!);
			} else {
				mergeRefs.add(ref!);
			}
		}

		if (duplicates.size > 0) {
			return createError(
				"Duplicates of merged cells found",
				`Duplicates: ${Array.from(duplicates).join(", ")}`,
			);
		}

		// Check for overlapping merge ranges
		const mergedRanges = Array.from(mergeRefs).map(ref => {
			const [start, end] = ref.split(":");
			return {
				endCol: end!.match(/[A-Z]+/)?.[0] || "",
				endRow: parseInt(end!.match(/\d+/)?.[0] || "0"),
				startCol: start!.match(/[A-Z]+/)?.[0] || "",
				startRow: parseInt(start!.match(/\d+/)?.[0] || "0"),
			};
		});

		for (let i = 0; i < mergedRanges.length; i++) {
			for (let j = i + 1; j < mergedRanges.length; j++) {
				const a = mergedRanges[i];
				const b = mergedRanges[j];

				if (rangesIntersect(a!, b!)) {
					return createError(
						"Found intersecting merged cells",
						`Intersecting: ${getRangeString(a!)} and ${getRangeString(b!)}`,
					);
				}
			}
		}

		// 6. Additional check: all mergeCell tags refer to existing cells
		for (const mergeTag of mergeCellTags) {
			const refMatch = mergeTag.match(/ref="([A-Z]+\d+:[A-Z]+\d+)"/);
			if (!refMatch) {
				return createError("Invalid merge cell format", `Tag: ${mergeTag}`);
			}

			const [cell1, cell2] = refMatch[1]!.split(":");
			const cell1Col = cell1!.match(/[A-Z]+/)?.[0];
			const cell1Row = parseInt(cell1!.match(/\d+/)?.[0] || "0");
			const cell2Col = cell2!.match(/[A-Z]+/)?.[0];
			const cell2Row = parseInt(cell2!.match(/\d+/)?.[0] || "0");

			if (!cell1Col || !cell2Col || isNaN(cell1Row) || isNaN(cell2Row)) {
				return createError("Invalid merged cell coordinates", `Merged cells: ${refMatch[1]}`);
			}

			// Check if the merged cells exist
			const cell1Exists = allCells.some(c => c.row === cell1Row && c.col === cell1Col);
			const cell2Exists = allCells.some(c => c.row === cell2Row && c.col === cell2Col);

			if (!cell1Exists || !cell2Exists) {
				return createError(
					"Merged cell reference points to non-existent cells",
					`Merged cells: ${refMatch[1]}, missing: ${!cell1Exists ? `${cell1Col}${cell1Row}` : `${cell2Col}${cell2Row}`}`,
				);
			}
		}
	}

	// 5. Check dimension and match with real data
	const dimensionMatch = xml.match(/<dimension\s+ref="([A-Z]+\d+:[A-Z]+\d+)"\s*\/>/);
	if (!dimensionMatch) {
		return createError("Data range (dimension) is not specified");
	}

	const [startCell, endCell] = dimensionMatch[1]!.split(":");
	const startCol = startCell!.match(/[A-Z]+/)?.[0];
	const startRow = parseInt(startCell!.match(/\d+/)?.[0] || "0");
	const endCol = endCell!.match(/[A-Z]+/)?.[0];
	const endRow = parseInt(endCell!.match(/\d+/)?.[0] || "0");

	if (!startCol || !endCol || isNaN(startRow) || isNaN(endRow)) {
		return createError("Invalid dimension format", `Dimension: ${dimensionMatch[1]}`);
	}

	const startColNum = colToNumber(startCol);
	const endColNum = colToNumber(endCol);

	// Check if all cells are within the dimension
	for (const cell of allCells) {
		const colNum = colToNumber(cell.col);

		if (cell.row < startRow || cell.row > endRow) {
			return createError(
				"Cell is outside the specified area (by row)",
				`Cell: ${cell.col}${cell.row}, dimension: ${dimensionMatch[1]}`,
			);
		}

		if (colNum < startColNum || colNum > endColNum) {
			return createError(
				"Cell is outside the specified area (by column)",
				`Cell: ${cell.col}${cell.row}, dimension: ${dimensionMatch[1]}`,
			);
		}
	}

	return { isValid: true };
}

// A function to check if two ranges intersect
function rangesIntersect(a: { startCol: string; startRow: number; endCol: string; endRow: number },
	b: { startCol: string; startRow: number; endCol: string; endRow: number }): boolean {
	const aStartColNum = colToNumber(a.startCol);
	const aEndColNum = colToNumber(a.endCol);
	const bStartColNum = colToNumber(b.startCol);
	const bEndColNum = colToNumber(b.endCol);

	// Check if the rows intersect
	const rowsIntersect = !(a.endRow < b.startRow || a.startRow > b.endRow);

	// Check if the columns intersect
	const colsIntersect = !(aEndColNum < bStartColNum || aStartColNum > bEndColNum);

	return rowsIntersect && colsIntersect;
}

// Function to get the range string1
function getRangeString(range: { startCol: string; startRow: number; endCol: string; endRow: number }): string {
	return `${range.startCol}${range.startRow}:${range.endCol}${range.endRow}`;
}

// Function to convert column letters to numbers
function colToNumber(col: string): number {
	let num = 0;
	for (let i = 0; i < col.length; i++) {
		num = num * 26 + (col.charCodeAt(i) - 64);
	}
	return num;
};
