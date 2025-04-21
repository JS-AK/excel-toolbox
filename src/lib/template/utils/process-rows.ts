/**
 * Processes a sheet XML by replacing table placeholders with real data and adjusting row numbers accordingly.
 *
 * @param data - An object containing the following properties:
 *   - `replacements`: An object where keys are table names and values are arrays of objects with table data.
 *   - `sharedIndexMap`: A Map of shared string indexes by their text content.
 *   - `mergeCellMatches`: An array of objects with `from` and `to` properties, describing the merge cells.
 *   - `sharedStrings`: An array of shared strings.
 *   - `sheetMergeCells`: An array of merge cell XML strings.
 *   - `sheetXml`: The sheet XML string.
 *
 * @returns An object with the following properties:
 *   - `lastIndex`: The last processed position in the sheet XML.
 *   - `resultRows`: An array of processed XML rows.
 *   - `rowShift`: The total row shift.
 */
export function processRows(data: {
	replacements: Record<string, unknown>;
	sharedIndexMap: Map<string, number>;
	mergeCellMatches: { from: string; to: string }[];
	sharedStrings: string[];
	sheetMergeCells: string[];
	sheetXml: string;
}) {
	const {
		mergeCellMatches,
		replacements,
		sharedIndexMap,
		sharedStrings,
		sheetMergeCells,
		sheetXml,
	} = data;
	const TABLE_REGEX = /\$\{table:([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/g;

	// Array for storing resulting XML rows
	const resultRows: string[] = [];

	// Previous position of processed part of XML
	let lastIndex = 0;

	// Shift for row numbers
	let rowShift = 0;

	// Regular expression for finding <row> elements
	const rowRegex = /<row[^>]*?>[\s\S]*?<\/row>/g;

	// Process each <row> element
	for (const match of sheetXml.matchAll(rowRegex)) {
		// Full XML row
		const fullRow = match[0];

		// Start position of the row in XML
		const matchStart = match.index!;

		// End position of the row in XML
		const matchEnd = matchStart + fullRow.length;

		// Add the intermediate XML chunk (if any) between the previous and the current row
		if (lastIndex !== matchStart) {
			resultRows.push(sheetXml.slice(lastIndex, matchStart));
		}

		lastIndex = matchEnd;

		// Get row number from r attribute
		const originalRowNumber = parseInt(fullRow.match(/<row[^>]* r="(\d+)"/)?.[1] ?? "1", 10);

		// Update row number based on rowShift
		const shiftedRowNumber = originalRowNumber + rowShift;

		// Find shared string indexes in cells of the current row
		const sharedValueIndexes: number[] = [];

		// Regular expression for finding a cell
		const cellRegex = /<c[^>]*?r="([A-Z]+\d+)"[^>]*?>([\s\S]*?)<\/c>/g;

		for (const cell of fullRow.matchAll(cellRegex)) {
			const cellTag = cell[0];
			// Check if the cell is a shared string
			const isShared = /t="s"/.test(cellTag);
			const valueMatch = cellTag.match(/<v>(\d+)<\/v>/);

			if (isShared && valueMatch) {
				sharedValueIndexes.push(parseInt(valueMatch[1]!, 10));
			}
		}

		// Get the text content of shared strings by their indexes
		const sharedTexts = sharedValueIndexes.map(i => sharedStrings[i]?.replace(/<\/?si>/g, "") ?? "");

		// Find table placeholders in shared strings
		const tablePlaceholders = sharedTexts.flatMap(e => [...e.matchAll(TABLE_REGEX)]);

		// If there are no table placeholders, just shift the row
		if (tablePlaceholders.length === 0) {
			const updatedRow = fullRow
				.replace(/(<row[^>]* r=")(\d+)(")/, `$1${shiftedRowNumber}$3`)
				.replace(/<c r="([A-Z]+)(\d+)"/g, (_, col) => `<c r="${col}${shiftedRowNumber}"`);

			resultRows.push(updatedRow);

			// Update mergeCells for regular row with rowShift
			const calculatedRowNumber = originalRowNumber + rowShift;

			for (const { from, to } of mergeCellMatches) {
				const [, fromCol, fromRow] = from.match(/^([A-Z]+)(\d+)$/)!;
				const [, toCol] = to.match(/^([A-Z]+)(\d+)$/)!;

				if (Number(fromRow) === calculatedRowNumber) {
					const newFrom = `${fromCol}${shiftedRowNumber}`;
					const newTo = `${toCol}${shiftedRowNumber}`;

					sheetMergeCells.push(`<mergeCell ref="${newFrom}:${newTo}"/>`);
				}
			}

			continue;
		}

		// Get the table name from the first placeholder
		const firstMatch = tablePlaceholders[0];
		const tableName = firstMatch?.[1];
		if (!tableName) throw new Error("Table name not found");

		// Get data for replacement from replacements
		const array = replacements[tableName];
		if (!array) continue;
		if (!Array.isArray(array)) throw new Error("Table data is not an array");

		const tableRowStart = shiftedRowNumber;

		// Find mergeCells to duplicate (mergeCells that start with the current row)
		const mergeCellsToDuplicate = mergeCellMatches.filter(({ from }) => {
			const match = from.match(/^([A-Z]+)(\d+)$/);

			if (!match) return false;

			// Row number of the merge cell start position is in the second group
			const rowNumber = match[2];

			return Number(rowNumber) === tableRowStart;
		});

		// Change the current row to multiple rows from the data array
		for (let i = 0; i < array.length; i++) {
			const rowData = array[i];
			let newRow = fullRow;

			// Replace placeholders in shared strings with real data
			sharedValueIndexes.forEach((originalIdx, idx) => {
				const originalText = sharedTexts[idx];
				if (!originalText) throw new Error("Shared value not found");

				// Replace placeholders ${tableName.field} with real data from array data
				const replacedText = originalText.replace(TABLE_REGEX, (_, tbl, field) =>
					tbl === tableName ? String(rowData?.[field] ?? "") : "",
				);

				// Add new text to shared strings if it doesn't exist
				let newIndex: number;

				if (sharedIndexMap.has(replacedText)) {
					newIndex = sharedIndexMap.get(replacedText)!;
				} else {
					newIndex = sharedStrings.length;
					sharedIndexMap.set(replacedText, newIndex);
					sharedStrings.push(`<si>${replacedText}</si>`);
				}

				// Replace the shared string index in the cell
				newRow = newRow.replace(`<v>${originalIdx}</v>`, `<v>${newIndex}</v>`);
			});

			// Update row number and cell references
			const newRowNum = shiftedRowNumber + i;
			newRow = newRow
				.replace(/<row[^>]* r="\d+"/, rowTag => rowTag.replace(/r="\d+"/, `r="${newRowNum}"`))
				.replace(/<c r="([A-Z]+)\d+"/g, (_, col) => `<c r="${col}${newRowNum}"`);

			resultRows.push(newRow);

			// Add duplicate mergeCells for new rows
			for (const { from, to } of mergeCellsToDuplicate) {
				const [, colFrom, rowFrom] = from.match(/^([A-Z]+)(\d+)$/)!;
				const [, colTo, rowTo] = to.match(/^([A-Z]+)(\d+)$/)!;
				const newFrom = `${colFrom}${Number(rowFrom) + i}`;
				const newTo = `${colTo}${Number(rowTo) + i}`;

				sheetMergeCells.push(`<mergeCell ref="${newFrom}:${newTo}"/>`);
			}
		}

		// It increases the row shift by the number of added rows minus one replaced
		rowShift += array.length - 1;

		const delta = array.length - 1;

		const calculatedRowNumber = originalRowNumber + rowShift - array.length + 1;

		if (delta > 0) {
			for (const merge of mergeCellMatches) {
				const fromRow = parseInt(merge.from.match(/\d+$/)![0], 10);
				if (fromRow > calculatedRowNumber) {
					merge.from = merge.from.replace(/\d+$/, r => `${parseInt(r) + delta}`);
					merge.to = merge.to.replace(/\d+$/, r => `${parseInt(r) + delta}`);
				}
			}
		}
	}

	return { lastIndex, resultRows, rowShift };
};
