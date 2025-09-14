import { CellStyle } from "./cell-style.js";
import { CellType } from "./cell-type.js";
import { CellValue } from "./cell-value.js";

/** Cell representation. */
export interface CellData {
	value: CellValue;

	isFormula?: boolean;

	type?: CellType;

	style?: CellStyle;
}
