import { BorderStyle } from "./border-style.js";

/** Style definition applied to a cell. */
export type CellStyle = {
	index?: number;
	font?: {
		name?: string;           // Font name, e.g. "Calibri"
		size?: number;           // Font size, e.g. 11, 14
		bold?: boolean;          // Bold
		italic?: boolean;        // Italic
		underline?: boolean | "single" | "double"; // Underline
		color?: string;          // Text color in HEX or ARGB, e.g. "#FF0000"
	};
	fill?: {
		type?: "pattern";        // Currently supports patternFill
		patternType?: string;    // e.g. "solid", "gray125", "none"
		fgColor?: string;        // Foreground (fill) color — HEX or ARGB
		bgColor?: string;        // Background color (rarely used)
	};
	border?: {
		top?: BorderStyle;
		bottom?: BorderStyle;
		left?: BorderStyle;
		right?: BorderStyle;
	};
	alignment?: {
		horizontal?: "left" | "center" | "right" | "justify";
		vertical?: "top" | "center" | "bottom";
		wrapText?: boolean;
		indent?: number;
	};
	numberFormat?: string;     // Number format, e.g. "0.00", "dd/mm/yyyy", "$#,##0.00"
};
