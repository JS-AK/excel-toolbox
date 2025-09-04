// Borders — basic empty border
export const border = () => ({
	children: [
		{ tag: "left" },
		{ tag: "right" },
		{ tag: "top" },
		{ tag: "bottom" },
	],
	tag: "border",
});

export const cellXf = () => ({
	borderId: 0,
	fillId: 0,
	fontId: 0,
	numFmtId: 0,
});

// Заливки — первый элемент пустой patternFill
export const fill = () => ({
	children: [
		{
			attrs: { patternType: "none" },
			tag: "patternFill",
		},
	],
	tag: "fill",
});

// Шрифты — пустой базовый шрифт
export const font = () => ({
	children: [
		{ attrs: { val: "11" }, tag: "sz" },
		{ attrs: { theme: "1" }, tag: "color" },
		{ attrs: { val: "Calibri" }, tag: "name" },
		{ attrs: { val: "2" }, tag: "family" },
		{ attrs: { val: "minor" }, tag: "scheme" },
	],
	tag: "font",
});

export const sheetName = () => "Sheet1";
