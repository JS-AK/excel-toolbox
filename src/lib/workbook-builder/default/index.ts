// Borders — basic empty border
export const borders = () => [
	{
		children: [
			{ tag: "left" },
			{ tag: "right" },
			{ tag: "top" },
			{ tag: "bottom" },
		],
		tag: "border",
	},
];

export const cellXfs = () => [
	{
		borderId: 0,
		fillId: 0,
		fontId: 0,
		numFmtId: 0,
	},
];

// Заливки — первый элемент пустой patternFill
export const fills = () => [
	{
		children: [
			{
				attrs: { patternType: "none" },
				tag: "patternFill",
			},
		],
		tag: "fill",
	},
];

// Шрифты — пустой базовый шрифт
export const fonts = () => [
	{
		children: [
			{ attrs: { val: "11" }, tag: "sz" },
			{ attrs: { theme: "1" }, tag: "color" },
			{ attrs: { val: "Calibri" }, tag: "name" },
		],
		tag: "font",
	},
];

// Форматы чисел
export const numFmts = () => [];

export const sheetName = () => "Sheet1";
