export type Alignment = {
	horizontal?: "left" | "center" | "right" | "justify";
	vertical?: "top" | "center" | "bottom";
	wrapText?: boolean;
	indent?: number;
};

export type CellXf = {
	borderId: number;
	fillId: number;
	fontId: number;
	numFmtId: number;
	alignment?: Alignment;
};
