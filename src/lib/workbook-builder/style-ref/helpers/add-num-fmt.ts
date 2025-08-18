export const addNumFmt = (payload: {
	formatCode: string;
	numFmts: { formatCode: string; id: number }[];
}) => {
	const { formatCode, numFmts } = payload;

	// 164+ зарезервировано для кастомных форматов
	const existing = numFmts.find(nf => nf.formatCode === formatCode);

	if (existing) return existing.id;

	const id = 164 + numFmts.length;

	numFmts.push({ formatCode, id });

	return id;
};
