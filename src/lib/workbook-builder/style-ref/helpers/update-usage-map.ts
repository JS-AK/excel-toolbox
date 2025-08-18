export const updateUsageMap = <T>(
	map: Map<string, Map<string, number>>,
	pageName: string,
	key?: T,
) => {
	if (key === undefined || key === null) return;

	const k = JSON.stringify(key);

	if (!map.has(k)) {
		map.set(k, new Map<string, number>());
	}

	const pageCounts = map.get(k)!;
	pageCounts.set(pageName, (pageCounts.get(pageName) ?? 0) + 1);
};
