export const removeFromUsageMap = <T>(
	map: Map<string, Map<string, number>>,
	pageName: string,
	key?: T,
): boolean => {
	if (key === undefined || key === null) {
		return false;
	}

	const k = JSON.stringify(key);
	const pageCounts = map.get(k);

	if (!pageCounts) {
		return false;
	}

	const currentCount = pageCounts.get(pageName) ?? 0;
	if (currentCount <= 1) {
		pageCounts.delete(pageName);
	} else {
		pageCounts.set(pageName, currentCount - 1);
	}

	if (pageCounts.size === 0) {
		map.delete(k);

		return true; // стиль больше нигде не используется
	}

	return false; // ещё есть страницы, где используется
};
