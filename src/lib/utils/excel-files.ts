export const normalizeExcelFileReference = (file: any) => {
	if (!file || typeof file !== 'object' || file.type !== 'excel') {
		return file;
	}

	const normalizedId = file.id ?? file.fileId;
	if (!normalizedId) {
		return file;
	}

	return {
		...file,
		id: normalizedId,
		fileId: file.fileId ?? normalizedId
	};
};

export const getFileIdentity = (file: any): string | null => {
	if (!file || typeof file !== 'object') {
		return null;
	}
	return file.id ?? file.fileId ?? null;
};

export const buildToolContextFiles = (
	chatFiles: any[],
	messages: any[],
	userMessageFiles: any[]
) => {
	const normalizedChatMessageFiles = messages
		.filter((message) => message.files)
		.flatMap((message) => message.files)
		.map((file) => normalizeExcelFileReference(file));

	const retainedChatFiles = chatFiles
		.map((file) => normalizeExcelFileReference(file))
		.filter((file) => {
			const fileId = getFileIdentity(file);
			return fileId
				? normalizedChatMessageFiles.some(
						(messageFile) => getFileIdentity(messageFile) === fileId
					)
				: false;
		});

	const assistantExcelFiles = messages
		.filter((message) => message.role === 'assistant' && message.files)
		.flatMap((message) => message.files)
		.map((file) => normalizeExcelFileReference(file))
		.filter((file) => file?.type === 'excel' && getFileIdentity(file));

	let files = JSON.parse(JSON.stringify(retainedChatFiles));
	files.push(...assistantExcelFiles);
	files.push(
		...(userMessageFiles ?? []).filter(
			(item) =>
				['doc', 'text', 'note', 'chat', 'collection'].includes(item.type) ||
				(item.type === 'file' && !(item?.content_type ?? '').startsWith('image/'))
		)
	);

	return files.filter(
		(item, index, array) =>
			array.findIndex((i) => JSON.stringify(i) === JSON.stringify(item)) === index
	);
};
