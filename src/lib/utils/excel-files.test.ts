import { describe, expect, it } from 'vitest';
import {
	buildToolContextFiles,
	getFileIdentity,
	normalizeExcelFileReference
} from './excel-files';

describe('excel-files utils', () => {
	it('normalizes excel file refs to include both id and fileId', () => {
		const normalized = normalizeExcelFileReference({
			type: 'excel',
			name: 'report.xlsx',
			url: '/api/v1/files/file-123/content',
			fileId: 'file-123'
		});

		expect(normalized.id).toBe('file-123');
		expect(normalized.fileId).toBe('file-123');
		expect(getFileIdentity(normalized)).toBe('file-123');
	});

	it('includes assistant excel artifacts in tool context files', () => {
		const messages = [
			{
				role: 'assistant',
				files: [
					{
						type: 'excel',
						name: 'sales.xlsx',
						url: '/api/v1/files/excel-1/content',
						fileId: 'excel-1'
					}
				]
			}
		];

		const files = buildToolContextFiles([], messages, []);
		const excelRef = files.find((file: any) => file.type === 'excel');

		expect(excelRef).toBeTruthy();
		expect(excelRef.id).toBe('excel-1');
		expect(excelRef.fileId).toBe('excel-1');
	});
});
