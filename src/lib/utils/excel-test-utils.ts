/**
 * Excel Artifact Test Utilities
 * 
 * Phase 1 - Stabilize & Baseline:
 * Helper functions for testing Excel artifact flows.
 */

import type { ExcelArtifact, ExcelFilesEvent } from '$lib/types/excel';

/**
 * Create a mock Excel artifact for testing
 */
export function createMockExcelArtifact(overrides?: Partial<ExcelArtifact>): ExcelArtifact {
	return {
		type: 'excel',
		url: '/api/v1/files/test-file-id/content',
		name: 'test-file.xlsx',
		fileId: 'test-file-id',
		meta: {
			sheetNames: ['Sheet1'],
			activeSheet: 'Sheet1'
		},
		...overrides
	};
}

/**
 * Create a mock files event for testing
 */
export function createMockFilesEvent(files: ExcelArtifact[]): ExcelFilesEvent {
	return {
		type: 'files',
		data: { files }
	};
}

/**
 * Validate event contract compliance
 * Returns array of violations, empty if compliant
 */
export function validateExcelEventContract(event: any): string[] {
	const violations: string[] = [];

	if (event?.type !== 'files') {
		violations.push(`Event type must be 'files', got '${event?.type}'`);
	}

	if (!event?.data?.files || !Array.isArray(event.data.files)) {
		violations.push('Event must have data.files array');
		return violations;
	}

	for (const file of event.data.files) {
		if (file.type === 'excel') {
			if (typeof file.url !== 'string') {
				violations.push('Excel artifact must have string url');
			}
			if (typeof file.name !== 'string') {
				violations.push('Excel artifact must have string name');
			}
			if (!file.url?.includes('/api/v1/files/')) {
				violations.push('Excel URL must follow /api/v1/files/{id}/content pattern');
			}
		}
	}

	return violations;
}

/**
 * Snapshot test helper - captures event structure for comparison
 */
export function captureEventSnapshot(event: ExcelFilesEvent): object {
	return {
		type: event.type,
		fileCount: event.data.files.length,
		fileTypes: event.data.files.map((f) => f.type),
		hasUrls: event.data.files.every((f) => !!f.url),
		hasNames: event.data.files.every((f) => !!f.name),
		hasMeta: event.data.files.every((f) => !!f.meta)
	};
}
