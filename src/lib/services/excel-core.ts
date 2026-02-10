/**
 * Excel Core Service
 * 
 * Phase 2 - Unified Excel Core:
 * Provides a single, consistent API for Excel operations across the application.
 * This service is the canonical source for Excel file handling.
 */

import type {
	ExcelArtifact,
	ExcelUpdateRequest,
	ExcelUpdateResponse,
	ExcelDownloadGateResponse
} from '$lib/types/excel';
import { isValidExcelArtifact, EXCEL_MIME_TYPE } from '$lib/types/excel';

/**
 * Excel Core Service singleton
 */
class ExcelCoreService {
	private static instance: ExcelCoreService;

	private constructor() {}

	static getInstance(): ExcelCoreService {
		if (!ExcelCoreService.instance) {
			ExcelCoreService.instance = new ExcelCoreService();
		}
		return ExcelCoreService.instance;
	}

	/**
	 * Validate an Excel artifact
	 */
	validateArtifact(artifact: any): { valid: boolean; errors: string[] } {
		const errors: string[] = [];

		if (!artifact) {
			errors.push('Artifact is null or undefined');
			return { valid: false, errors };
		}

		if (artifact.type !== 'excel') {
			errors.push(`Expected type 'excel', got '${artifact.type}'`);
		}

		if (typeof artifact.url !== 'string' || !artifact.url) {
			errors.push('Missing or invalid url');
		}

		if (typeof artifact.name !== 'string' || !artifact.name) {
			errors.push('Missing or invalid name');
		}

		return { valid: errors.length === 0, errors };
	}

	/**
	 * Fetch Excel file from URL and return as ArrayBuffer
	 */
	async fetchExcelFile(url: string): Promise<ArrayBuffer> {
		// Add cache-busting for freshness
		const cacheBustUrl = url + (url.includes('?') ? '&' : '?') + '_t=' + Date.now();

		const response = await fetch(cacheBustUrl);
		if (!response.ok) {
			throw new Error(`Failed to fetch Excel file: ${response.statusText}`);
		}

		const contentType = response.headers.get('content-type') || '';
		if (!contentType.includes('spreadsheet') && !contentType.includes('excel')) {
			console.warn('Response content-type is not Excel:', contentType);
		}

		return response.arrayBuffer();
	}

	/**
	 * Save Excel changes to server
	 */
	async saveChanges(request: ExcelUpdateRequest): Promise<ExcelUpdateResponse> {
		const response = await fetch('/api/v1/excel/update', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(request)
		});

		if (!response.ok) {
			const errorData = await response.json().catch(() => ({}));
			throw new Error(errorData.detail || 'Failed to save Excel changes');
		}

		return response.json();
	}


	async checkDownloadReady(params: {
		fileId: string;
		strictMode?: boolean;
		allowLlmRepair?: boolean;
		llmModelId?: string;
		valveLlmModelId?: string;
		fallbackModelId?: string;
	}): Promise<ExcelDownloadGateResponse> {
		const response = await fetch('/api/v1/excel/download-ready', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(params)
		});

		if (!response.ok) {
			const errorData = await response.json().catch(() => ({}));
			throw new Error(errorData.detail || 'Failed to run download QC gate');
		}

		return response.json();
	}

	/**
	 * Download Excel file from blob
	 */
	async downloadExcel(blob: Blob, filename: string): Promise<void> {
		const { saveAs } = await import('file-saver');
		saveAs(blob, filename);
	}

	/**
	 * Check if a file type is Excel
	 */
	isExcelFile(file: { type?: string; name?: string }): boolean {
		const name = file.name?.toLowerCase() || '';
		const type = file.type || '';

		return (
			name.endsWith('.xlsx') ||
			name.endsWith('.xlsm') ||
			name.endsWith('.xls') ||
			type === EXCEL_MIME_TYPE ||
			type.includes('spreadsheet')
		);
	}

	/**
	 * Get file extension from filename
	 */
	getExtension(filename: string): string {
		const match = filename.match(/\.([^.]+)$/);
		return match ? match[1].toLowerCase() : '';
	}
}

// Export singleton instance
export const excelCore = ExcelCoreService.getInstance();

// Export for testing
export { ExcelCoreService };
