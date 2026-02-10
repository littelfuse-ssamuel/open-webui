/**
 * Excel Artifact Types for Open WebUI
 * 
 * EVENT CONTRACT (FROZEN - DO NOT MODIFY WITHOUT MIGRATION PLAN):
 * 
 * Event Type: "files"
 * Event Structure:
 * {
 *   type: "files",
 *   data: {
 *     files: ExcelArtifact[]
 *   }
 * }
 * 
 * Required fields for Excel artifacts:
 * - type: "excel" (literal)
 * - url: string (download URL from /api/v1/files/{id}/content)
 * - name: string (filename with .xlsx extension)
 * 
 * Optional fields:
 * - fileId: string (UUID for save operations)
 * - meta.sheetNames: string[] (list of sheet names)
 * - meta.activeSheet: string (default sheet to display)
 */

/** Excel artifact metadata */
export interface ExcelArtifactMeta {
	/** List of sheet names in the workbook */
	sheetNames?: string[];
	/** Name of the active/default sheet to display */
	activeSheet?: string;
	/** Content type (should always be Excel MIME type) */
	content_type?: string;
	/** File size in bytes */
	size?: number;
}

/** Excel artifact structure (matches backend emission) */
export interface ExcelArtifact {
	/** Artifact type - must be "excel" */
	type: 'excel';
	/** Download URL for the Excel file */
	url: string;
	/** Filename with extension */
	name: string;
	/** File ID for save/update operations */
	fileId?: string;
	/** Excel-specific metadata */
	meta?: ExcelArtifactMeta;
}

/** Backend Excel file event structure */
export interface ExcelFilesEvent {
	type: 'files';
	data: {
		files: ExcelArtifact[];
	};
}

/** Excel update request payload */
export interface ExcelUpdateRequest {
	fileId: string;
	sheet: string;
	changes: ExcelCellChange[];
}

/** Individual cell change for save operations */
export interface ExcelCellChange {
	row: number;
	col: number;
	value: any;
	isFormula: boolean;
}

export interface ExcelQcIssue {
	sheet: string;
	cell: string;
	severity: 'critical' | 'warning';
	issueType: string;
	message: string;
	originalFormula?: string;
	repairedFormula?: string;
}

export interface ExcelQcReport {
	blocked: boolean;
	blockReason: string;
	criticalUnresolved: number;
	issues: ExcelQcIssue[];
	recommendedActions: string[];
}

/** Excel update response */
export interface ExcelUpdateResponse {
	status: 'ok' | 'blocked' | 'error';
	message?: string;
	qcReport?: ExcelQcReport;
}

export interface ExcelDownloadGateResponse {
	status: 'ok' | 'blocked' | 'error';
	downloadUrl?: string;
	qcReport?: ExcelQcReport;
	selectedLlmModelId?: string;
	selectedLlmModelSource?: 'request' | 'valve' | 'fallback';
}

/** Supported Excel file extensions */
export const EXCEL_EXTENSIONS = ['.xlsx', '.xlsm', '.xls'] as const;
export type ExcelExtension = (typeof EXCEL_EXTENSIONS)[number];

/** Excel MIME type constant */
export const EXCEL_MIME_TYPE =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Validate if an artifact is a valid Excel artifact */
export function isValidExcelArtifact(artifact: any): artifact is ExcelArtifact {
	return (
		artifact &&
		artifact.type === 'excel' &&
		typeof artifact.url === 'string' &&
		typeof artifact.name === 'string'
	);
}

/** Validate if a file event contains Excel artifacts */
export function hasExcelArtifacts(event: any): boolean {
	return (
		event?.type === 'files' &&
		Array.isArray(event?.data?.files) &&
		event.data.files.some((f: any) => f?.type === 'excel')
	);
}
