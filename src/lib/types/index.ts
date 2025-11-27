export type Banner = {
	id: string;
	type: string;
	title?: string;
	content: string;
	url?: string;
	dismissible?: boolean;
	timestamp: number;
};

export enum TTS_RESPONSE_SPLIT {
	PUNCTUATION = 'punctuation',
	PARAGRAPHS = 'paragraphs',
	NONE = 'none'
}

// Excel artifact types for file handling
export type ExcelArtifact = {
	type: 'excel';
	url: string;
	name: string;
	fileId?: string;
	meta?: {
		sheetNames?: string[];
		activeSheet?: string;
	};
};

// Presentation artifact types for Reveal.js presentations
export type PresentationArtifact = {
	type: 'presentation';
	content: string;
};

export type FileArtifact = {
	type: 'image' | 'audio' | 'file' | 'excel' | 'presentation';
	url: string;
	name?: string;
	fileId?: string;
	meta?: Record<string, any>;
};
