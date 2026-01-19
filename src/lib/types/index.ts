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

// PPTX artifact types for PowerPoint presentations
export type PptxContentItem = {
	type: 'text' | 'bullet' | 'table' | 'image';
	text?: string;
	items?: string[];
	headers?: string[];
	rows?: string[][];
	src?: string;
	alt?: string;
};

export type PptxSlide = {
	title?: string;
	backgroundColor?: string;
	content?: PptxContentItem[];
	notes?: string;
};

export type PptxArtifact = {
	type: 'pptx';
	title: string;
	slides: PptxSlide[];
	fileId?: string;
	url?: string;
};

export type FileArtifact = {
	type: 'image' | 'audio' | 'file' | 'excel' | 'presentation';
	url: string;
	name?: string;
	fileId?: string;
	meta?: Record<string, any>;
};
