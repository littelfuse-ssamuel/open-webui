<script lang="ts">
	import { onMount, onDestroy, createEventDispatcher } from 'svelte';

	const dispatch = createEventDispatcher();

	export let content: string;

	// State
	let containerElement: HTMLDivElement;
	let iframeElement: HTMLIFrameElement;
	let loading = true;
	let error: string | null = null;
	let isFullscreen = false;

	// Build the full HTML document with Reveal.js CDN resources
	function buildPresentationHTML(htmlContent: string): string {
		// Check if the content already has full HTML structure
		const hasDoctype = htmlContent.trim().toLowerCase().startsWith('<!doctype');
		const hasHtmlTag = /<html[\s>]/i.test(htmlContent);
		
		if (hasDoctype || hasHtmlTag) {
			// Content already has HTML structure, inject Reveal.js if not present
			if (!htmlContent.includes('reveal.js') && !htmlContent.includes('Reveal.initialize')) {
				// Inject Reveal.js CSS before </head>
				const revealCSS = `
					<link rel="stylesheet" href="https://unpkg.com/reveal.js@5.1.0/dist/reveal.css">
					<link rel="stylesheet" href="https://unpkg.com/reveal.js@5.1.0/dist/theme/black.css">
				`;
				htmlContent = htmlContent.replace('</head>', `${revealCSS}</head>`);
				
				// Inject Reveal.js script and initialization before </body>
				const revealJS = `
					<script src="https://unpkg.com/reveal.js@5.1.0/dist/reveal.js"><\/script>
					<script>
						Reveal.initialize({
							hash: false,
							controls: true,
							progress: true,
							center: true,
							transition: 'slide'
						});
					<\/script>
				`;
				htmlContent = htmlContent.replace('</body>', `${revealJS}</body>`);
			}
			return htmlContent;
		}

		// Content is just the reveal div structure, wrap it in full HTML
		return `<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>Presentation</title>
	<link rel="stylesheet" href="https://unpkg.com/reveal.js@5.1.0/dist/reveal.css">
	<link rel="stylesheet" href="https://unpkg.com/reveal.js@5.1.0/dist/theme/black.css">
	<style>
		body {
			margin: 0;
			padding: 0;
			overflow: hidden;
		}
		.reveal {
			height: 100vh;
		}
	</style>
</head>
<body>
	${htmlContent}
	<script src="https://unpkg.com/reveal.js@5.1.0/dist/reveal.js"><\/script>
	<script>
		Reveal.initialize({
			hash: false,
			controls: true,
			progress: true,
			center: true,
			transition: 'slide'
		});
	<\/script>
</body>
</html>`;
	}

	function handleIframeLoad() {
		loading = false;
		
		// Add click handler to prevent external navigation
		try {
			iframeElement.contentWindow?.addEventListener('click', function(e: MouseEvent) {
				const target = (e.target as HTMLElement)?.closest('a');
				if (target && (target as HTMLAnchorElement).href) {
					const url = new URL((target as HTMLAnchorElement).href, iframeElement.baseURI);
					if (url.origin !== window.location.origin) {
						e.preventDefault();
						console.info('External navigation blocked:', url.href);
					}
				}
			}, true);
		} catch (e) {
			// Cross-origin restrictions may prevent this
			console.warn('Could not attach click handler to iframe:', e);
		}
	}

	function handleIframeError() {
		loading = false;
		error = 'Failed to load presentation';
	}

	// Toggle fullscreen
	export function toggleFullscreen() {
		const wrapper = containerElement?.closest('.presentation-viewer-wrapper');
		if (!wrapper) return;

		if (!document.fullscreenElement) {
			wrapper.requestFullscreen?.();
			isFullscreen = true;
		} else {
			document.exitFullscreen?.();
			isFullscreen = false;
		}
	}

	// Handle fullscreen change
	function handleFullscreenChange() {
		isFullscreen = !!document.fullscreenElement;
	}

	onMount(() => {
		document.addEventListener('fullscreenchange', handleFullscreenChange);
	});

	onDestroy(() => {
		document.removeEventListener('fullscreenchange', handleFullscreenChange);
	});

	// Reactive: rebuild presentation when content changes
	$: presentationHTML = content ? buildPresentationHTML(content) : '';
</script>

<div class="presentation-viewer-wrapper" bind:this={containerElement}>
	<div class="presentation-viewer">
		{#if loading}
			<div class="presentation-loading">
				<div class="presentation-loading-spinner"></div>
				<span>Loading presentation...</span>
			</div>
		{/if}

		{#if error}
			<div class="presentation-error">
				<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
					<circle cx="12" cy="12" r="10"/>
					<line x1="12" y1="8" x2="12" y2="12"/>
					<line x1="12" y1="16" x2="12.01" y2="16"/>
				</svg>
				<span>{error}</span>
			</div>
		{/if}

		<iframe
			bind:this={iframeElement}
			title="Presentation"
			srcdoc={presentationHTML}
			class="presentation-iframe"
			style:visibility={loading || error ? 'hidden' : 'visible'}
			sandbox="allow-scripts allow-same-origin"
			on:load={handleIframeLoad}
			on:error={handleIframeError}
		></iframe>
	</div>
</div>

<style>
	.presentation-viewer-wrapper {
		width: 100%;
		height: 100%;
		display: flex;
		flex-direction: column;
		background: #1a1a1a;
	}

	.presentation-viewer-wrapper:fullscreen {
		background: #000;
	}

	.presentation-viewer {
		display: flex;
		flex-direction: column;
		height: 100%;
		width: 100%;
		overflow: hidden;
		position: relative;
	}

	.presentation-iframe {
		flex: 1;
		width: 100%;
		height: 100%;
		border: none;
		background: #000;
	}

	.presentation-loading,
	.presentation-error {
		position: absolute;
		top: 0;
		left: 0;
		right: 0;
		bottom: 0;
		display: flex;
		flex-direction: column;
		align-items: center;
		justify-content: center;
		gap: 12px;
		color: #9ca3af;
		background: #1a1a1a;
		z-index: 10;
	}

	.presentation-loading-spinner {
		width: 32px;
		height: 32px;
		border: 3px solid #374151;
		border-top-color: #3b82f6;
		border-radius: 50%;
		animation: spin 1s linear infinite;
	}

	@keyframes spin {
		to {
			transform: rotate(360deg);
		}
	}

	.presentation-error {
		color: #f87171;
	}

	.presentation-error svg {
		color: #f87171;
	}

	/* Dark mode support */
	:global(.dark) .presentation-viewer-wrapper {
		background: #1a1a1a;
	}

	:global(.dark) .presentation-loading,
	:global(.dark) .presentation-error {
		background: #1a1a1a;
	}
</style>