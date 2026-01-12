<script lang="ts">
	import { onMount, onDestroy, createEventDispatcher } from 'svelte';
	import { toast } from 'svelte-sonner';

	const dispatch = createEventDispatcher();

	// Props - expecting the parsed JSON slide data
	export let slideData: {
		type: 'pptx';
		title: string;
		slides: Array<{
			title?: string;
			backgroundColor?: string;
			content?: Array<{
				type: 'text' | 'bullet' | 'table' | 'image';
				text?: string;
				items?: string[];
				headers?: string[];
				rows?: string[][];
				src?: string;
				alt?: string;
			}>;
			notes?: string;
		}>;
		// These are set after generation
		fileId?: string;
		url?: string;
	};

	// State
	let currentSlide = 0;
	let fileId: string | null = null;
	let downloadUrl: string | null = null;
	let loading = false;
	let generating = false;
	let error: string | null = null;
	let isFullscreen = false;
	let showNotes = false;
	let containerElement: HTMLDivElement;

	// Navigation
	function nextSlide() {
		if (currentSlide < slideData.slides.length - 1) {
			currentSlide++;
		}
	}

	function prevSlide() {
		if (currentSlide > 0) {
			currentSlide--;
		}
	}

	function goToSlide(index: number) {
		if (index >= 0 && index < slideData.slides.length) {
			currentSlide = index;
		}
	}

	// Keyboard navigation
	function handleKeydown(event: KeyboardEvent) {
		// Don't handle if user is typing in an input
		if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement) {
			return;
		}
		
		if (event.key === 'ArrowRight' || event.key === ' ') {
			event.preventDefault();
			nextSlide();
		} else if (event.key === 'ArrowLeft') {
			event.preventDefault();
			prevSlide();
		} else if (event.key === 'Home') {
			event.preventDefault();
			currentSlide = 0;
		} else if (event.key === 'End') {
			event.preventDefault();
			currentSlide = slideData.slides.length - 1;
		} else if (event.key === 'Escape' && isFullscreen) {
			toggleFullscreen();
		}
	}

	// Generate PPTX file via backend API
	async function generatePptx() {
		if (generating) return;
		
		// If already have file, don't regenerate
		if (fileId && downloadUrl) {
			return;
		}
		
		generating = true;
		error = null;
		
		try {
			const token = localStorage.getItem('token');
			const response = await fetch('/api/v1/pptx/generate', {
				method: 'POST',
				headers: {
					'Content-Type': 'application/json',
					...(token ? { 'Authorization': `Bearer ${token}` } : {})
				},
				body: JSON.stringify({
					title: slideData.title,
					slides: slideData.slides,
					use_template: true
				})
			});
			
			if (!response.ok) {
				const errorData = await response.json().catch(() => ({}));
				throw new Error(errorData.detail || `HTTP ${response.status}`);
			}
			
			const data = await response.json();
			fileId = data.file_id;
			downloadUrl = data.download_url;
			
			// Update slideData with file info for parent components
			slideData.fileId = fileId;
			slideData.url = downloadUrl;
			
			toast.success('PowerPoint file ready for download');
			
		} catch (e) {
			console.error('Error generating PPTX:', e);
			error = e instanceof Error ? e.message : 'Failed to generate PowerPoint file';
			toast.error(error);
		} finally {
			generating = false;
		}
	}

	// Download the generated PPTX
	async function downloadPptx() {
		// Generate if not already done
		if (!downloadUrl) {
			await generatePptx();
		}
		
		if (!downloadUrl) {
			toast.error('Failed to generate file');
			return;
		}
		
		loading = true;
		
		try {
			const token = localStorage.getItem('token');
			const response = await fetch(downloadUrl, {
				headers: {
					...(token ? { 'Authorization': `Bearer ${token}` } : {})
				}
			});
			
			if (!response.ok) {
				throw new Error('Download failed');
			}
			
			const blob = await response.blob();
			const url = URL.createObjectURL(blob);
			
			// Create download link
			const link = document.createElement('a');
			link.href = url;
			link.download = `${slideData.title || 'presentation'}.pptx`;
			document.body.appendChild(link);
			link.click();
			document.body.removeChild(link);
			
			// Cleanup
			URL.revokeObjectURL(url);
			
		} catch (e) {
			console.error('Download error:', e);
			toast.error('Failed to download file');
		} finally {
			loading = false;
		}
	}

	// Fullscreen toggle
	function toggleFullscreen() {
		const wrapper = containerElement?.closest('.pptx-viewer-wrapper');
		if (!wrapper) return;

		if (!document.fullscreenElement) {
			wrapper.requestFullscreen?.();
			isFullscreen = true;
		} else {
			document.exitFullscreen?.();
			isFullscreen = false;
		}
	}

	function handleFullscreenChange() {
		isFullscreen = !!document.fullscreenElement;
	}

	// Lifecycle
	onMount(() => {
		document.addEventListener('fullscreenchange', handleFullscreenChange);
		
		// Check if file already exists
		if (slideData.fileId && slideData.url) {
			fileId = slideData.fileId;
			downloadUrl = slideData.url;
		} else {
			// Auto-generate PPTX on mount for quick download
			generatePptx();
		}
	});

	onDestroy(() => {
		document.removeEventListener('fullscreenchange', handleFullscreenChange);
	});

	// Reactive
	$: slide = slideData?.slides?.[currentSlide];
	$: slideCount = slideData?.slides?.length || 0;
	$: canDownload = !generating && !loading;
</script>

<svelte:window on:keydown={handleKeydown} />

<div class="pptx-viewer-wrapper" bind:this={containerElement}>
	<div class="pptx-viewer">
		<!-- Header -->
		<div class="pptx-header">
			<div class="pptx-header-left">
				<svg class="pptx-icon" xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
					<path d="M2 3h20v18H2z"/>
					<path d="M2 7h20"/>
					<path d="M6 3v4"/>
				</svg>
				<span class="pptx-title">{slideData?.title || 'Presentation'}</span>
				{#if generating}
					<span class="pptx-status generating">Generating...</span>
				{:else if fileId}
					<span class="pptx-status ready">Ready</span>
				{/if}
			</div>
			<div class="pptx-header-right">
				{#if slide?.notes}
					<button 
						class="pptx-btn"
						on:click={() => showNotes = !showNotes}
						title={showNotes ? 'Hide notes' : 'Show notes'}
						class:active={showNotes}
					>
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/>
							<polyline points="14,2 14,8 20,8"/>
							<line x1="16" y1="13" x2="8" y2="13"/>
							<line x1="16" y1="17" x2="8" y2="17"/>
						</svg>
					</button>
				{/if}
				<button
					class="pptx-btn"
					on:click={toggleFullscreen}
					title={isFullscreen ? 'Exit fullscreen' : 'Fullscreen'}
				>
					{#if isFullscreen}
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<polyline points="4 14 10 14 10 20"/>
							<polyline points="20 10 14 10 14 4"/>
							<line x1="14" y1="10" x2="21" y2="3"/>
							<line x1="3" y1="21" x2="10" y2="14"/>
						</svg>
					{:else}
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<polyline points="15 3 21 3 21 9"/>
							<polyline points="9 21 3 21 3 15"/>
							<line x1="21" y1="3" x2="14" y2="10"/>
							<line x1="3" y1="21" x2="10" y2="14"/>
						</svg>
					{/if}
				</button>
				<button
					class="pptx-btn pptx-btn-primary"
					on:click={downloadPptx}
					disabled={!canDownload}
					title="Download PowerPoint"
				>
					{#if generating || loading}
						<svg class="animate-spin" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
							<circle cx="12" cy="12" r="10" stroke-dasharray="32" stroke-dashoffset="12"/>
						</svg>
						<span>{generating ? 'Generating...' : 'Downloading...'}</span>
					{:else}
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
							<polyline points="7 10 12 15 17 10"/>
							<line x1="12" y1="15" x2="12" y2="3"/>
						</svg>
						<span>Download .pptx</span>
					{/if}
				</button>
			</div>
		</div>

		<!-- Main slide area -->
		<div class="pptx-main">
			<div class="pptx-slide-container">
				{#if slide}
					<div 
						class="pptx-slide"
						style:background-color={slide.backgroundColor || '#ffffff'}
					>
						{#if slide.title}
							<h1 class="slide-title">{slide.title}</h1>
						{/if}
						
						{#if slide.content && slide.content.length > 0}
							<div class="slide-content">
								{#each slide.content as item}
									{#if item.type === 'text'}
										<p class="slide-text">{item.text}</p>
									{:else if item.type === 'bullet'}
										<ul class="slide-bullets">
											{#each item.items || [] as bullet}
												<li>{bullet}</li>
											{/each}
										</ul>
									{:else if item.type === 'image' && item.src}
										<div class="slide-image-container">
											<img 
												src={item.src} 
												alt={item.alt || ''} 
												class="slide-image"
											/>
										</div>
									{:else if item.type === 'table' && item.headers}
										<div class="slide-table-container">
											<table class="slide-table">
												<thead>
													<tr>
														{#each item.headers as header}
															<th>{header}</th>
														{/each}
													</tr>
												</thead>
												<tbody>
													{#each item.rows || [] as row}
														<tr>
															{#each row as cell}
																<td>{cell}</td>
															{/each}
														</tr>
													{/each}
												</tbody>
											</table>
										</div>
									{/if}
								{/each}
							</div>
						{:else if !slide.title}
							<div class="slide-empty">
								<p>Empty slide</p>
							</div>
						{/if}
					</div>
				{:else}
					<div class="pptx-empty">
						<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round" stroke-linejoin="round">
							<path d="M2 3h20v18H2z"/>
							<path d="M2 7h20"/>
						</svg>
						<p>No slides available</p>
					</div>
				{/if}
			</div>

			<!-- Speaker notes (collapsible) -->
			{#if showNotes && slide?.notes}
				<div class="pptx-notes">
					<div class="pptx-notes-header">
						<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/>
						</svg>
						<span>Speaker Notes</span>
					</div>
					<div class="pptx-notes-content">{slide.notes}</div>
				</div>
			{/if}
		</div>

		<!-- Navigation controls -->
		<div class="pptx-navigation">
			<button 
				class="pptx-nav-btn"
				on:click={prevSlide}
				disabled={currentSlide === 0}
				title="Previous slide (←)"
			>
				<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
					<polyline points="15 18 9 12 15 6"/>
				</svg>
				<span class="nav-text">Previous</span>
			</button>
			
			<div class="pptx-slide-info">
				<!-- Thumbnail strip -->
				<div class="pptx-thumbnails">
					{#each slideData?.slides || [] as thumbSlide, index}
						<button
							class="pptx-thumbnail"
							class:active={index === currentSlide}
							on:click={() => goToSlide(index)}
							title={thumbSlide.title || `Slide ${index + 1}`}
						>
							<span class="pptx-thumbnail-number">{index + 1}</span>
						</button>
					{/each}
				</div>
				
				<div class="pptx-slide-indicator">
					<span class="pptx-slide-current">{currentSlide + 1}</span>
					<span class="pptx-slide-separator">/</span>
					<span class="pptx-slide-total">{slideCount}</span>
				</div>
			</div>

			<button 
				class="pptx-nav-btn"
				on:click={nextSlide}
				disabled={currentSlide === slideCount - 1}
				title="Next slide (→)"
			>
				<span class="nav-text">Next</span>
				<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
					<polyline points="9 18 15 12 9 6"/>
				</svg>
			</button>
		</div>
	</div>
</div>

<style>
	.pptx-viewer-wrapper {
		width: 100%;
		height: 100%;
		display: flex;
		flex-direction: column;
		background: #1a1a2e;
		font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
		color: #e0e0e0;
	}

	.pptx-viewer-wrapper:fullscreen {
		background: #0d0d1a;
	}

	.pptx-viewer {
		display: flex;
		flex-direction: column;
		height: 100%;
		width: 100%;
		overflow: hidden;
	}

	/* Header */
	.pptx-header {
		display: flex;
		justify-content: space-between;
		align-items: center;
		padding: 10px 16px;
		background: linear-gradient(180deg, #252540 0%, #1e1e35 100%);
		border-bottom: 1px solid #3a3a5c;
		flex-shrink: 0;
	}

	.pptx-header-left {
		display: flex;
		align-items: center;
		gap: 10px;
	}

	.pptx-icon {
		color: #c44e23;
	}

	.pptx-title {
		font-size: 14px;
		font-weight: 600;
		color: #ffffff;
		white-space: nowrap;
		overflow: hidden;
		text-overflow: ellipsis;
		max-width: 280px;
	}

	.pptx-status {
		font-size: 11px;
		padding: 2px 8px;
		border-radius: 10px;
		font-weight: 500;
	}

	.pptx-status.generating {
		background: #3b3b5c;
		color: #a0a0c0;
	}

	.pptx-status.ready {
		background: #1e4620;
		color: #4ade80;
	}

	.pptx-header-right {
		display: flex;
		align-items: center;
		gap: 8px;
	}

	.pptx-btn {
		display: flex;
		align-items: center;
		justify-content: center;
		gap: 6px;
		padding: 7px 12px;
		background: #2a2a45;
		border: 1px solid #3a3a5c;
		border-radius: 6px;
		color: #c0c0d0;
		font-size: 13px;
		cursor: pointer;
		transition: all 0.15s ease;
	}

	.pptx-btn:hover:not(:disabled) {
		background: #35355a;
		border-color: #4a4a7c;
		color: #ffffff;
	}

	.pptx-btn:disabled {
		opacity: 0.5;
		cursor: not-allowed;
	}

	.pptx-btn.active {
		background: #3a3a6c;
		border-color: #5a5a9c;
	}

	.pptx-btn-primary {
		background: #c44e23;
		border-color: #d55a2f;
		color: white;
		font-weight: 500;
	}

	.pptx-btn-primary:hover:not(:disabled) {
		background: #d55a2f;
		border-color: #e66a3f;
	}

	/* Main slide area */
	.pptx-main {
		flex: 1;
		display: flex;
		flex-direction: column;
		padding: 20px;
		overflow: hidden;
		background: linear-gradient(135deg, #1a1a2e 0%, #16162a 100%);
	}

	.pptx-slide-container {
		flex: 1;
		display: flex;
		align-items: center;
		justify-content: center;
		overflow: hidden;
	}

	.pptx-slide {
		width: 100%;
		max-width: 900px;
		aspect-ratio: 16 / 9;
		background: white;
		border-radius: 8px;
		box-shadow: 
			0 4px 6px rgba(0, 0, 0, 0.3),
			0 10px 40px rgba(0, 0, 0, 0.4),
			0 0 0 1px rgba(255, 255, 255, 0.05);
		padding: 32px 40px;
		overflow-y: auto;
		box-sizing: border-box;
		color: #1a1a1a;
	}

	.pptx-empty {
		display: flex;
		flex-direction: column;
		align-items: center;
		gap: 16px;
		color: #6a6a8a;
	}

	.slide-empty {
		display: flex;
		align-items: center;
		justify-content: center;
		height: 100%;
		color: #999;
		font-style: italic;
	}

	/* Slide content styles */
	.slide-title {
		font-size: 28px;
		font-weight: 700;
		color: #1a1a1a;
		margin: 0 0 20px 0;
		padding-bottom: 12px;
		border-bottom: 3px solid #c44e23;
	}

	.slide-content {
		display: flex;
		flex-direction: column;
		gap: 16px;
	}

	.slide-text {
		font-size: 17px;
		line-height: 1.6;
		color: #333;
		margin: 0;
	}

	.slide-bullets {
		font-size: 16px;
		line-height: 1.7;
		color: #333;
		margin: 0;
		padding-left: 24px;
	}

	.slide-bullets li {
		margin: 6px 0;
	}

	.slide-bullets li::marker {
		color: #c44e23;
	}

	.slide-image-container {
		display: flex;
		justify-content: center;
	}

	.slide-image {
		max-width: 100%;
		max-height: 280px;
		object-fit: contain;
		border-radius: 6px;
		box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
	}

	.slide-table-container {
		overflow-x: auto;
	}

	.slide-table {
		width: 100%;
		border-collapse: collapse;
		font-size: 14px;
	}

	.slide-table th,
	.slide-table td {
		border: 1px solid #ddd;
		padding: 10px 12px;
		text-align: left;
	}

	.slide-table th {
		background: #c44e23;
		color: white;
		font-weight: 600;
	}

	.slide-table tr:nth-child(even) {
		background: #f8f8f8;
	}

	/* Speaker notes */
	.pptx-notes {
		margin-top: 16px;
		background: #252540;
		border: 1px solid #3a3a5c;
		border-radius: 8px;
		overflow: hidden;
		max-height: 100px;
	}

	.pptx-notes-header {
		display: flex;
		align-items: center;
		gap: 6px;
		padding: 8px 12px;
		background: #2a2a4a;
		font-size: 11px;
		font-weight: 600;
		color: #8a8aaa;
		text-transform: uppercase;
		letter-spacing: 0.5px;
	}

	.pptx-notes-content {
		padding: 10px 12px;
		font-size: 13px;
		color: #b0b0c0;
		line-height: 1.5;
		overflow-y: auto;
		max-height: 60px;
	}

	/* Navigation */
	.pptx-navigation {
		display: flex;
		align-items: center;
		justify-content: space-between;
		padding: 12px 16px;
		background: linear-gradient(180deg, #1e1e35 0%, #252540 100%);
		border-top: 1px solid #3a3a5c;
		flex-shrink: 0;
	}

	.pptx-nav-btn {
		display: flex;
		align-items: center;
		gap: 6px;
		padding: 8px 14px;
		background: #2a2a45;
		border: 1px solid #3a3a5c;
		border-radius: 6px;
		color: #c0c0d0;
		font-size: 13px;
		font-weight: 500;
		cursor: pointer;
		transition: all 0.15s ease;
		min-width: 100px;
		justify-content: center;
	}

	.pptx-nav-btn:hover:not(:disabled) {
		background: #35355a;
		border-color: #4a4a7c;
		color: #ffffff;
	}

	.pptx-nav-btn:disabled {
		opacity: 0.35;
		cursor: not-allowed;
	}

	.nav-text {
		display: inline;
	}

	.pptx-slide-info {
		display: flex;
		flex-direction: column;
		align-items: center;
		gap: 8px;
	}

	.pptx-slide-indicator {
		display: flex;
		align-items: center;
		gap: 4px;
		font-size: 13px;
		color: #8a8aaa;
	}

	.pptx-slide-current {
		font-weight: 600;
		color: #ffffff;
		font-size: 15px;
	}

	.pptx-slide-separator {
		color: #5a5a7a;
	}

	.pptx-slide-total {
		color: #8a8aaa;
	}

	/* Thumbnails */
	.pptx-thumbnails {
		display: flex;
		gap: 6px;
		overflow-x: auto;
		padding: 4px;
		max-width: 320px;
	}

	.pptx-thumbnail {
		width: 36px;
		height: 26px;
		background: #2a2a45;
		border: 2px solid #3a3a5c;
		border-radius: 4px;
		cursor: pointer;
		display: flex;
		align-items: center;
		justify-content: center;
		transition: all 0.15s ease;
		flex-shrink: 0;
	}

	.pptx-thumbnail:hover {
		border-color: #c44e23;
		background: #35355a;
	}

	.pptx-thumbnail.active {
		border-color: #c44e23;
		background: #c44e23;
	}

	.pptx-thumbnail-number {
		font-size: 10px;
		font-weight: 600;
		color: #8a8aaa;
	}

	.pptx-thumbnail.active .pptx-thumbnail-number {
		color: white;
	}

	/* Animations */
	.animate-spin {
		animation: spin 1s linear infinite;
	}

	@keyframes spin {
		from { transform: rotate(0deg); }
		to { transform: rotate(360deg); }
	}

	/* Responsive adjustments */
	@media (max-width: 640px) {
		.nav-text {
			display: none;
		}
		
		.pptx-nav-btn {
			min-width: auto;
			padding: 8px 10px;
		}
		
		.pptx-title {
			max-width: 150px;
		}
		
		.pptx-thumbnails {
			max-width: 180px;
		}
	}

	/* Dark mode is default, but add light mode support */
	:global(.light) .pptx-viewer-wrapper {
		background: #f0f0f5;
	}
	
	:global(.light) .pptx-header {
		background: linear-gradient(180deg, #ffffff 0%, #f5f5fa 100%);
		border-bottom-color: #e0e0e8;
	}
	
	:global(.light) .pptx-title {
		color: #1a1a2e;
	}
</style>