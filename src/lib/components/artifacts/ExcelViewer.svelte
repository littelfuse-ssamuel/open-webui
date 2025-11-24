<script lang="ts">
	import { onMount, onDestroy } from 'svelte';
	import type { ExcelArtifact } from '$lib/types';
	// @ts-ignore - no types for file-saver
	import fileSaver from 'file-saver';
	// @ts-ignore - luckyexcel doesn't have types
	import LuckyExcel from 'luckyexcel';

	const { saveAs } = fileSaver;

	export let file: ExcelArtifact;

	let containerElement: HTMLDivElement;
	let loading = true;
	let error: string | null = null;
	let luckysheet: any = null;

	async function loadFortuneSheet() {
		try {
			loading = true;
			error = null;

			// Fetch the Excel file
			const resp = await fetch(file.url);
			if (!resp.ok) {
				throw new Error(`Failed to fetch file: ${resp.statusText}`);
			}

			const arrayBuffer = await resp.arrayBuffer();

			// Convert Excel to LuckySheet format using LuckyExcel
			LuckyExcel.transformExcelToLucky(
				arrayBuffer,
				async (exportJson: any, luckyFile: any) => {
					if (exportJson.sheets == null || exportJson.sheets.length === 0) {
						error = 'Failed to parse Excel file';
						loading = false;
						return;
					}

					// Dynamically import FortuneSheet
					// @ts-ignore - dynamic import
					const FortuneSheet = (await import('@fortune-sheet/core')).default;
					// Import FortuneSheet styles
					await import('@fortune-sheet/core/dist/index.css');

					// Get active sheet from metadata or use first sheet
					const activeSheetName = file.meta?.activeSheet || exportJson.sheets[0].name;
					const activeSheetIndex = exportJson.sheets.findIndex(
						(s: any) => s.name === activeSheetName
					);

					// Initialize FortuneSheet
					// @ts-ignore - FortuneSheet API
					luckysheet = FortuneSheet.create({
						container: containerElement,
						data: exportJson.sheets,
						options: {
							container: containerElement,
							showinfobar: false,
							showsheetbar: true,
							showsheetbarConfig: {
								add: false,
								menu: false
							},
							sheetFormulaBar: true,
							enableAddRow: false,
							enableAddCol: false,
							userInfo: false,
							myFolderUrl: '',
							title: file.name,
							lang: 'en',
							row: exportJson.sheets[0]?.row || 60,
							column: exportJson.sheets[0]?.column || 26,
							showtoolbar: true,
							showtoolbarConfig: {
								undoRedo: true,
								paintFormat: false,
								currencyFormat: false,
								percentageFormat: false,
								numberDecrease: false,
								numberIncrease: false,
								moreFormats: false,
								font: false,
								fontSize: false,
								bold: false,
								italic: false,
								strikethrough: false,
								underline: false,
								textColor: false,
								fillColor: false,
								border: false,
								mergeCell: false,
								horizontalAlignMode: false,
								verticalAlignMode: false,
								textWrapMode: false,
								textRotateMode: false,
								image: false,
								link: false,
								chart: false,
								postil: false,
								pivotTable: false,
								function: false,
								frozenMode: false,
								sortAndFilter: false,
								conditionalFormat: false,
								dataVerification: false,
								splitColumn: false,
								screenshot: false,
								findAndReplace: false,
								protection: false,
								print: false
							}
						}
					});

					// Set active sheet if specified
					if (activeSheetIndex > 0) {
						setTimeout(() => {
							luckysheet.setSheetActive(activeSheetIndex);
						}, 100);
					}

					loading = false;
				},
				(err: any) => {
					console.error('Error transforming Excel:', err);
					error = 'Failed to load Excel file';
					loading = false;
				}
			);
		} catch (e) {
			console.error('Error loading Excel file:', e);
			error = e instanceof Error ? e.message : 'Failed to load Excel file';
			loading = false;
		}
	}

	function downloadExcel() {
		if (!file?.url) return;

		// Download the original file
		fetch(file.url)
			.then((response) => response.blob())
			.then((blob) => {
				saveAs(blob, file.name || 'download.xlsx');
			})
			.catch((e) => {
				console.error('Error downloading Excel file:', e);
			});
	}

	onMount(() => {
		if (file?.url) {
			loadFortuneSheet();
		}
	});

	onDestroy(() => {
		if (luckysheet) {
			try {
				luckysheet.destroy();
			} catch (e) {
				// Ignore cleanup errors
			}
			luckysheet = null;
		}
	});

	$: if (file?.url) {
		if (luckysheet) {
			luckysheet.destroy();
			luckysheet = null;
		}
		loadFortuneSheet();
	}
</script>

<div class="excel-viewer">
	<div class="excel-header">
		<div class="excel-header-left">
			<span class="excel-filename">{file.name}</span>
		</div>
		<div class="excel-header-right">
			<button
				class="excel-download-button"
				on:click={downloadExcel}
				disabled={loading}
				title="Download Excel file"
			>
				<svg
					xmlns="http://www.w3.org/2000/svg"
					width="16"
					height="16"
					viewBox="0 0 24 24"
					fill="none"
					stroke="currentColor"
					stroke-width="2"
					stroke-linecap="round"
					stroke-linejoin="round"
				>
					<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
					<polyline points="7 10 12 15 17 10" />
					<line x1="12" y1="15" x2="12" y2="3" />
				</svg>
				Download
			</button>
		</div>
	</div>

	{#if loading}
		<div class="excel-loading">Loading Excel file...</div>
	{:else if error}
		<div class="excel-error">
			<strong>Error:</strong>
			{error}
		</div>
	{:else}
		<div class="excel-container" bind:this={containerElement}></div>
	{/if}
</div>

<style>
	.excel-viewer {
		display: flex;
		flex-direction: column;
		height: 100%;
		background: var(--color-gray-50);
		border-radius: 0.5rem;
		overflow: hidden;
	}

	.excel-header {
		display: flex;
		justify-content: space-between;
		align-items: center;
		padding: 0.75rem 1rem;
		background: white;
		border-bottom: 1px solid var(--color-gray-200);
		gap: 0.5rem;
		flex-wrap: wrap;
		z-index: 10;
	}

	.excel-header-left {
		display: flex;
		align-items: center;
		gap: 0.75rem;
		flex: 1;
		min-width: 0;
	}

	.excel-header-right {
		display: flex;
		align-items: center;
		gap: 0.5rem;
	}

	.excel-filename {
		font-weight: 600;
		color: #000;
		white-space: nowrap;
		overflow: hidden;
		text-overflow: ellipsis;
	}

	.excel-download-button {
		display: flex;
		align-items: center;
		gap: 0.375rem;
		padding: 0.375rem 0.75rem;
		border: none;
		border-radius: 0.25rem;
		font-size: 0.875rem;
		font-weight: 600;
		cursor: pointer;
		transition: background 0.2s;
		background: var(--color-gray-600);
		color: white;
	}

	.excel-download-button:hover:not(:disabled) {
		background: var(--color-gray-700);
	}

	.excel-download-button:disabled {
		background: var(--color-gray-300);
		color: var(--color-gray-500);
		cursor: not-allowed;
	}

	.excel-loading,
	.excel-error {
		padding: 2rem;
		text-align: center;
		color: var(--color-gray-600);
	}

	.excel-error {
		color: var(--color-red-600);
	}

	.excel-container {
		flex: 1;
		overflow: hidden;
		position: relative;
		background: white;
	}

	/* Override FortuneSheet container to fill space */
	.excel-container :global(.luckysheet) {
		height: 100% !important;
	}
</style>
