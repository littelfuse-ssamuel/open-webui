<script lang="ts">
	import { onMount, onDestroy, createEventDispatcher } from 'svelte';
	import { toast } from 'svelte-sonner';
	import type { ExcelArtifact, ExcelCellChange } from '$lib/types/excel';
	import { isValidExcelArtifact, EXCEL_MIME_TYPE } from '$lib/types/excel';
	import { excelCore } from '$lib/services/excel-core';

	const dispatch = createEventDispatcher();

	export let file: ExcelArtifact;

	// Constants
	// Delay to allow fullscreen transition to complete before triggering Univer resize
	const FULLSCREEN_TRANSITION_DELAY = 100;

	// State
	let containerElement: HTMLDivElement;
	let wrapperElement: HTMLDivElement;
	let loading = true;
	let error: string | null = null;
	let saving = false;
	let saveMessage = '';
	let hasUnsavedChanges = false;
	let dirtyChanges: Map<string, ExcelCellChange> = new Map();
	let isFullscreen = false;
	let hasCharts = false;
	// Univer instances (dynamically loaded)
	let univer: any = null;
	let univerAPI: any = null;
	let workbookData: any = null;

	// Track if Univer modules are loaded
	let univerLoaded = false;

	// Dynamically import Univer modules
	async function loadUniverModules() {
		try {
			const [
				// Core modules
				{ Univer, LocaleType, UniverInstanceType },
				{ defaultTheme },
				{ UniverRenderEnginePlugin },
				{ UniverFormulaEnginePlugin },
				{ UniverUIPlugin },
				{ UniverDocsPlugin },
				{ UniverDocsUIPlugin },
				{ UniverSheetsPlugin },
				{ UniverSheetsUIPlugin },
				{ UniverSheetsFormulaPlugin },
				{ UniverSheetsFormulaUIPlugin },
				{ FUniver },
				{ UniverDrawingPlugin },
				{ UniverDrawingUIPlugin },
				{ UniverSheetsDrawingPlugin },
				{ UniverSheetsDrawingUIPlugin },
				{ UniverSheetsChartPlugin },
				{ UniverSheetsChartUIPlugin },
				{ UniverSheetsFilterPlugin },
				{ UniverSheetsFilterUIPlugin },
				{ UniverSheetsSortPlugin },
				{ UniverSheetsSortUIPlugin }
			] = await Promise.all([
				// Core modules
				import('@univerjs/core'),
				import('@univerjs/design'),
				import('@univerjs/engine-render'),
				import('@univerjs/engine-formula'),
				import('@univerjs/ui'),
				import('@univerjs/docs'),
				import('@univerjs/docs-ui'),
				import('@univerjs/sheets'),
				import('@univerjs/sheets-ui'),
				import('@univerjs/sheets-formula'),
				import('@univerjs/sheets-formula-ui'),
				import('@univerjs/facade'),
				import('@univerjs/drawing'),
				import('@univerjs/drawing-ui'),
				import('@univerjs/sheets-drawing'),
				import('@univerjs/sheets-drawing-ui'),
				import('@univerjs-pro/sheets-chart'),
				import('@univerjs-pro/sheets-chart-ui'),
				import('@univerjs/sheets-filter'),
				import('@univerjs/sheets-filter-ui'),
				import('@univerjs/sheets-sort'),
				import('@univerjs/sheets-sort-ui')
			]);

			// Import styles
			await Promise.all([
				import('@univerjs/design/lib/index.css'),
				import('@univerjs/ui/lib/index.css'),
				import('@univerjs/docs-ui/lib/index.css'),
				import('@univerjs/sheets-ui/lib/index.css'),
				import('@univerjs/sheets-formula-ui/lib/index.css'),
				import('@univerjs/drawing-ui/lib/index.css'),
				import('@univerjs/sheets-drawing-ui/lib/index.css'),
				import('@univerjs-pro/sheets-chart-ui/lib/index.css'),
				import('@univerjs/sheets-filter-ui/lib/index.css'),
				import('@univerjs/sheets-sort-ui/lib/index.css')
			]);

			// Import locale
			const [
				{ default: DesignEnUS },
				{ default: UIEnUS },
				{ default: DocsUIEnUS },
				{ default: SheetsEnUS },
				{ default: SheetsUIEnUS },
				{ default: SheetsFormulaUIEnUS },
				{ default: DrawingUIEnUS },
				{ default: SheetsDrawingUIEnUS },
				{ default: SheetsChartEnUS },
				{ default: SheetsChartUIEnUS },
				{ default: SheetsFilterUIEnUS },
				{ default: SheetsSortUIEnUS }
			] = await Promise.all([
				import('@univerjs/design/locale/en-US'),
				import('@univerjs/ui/locale/en-US'),
				import('@univerjs/docs-ui/locale/en-US'),
				import('@univerjs/sheets/locale/en-US'),
				import('@univerjs/sheets-ui/locale/en-US'),
				import('@univerjs/sheets-formula-ui/locale/en-US'),
				import('@univerjs/drawing-ui/locale/en-US'),
				import('@univerjs/sheets-drawing-ui/locale/en-US'),
				import('@univerjs-pro/sheets-chart/locale/en-US'),
				import('@univerjs-pro/sheets-chart-ui/locale/en-US'),
				import('@univerjs/sheets-filter-ui/locale/en-US'),
				import('@univerjs/sheets-sort-ui/locale/en-US')
			]);

			return {
				Univer,
				LocaleType,
				UniverInstanceType,
				defaultTheme,
				UniverRenderEnginePlugin,
				UniverFormulaEnginePlugin,
				UniverUIPlugin,
				UniverDocsPlugin,
				UniverDocsUIPlugin,
				UniverSheetsPlugin,
				UniverSheetsUIPlugin,
				UniverSheetsFormulaPlugin,
				UniverSheetsFormulaUIPlugin,
				FUniver,
				UniverDrawingPlugin,
				UniverDrawingUIPlugin,
				UniverSheetsDrawingPlugin,
				UniverSheetsDrawingUIPlugin,
				UniverSheetsChartPlugin,
				UniverSheetsChartUIPlugin,
				UniverSheetsFilterPlugin,
				UniverSheetsFilterUIPlugin,
				UniverSheetsSortPlugin,
				UniverSheetsSortUIPlugin,
				locales: {
					...DesignEnUS,
					...UIEnUS,
					...DocsUIEnUS,
					...SheetsEnUS,
					...SheetsUIEnUS,
					...SheetsFormulaUIEnUS,
					...DrawingUIEnUS,
					...SheetsDrawingUIEnUS,
					...SheetsChartEnUS,
					...SheetsChartUIEnUS,
					...SheetsFilterUIEnUS,
					...SheetsSortUIEnUS
				}
			};
		} catch (e) {
			console.error('Failed to load Univer modules:', e);
			throw new Error('Failed to load spreadsheet components. Please refresh the page.');
		}
	}

	function detectChartsInWorkbook(workbook: any): boolean {
		try {
			// Method 1: Check workbook's internal file list if available
			if (workbook.Directory) {
				const hasChartFiles = Object.keys(workbook.Directory).some(
					(key) => key.includes('chart') || key.includes('drawing')
				);
				if (hasChartFiles) return true;
			}

			// Method 2: Check for chart-related properties in sheets
			for (const sheetName of workbook.SheetNames) {
				const sheet = workbook.Sheets[sheetName];
				// Check for drawings reference
				if (sheet['!drawings'] || sheet['!charts']) {
					return true;
				}
				// Check for legacy chart indicators
				if (sheet['!objects'] && Array.isArray(sheet['!objects'])) {
					const hasChart = sheet['!objects'].some((obj: any) => 
						obj?.Type === 'Chart' || obj?.type === 'chart'
					);
					if (hasChart) return true;
				}
			}

			// Method 3: Check workbook metadata
			if (workbook.Workbook?.Sheets) {
				const hasChartSheet = workbook.Workbook.Sheets.some(
					(s: any) => s?.Chart || s?.Drawing
				);
				if (hasChartSheet) return true;
			}

			return false;
		} catch (e) {
			console.warn('Error detecting charts:', e);
			return false;
		}
	}

	// Convert XLSX ArrayBuffer to Univer workbook data format using xlsx library
	async function convertXLSXToUniverData(arrayBuffer: ArrayBuffer) {
		const XLSX = await import('xlsx');
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellFormula: true, cellStyles: true });
		hasCharts = detectChartsInWorkbook(workbook);
		// Build Univer-compatible workbook data
		const sheets: Record<string, any> = {};
		let sheetOrder: string[] = [];

		workbook.SheetNames.forEach((sheetName, sheetIndex) => {
			const ws = workbook.Sheets[sheetName];
			const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

			const cellData: Record<number, Record<number, any>> = {};

			for (let r = range.s.r; r <= range.e.r; r++) {
				cellData[r] = {};
				for (let c = range.s.c; c <= range.e.c; c++) {
					const cellAddress = XLSX.utils.encode_cell({ r, c });
					const cell = ws[cellAddress];

					if (cell) {
						const cellObj: any = {
							v: cell.v ?? '',
							t: cell.t === 'n' ? 1 : cell.t === 'b' ? 2 : 0 // CellValueType: String=0, Number=1, Boolean=2
						};

						// Handle formulas
						if (cell.f) {
							cellObj.f = cell.f;
						}

						// Handle styles (basic)
						if (cell.s) {
							cellObj.s = convertCellStyle(cell.s);
						}

						cellData[r][c] = cellObj;
					}
				}
			}

			// Convert column widths from SheetJS to Univer format
			const columnData: Record<number, any> = {};
			if (ws['!cols']) {
				ws['!cols'].forEach((col: any, idx: number) => {
					if (col) {
						columnData[idx] = {
							w: col.wpx || (col.wch ? Math.round(col.wch * 7.5) : undefined),
							hd: col.hidden ? 1 : 0
						};
					}
				});
			}

			// Convert row heights from SheetJS to Univer format
			const rowData: Record<number, any> = {};
			if (ws['!rows']) {
				ws['!rows'].forEach((row: any, idx: number) => {
					if (row) {
						rowData[idx] = {
							h: row.hpx || row.hpt || undefined,
							hd: row.hidden ? 1 : 0
						};
					}
				});
			}

			// Convert merged cells from SheetJS to Univer format
			const mergeData: any[] = [];
			if (ws['!merges']) {
				ws['!merges'].forEach((merge: any) => {
					mergeData.push({
						startRow: merge.s.r,
						startColumn: merge.s.c,
						endRow: merge.e.r,
						endColumn: merge.e.c
					});
				});
			}

			const sheetId = `sheet_${sheetIndex}`;
			sheetOrder.push(sheetId);
			let freeze: any = undefined;
			if (ws['!freeze']) {
				freeze = {
					startRow: ws['!freeze'].ySplit || 0,
					startColumn: ws['!freeze'].xSplit || 0,
					ySplit: ws['!freeze'].ySplit || 0,
					xSplit: ws['!freeze'].xSplit || 0
				};
			}

			sheets[sheetId] = {
				id: sheetId,
				name: sheetName,
				rowCount: Math.max(range.e.r + 1, 100),
				columnCount: Math.max(range.e.c + 1, 26),
				cellData,
				rowData,
				columnData,
				mergeData,
				freeze,
				defaultRowHeight: 24,
				defaultColumnWidth: 88
			};
		});

		return {
			id: 'workbook_1',
			sheetOrder,
			name: file.name || 'Workbook',
			appVersion: '1.0.0',
			locale: 'en-US' as any,
			styles: {},
			sheets
		};
	}

	// Convert xlsx cell style to Univer style format
	function convertCellStyle(xlsxStyle: any): any {
		const style: any = {};

		if (xlsxStyle.font) {
			if (xlsxStyle.font.bold) style.bl = 1;
			if (xlsxStyle.font.italic) style.it = 1;
			if (xlsxStyle.font.underline) style.ul = { s: 1 };
			if (xlsxStyle.font.sz) style.fs = xlsxStyle.font.sz;
			if (xlsxStyle.font.color?.rgb) style.cl = { rgb: xlsxStyle.font.color.rgb };
			if (xlsxStyle.font.name) style.ff = xlsxStyle.font.name; // Font family
		}

		if (xlsxStyle.fill?.fgColor?.rgb) {
			style.bg = { rgb: xlsxStyle.fill.fgColor.rgb };
		}

		if (xlsxStyle.alignment) {
			if (xlsxStyle.alignment.horizontal) {
				const hMap: Record<string, number> = { left: 1, center: 2, right: 3 };
				style.ht = hMap[xlsxStyle.alignment.horizontal] || 1;
			}
			if (xlsxStyle.alignment.vertical) {
				const vMap: Record<string, number> = { top: 1, center: 2, bottom: 3 };
				style.vt = vMap[xlsxStyle.alignment.vertical] || 2;
			}
			if (xlsxStyle.alignment.wrapText) {
				style.tb = 1; // Text wrap
			}
		}

		// Border handling
		if (xlsxStyle.border) {
			style.bd = {};
			const mapBorderStyle = (styleName: string): number => {
				const styleMap: Record<string, number> = {
					thin: 1, medium: 2, thick: 3, dashed: 4, dotted: 5, double: 6
				};
				return styleMap[styleName] || 1;
			};
			const mapBorder = (b: any) => {
				if (!b) return undefined;
				return { s: mapBorderStyle(b.style || 'thin'), cl: { rgb: b.color?.rgb || '000000' } };
			};
			if (xlsxStyle.border.top) style.bd.t = mapBorder(xlsxStyle.border.top);
			if (xlsxStyle.border.right) style.bd.r = mapBorder(xlsxStyle.border.right);
			if (xlsxStyle.border.bottom) style.bd.b = mapBorder(xlsxStyle.border.bottom);
			if (xlsxStyle.border.left) style.bd.l = mapBorder(xlsxStyle.border.left);
		}

		return style;
	}

	// Initialize Univer with workbook data
	async function initUniver(data: any) {
		if (!containerElement) return;

		const modules = await loadUniverModules();
		const {
			Univer,
			LocaleType,
			UniverInstanceType,
			defaultTheme,
			UniverRenderEnginePlugin,
			UniverFormulaEnginePlugin,
			UniverUIPlugin,
			UniverDocsPlugin,
			UniverDocsUIPlugin,
			UniverSheetsPlugin,
			UniverSheetsUIPlugin,
			UniverSheetsFormulaPlugin,
			UniverSheetsFormulaUIPlugin,
			FUniver,
			UniverDrawingPlugin,
			UniverDrawingUIPlugin,
			UniverSheetsDrawingPlugin,
			UniverSheetsDrawingUIPlugin,
			UniverSheetsChartPlugin,
			UniverSheetsChartUIPlugin,
			UniverSheetsFilterPlugin,
			UniverSheetsFilterUIPlugin,
			UniverSheetsSortPlugin,
			UniverSheetsSortUIPlugin,
			locales
		} = modules;

		// Create Univer instance
		univer = new Univer({
			theme: defaultTheme,
			locale: LocaleType.EN_US,
			locales: {
				[LocaleType.EN_US]: locales
			}
		});

		// Register plugins
		univer.registerPlugin(UniverRenderEnginePlugin);
		univer.registerPlugin(UniverFormulaEnginePlugin);
		univer.registerPlugin(UniverUIPlugin, {
			container: containerElement,
			// FIX #2: Enable footer to show sheet tabs for switching between sheets
			footer: true
		});
		univer.registerPlugin(UniverDocsPlugin);
		univer.registerPlugin(UniverDocsUIPlugin);
		univer.registerPlugin(UniverSheetsPlugin, {
			notExecuteFormula: false
		});
		univer.registerPlugin(UniverSheetsUIPlugin, {
			clipboardConfig: {
				enabled: true
			}
		});
		univer.registerPlugin(UniverSheetsFormulaPlugin);
		// FIX #1: Register Formula UI plugin for editable formula bar
		univer.registerPlugin(UniverSheetsFormulaUIPlugin);

		// FIX #3: Register Drawing plugins (required for charts and floating images)
		univer.registerPlugin(UniverDrawingPlugin);
		univer.registerPlugin(UniverDrawingUIPlugin);
		univer.registerPlugin(UniverSheetsDrawingPlugin);
		univer.registerPlugin(UniverSheetsDrawingUIPlugin);

		univer.registerPlugin(UniverSheetsChartPlugin);
		univer.registerPlugin(UniverSheetsChartUIPlugin);

		univer.registerPlugin(UniverSheetsFilterPlugin);
		univer.registerPlugin(UniverSheetsFilterUIPlugin);

		univer.registerPlugin(UniverSheetsSortPlugin);
		univer.registerPlugin(UniverSheetsSortUIPlugin);

		// Create workbook with data
		univer.createUnit(UniverInstanceType.UNIVER_SHEET, data);

		// Get Facade API for easier interaction
		univerAPI = FUniver.newAPI(univer);

		// Note: Univer handles focus automatically when users click cells
		// Do not manually manipulate focus/tabindex as it interferes with Univer's internal editor system

		// Note: Univer handles focus automatically when users click cells
		// Do not manually manipulate focus/tabindex as it interferes with Univer's internal editor system

		// Listen for changes to track unsaved state AND capture actual cell changes
		univerAPI.onCommandExecuted((command: any) => {
			const editCommands = [
				'sheet.mutation.set-range-values',
				'sheet.command.set-range-values',
				'sheet.mutation.insert-row',
				'sheet.mutation.insert-col',
				'sheet.mutation.remove-row',
				'sheet.mutation.remove-col'
			];

			// Check if this is an edit command
			if (editCommands.some((cmd) => command.id?.includes(cmd) || command.id?.includes('set'))) {
				hasUnsavedChanges = true;
				saveMessage = '';

				// For set-range-values commands, extract the actual changed cells
				if (command.id?.includes('set-range-values') && command.params) {
					try {
						const params = command.params;
						const subUnitId = params.subUnitId; // Sheet ID
						const cellValue = params.cellValue; // The cell data object
						
						if (subUnitId && cellValue) {
							// cellValue is structured as { [row]: { [col]: cellData } }
							Object.entries(cellValue).forEach(([rowStr, rowData]: [string, any]) => {
								if (rowData && typeof rowData === 'object') {
									Object.entries(rowData).forEach(([colStr, cellData]: [string, any]) => {
										const row = parseInt(rowStr) + 1; // Convert to 1-indexed
										const col = parseInt(colStr) + 1; // Convert to 1-indexed
										
										// Extract value - check for formula first
										const hasFormula = !!(cellData?.f);
										const value = hasFormula ? `=${cellData.f}` : cellData?.v;
										
										// Only track if we have a valid value (not undefined)
										if (value !== undefined) {
											const key = `${subUnitId}:${row}:${col}`;
											dirtyChanges.set(key, {
												row,
												col,
												value,
												isFormula: hasFormula
											});
											console.debug(`Tracked change: ${key} = ${value}`);
										}
									});
								}
							});
						}
					} catch (e) {
						console.warn('Could not extract cell changes from command:', e);
					}
				}
			}
		});

		univerLoaded = true;
		console.log('Univer initialized successfully with formula bar, sheet tabs, and chart support');
	}

	// Load workbook from URL
	async function loadWorkbook() {
		if (!file?.url) return;

		// Phase 2: Validate artifact structure
		if (!isValidExcelArtifact(file)) {
			console.error('Invalid Excel artifact structure:', file);
			error = 'Invalid Excel file data';
			loading = false;
			return;
		}

		// Phase 2: Add file type validation
		const fileName = file.name?.toLowerCase() || '';
		const validExtensions = ['.xlsx', '.xlsm', '.xls'];
		const hasValidExtension = validExtensions.some((ext) => fileName.endsWith(ext));

		if (!hasValidExtension) {
			console.warn('File does not have standard Excel extension:', fileName);
			// Continue anyway - file might still be valid Excel
		}

		try {
			loading = true;
			dirtyChanges.clear();
			error = null;

			// Use excelCore service for consistent fetch with cache-busting
			const arrayBuffer = await excelCore.fetchExcelFile(file.url);
			workbookData = await convertXLSXToUniverData(arrayBuffer);

			await initUniver(workbookData);

			loading = false;

			// Force Univer to recalculate dimensions after loading spinner hides
			// and the container becomes visible with its final flex-computed size
			requestAnimationFrame(() => {
				window.dispatchEvent(new Event('resize'));
			});
		} catch (e) {
			console.error('Error loading workbook:', e);
			error = e instanceof Error ? e.message : 'Failed to load Excel file';
			loading = false;
		}
	}

	// Save only the cells that were actually changed (dirty cell tracking)
	async function saveChanges() {
		if (!univerAPI || !file.fileId) {
			saveMessage = 'No changes to save';
			return;
		}

		if (dirtyChanges.size === 0) {
			saveMessage = 'No changes to save';
			hasUnsavedChanges = false;
			return;
		}

		try {
			saving = true;
			saveMessage = '';

			// Get current workbook data from Univer for sheet name resolution
			const workbook = univerAPI.getActiveWorkbook();
			if (!workbook) {
				throw new Error('No active workbook');
			}

			// End any active cell editing to ensure data is synced
			try {
				await workbook.endEditingAsync(true);
			} catch (e) {
				console.warn('endEditingAsync not available, trying command fallback');
				try {
					univerAPI.executeCommand('sheet.operation.set-cell-edit-visible', {
						visible: false,
						_eventType: 2
					});
					await new Promise(resolve => setTimeout(resolve, 50));
				} catch (cmdErr) {
					console.warn('Could not end editing:', cmdErr);
				}
			}

			const snapshot = workbook.save();

			// Group dirty changes by sheet
			const changesBySheet = new Map<string, { sheetName: string; changes: ExcelCellChange[] }>();

			dirtyChanges.forEach((change, key) => {
				const [subUnitId] = key.split(':');
				
				// Get sheet name from snapshot
				const sheetData = snapshot.sheets?.[subUnitId];
				const sheetName = sheetData?.name || 'Sheet1';

				if (!changesBySheet.has(subUnitId)) {
					changesBySheet.set(subUnitId, { sheetName, changes: [] });
				}
				changesBySheet.get(subUnitId)!.changes.push(change);
			});

			let totalChangesApplied = 0;
			const errors: string[] = [];

			// Save changes for each sheet
			for (const [subUnitId, { sheetName, changes }] of changesBySheet) {
				if (changes.length > 0) {
					try {
						await excelCore.saveChanges({
							fileId: file.fileId!,
							sheet: sheetName,
							changes
						});
						totalChangesApplied += changes.length;
						console.log(`Saved ${changes.length} changed cells to sheet "${sheetName}"`);
					} catch (e) {
						errors.push(`Sheet "${sheetName}": ${e instanceof Error ? e.message : 'Save failed'}`);
					}
				}
			}

			if (errors.length > 0) {
				throw new Error(errors.join('; '));
			}

			// Clear dirty changes after successful save
			dirtyChanges.clear();
			hasUnsavedChanges = false;
			saveMessage = `Successfully saved ${totalChangesApplied} cell(s)`;
			saving = false;

			// Auto-dismiss success message
			setTimeout(() => {
				if (saveMessage && !saveMessage.includes('Failed')) {
					saveMessage = '';
				}
			}, 3000);

		} catch (e) {
			console.error('Error saving changes:', e);
			saveMessage = e instanceof Error ? e.message : 'Failed to save changes';
			saving = false;
		}
	}

	// Download the current workbook
	async function downloadExcel() {
		if (!univerAPI) {
			toast.error('No workbook loaded');
			return;
		}

		try {
			const workbook = univerAPI.getActiveWorkbook();
			if (!workbook) {
				throw new Error('No active workbook');
			}

			const snapshot = workbook.save();

			// Convert Univer snapshot back to xlsx format
			const XLSX = await import('xlsx');
			const wb = XLSX.utils.book_new();

			// Process each sheet
			Object.values(snapshot.sheets || {}).forEach((sheetData: any) => {
				const wsData: any[][] = [];
				const cellData = sheetData.cellData || {};

				// Find max row and column
				let maxRow = 0;
				let maxCol = 0;
				Object.entries(cellData).forEach(([rowStr, rowData]: [string, any]) => {
					maxRow = Math.max(maxRow, parseInt(rowStr));
					Object.keys(rowData).forEach((colStr) => {
						maxCol = Math.max(maxCol, parseInt(colStr));
					});
				});

				// Build 2D array
				for (let r = 0; r <= maxRow; r++) {
					wsData[r] = [];
					for (let c = 0; c <= maxCol; c++) {
						const cell = cellData[r]?.[c];
						if (cell) {
							wsData[r][c] = cell.f ? { f: cell.f, v: cell.v } : cell.v;
						} else {
							wsData[r][c] = '';
						}
					}
				}

				const ws = XLSX.utils.aoa_to_sheet(wsData);
				XLSX.utils.book_append_sheet(wb, ws, sheetData.name || 'Sheet1');
			});

			// Generate and download
			const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
			const blob = new Blob([wbout], {
				type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
			});

			// Use excelCore service for consistent download
			await excelCore.downloadExcel(blob, file.name || 'download.xlsx');
		} catch (e) {
			console.error('Error downloading Excel file:', e);
			toast.error('Failed to download file');
		}
	}

	// Toggle fullscreen
	function toggleFullscreen() {
		if (!wrapperElement) return;

		if (!document.fullscreenElement) {
			wrapperElement.requestFullscreen?.();
		} else {
			document.exitFullscreen?.();
		}
		// Note: isFullscreen state is updated via handleFullscreenChange event listener
		// which ensures it stays in sync regardless of how fullscreen is entered/exited
	}

	// Handle fullscreen change
	function handleFullscreenChange() {
		isFullscreen = !!document.fullscreenElement;
		
		// Force Univer to recalculate by explicitly setting container dimensions
		// Univer's render engine uses ResizeObserver on its container, but since
		// the CSS width stays "100%" (of a parent that was sized at pane width),
		// the ResizeObserver doesn't detect the change. Setting explicit pixel
		// values forces a true dimension change that Univer will pick up.
		requestAnimationFrame(() => {
			if (containerElement) {
				if (isFullscreen) {
					// Set explicit viewport dimensions to force ResizeObserver trigger
					containerElement.style.width = `${window.innerWidth}px`;
					containerElement.style.height = `${window.innerHeight - (containerElement.getBoundingClientRect().top - wrapperElement.getBoundingClientRect().top)}px`;
				} else {
					// Reset to CSS-driven sizing
					containerElement.style.width = '100%';
					containerElement.style.height = '';
				}
			}
			// Also dispatch resize as a fallback
			setTimeout(() => {
				window.dispatchEvent(new Event('resize'));
			}, 100);
		});
 	}

	// Warn about unsaved changes before leaving
	function handleBeforeUnload(e: BeforeUnloadEvent) {
		if (hasUnsavedChanges) {
			e.preventDefault();
			e.returnValue = '';
			return '';
		}
	}

	onMount(() => {
		document.addEventListener('fullscreenchange', handleFullscreenChange);
		window.addEventListener('beforeunload', handleBeforeUnload);
	});

	onDestroy(() => {
		document.removeEventListener('fullscreenchange', handleFullscreenChange);
		window.removeEventListener('beforeunload', handleBeforeUnload);

		// Cleanup Univer instance
		if (univer) {
			try {
				univer.dispose();
			} catch (e) {
				console.warn('Error disposing Univer:', e);
			}
			univer = null;
			univerAPI = null;
		}
	});

	// Add this variable to track the loaded URL preventing loops
	let currentLoadedUrl = '';

	// Reactive: reload ONLY if the specific file URL has changed
	$: if (file?.url && containerElement && file.url !== currentLoadedUrl) {
		currentLoadedUrl = file.url;
		// Reset loading state if needed when URL changes
		univerLoaded = false; 
		loadWorkbook();
	}

	// Dispatch unsaved changes state to parent
	$: dispatch('unsavedChanges', { hasUnsavedChanges });
</script>

<div class="excel-viewer-wrapper" bind:this={wrapperElement}>
	<div class="excel-viewer">
		<!-- Header toolbar -->
		<div class="excel-header">
			<div class="excel-header-left">
				<span class="excel-filename" title={file.name}>{file.name}</span>
				{#if hasUnsavedChanges}
					<span class="excel-unsaved-indicator">•</span>
				{/if}
			</div>

			<div class="excel-header-right">
				{#if hasUnsavedChanges}
					<span class="excel-changes-badge">Unsaved changes</span>
				{/if}

				<button
					class="excel-toolbar-btn"
					on:click={toggleFullscreen}
					title={isFullscreen ? 'Exit fullscreen' : 'Fullscreen'}
				>
					{#if isFullscreen}
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<path d="M8 3v3a2 2 0 0 1-2 2H3m18 0h-3a2 2 0 0 1-2-2V3m0 18v-3a2 2 0 0 1 2-2h3M3 16h3a2 2 0 0 1 2 2v3"/>
						</svg>
					{:else}
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/>
						</svg>
					{/if}
				</button>

				<button
					class="excel-toolbar-btn"
					on:click={downloadExcel}
					disabled={loading || !univerLoaded}
					title="Download Excel file"
				>
					<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
						<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
						<polyline points="7 10 12 15 17 10"/>
						<line x1="12" y1="15" x2="12" y2="3"/>
					</svg>
				</button>

				<button
					class="excel-save-btn"
					on:click={saveChanges}
					disabled={saving || !hasUnsavedChanges || !file.fileId}
					title={!file.fileId ? 'Save not available for this file' : 'Save changes'}
				>
					{#if saving}
						<svg class="animate-spin" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
							<path d="M21 12a9 9 0 1 1-6.219-8.56"/>
						</svg>
						Saving...
					{:else}
						<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
							<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
							<polyline points="17 21 17 13 7 13 7 21"/>
							<polyline points="7 3 7 8 15 8"/>
						</svg>
						Save
					{/if}
				</button>
			</div>
		</div>

		<!-- Status message -->
		{#if saveMessage}
			<div class="excel-message" class:excel-message-success={!saveMessage.includes('Failed')}>
				{saveMessage}
			</div>
		{/if}

		<!-- Chart notice banner -->
		{#if hasCharts && !loading && !error}
			<div class="excel-chart-notice">
				<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
					<line x1="18" y1="20" x2="18" y2="10"/>
					<line x1="12" y1="20" x2="12" y2="4"/>
					<line x1="6" y1="20" x2="6" y2="14"/>
				</svg>
				<span>This file contains charts. Charts are preserved in the downloaded file but cannot be displayed in the web viewer.</span>
			</div>
		{/if}

		<!-- Main content area -->
		{#if loading}
			<div class="excel-loading">
				<div class="excel-loading-spinner"></div>
				<span>Loading spreadsheet...</span>
			</div>
		{/if}
		
		{#if error}
			<div class="excel-error">
				<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
					<circle cx="12" cy="12" r="10"/>
					<line x1="12" y1="8" x2="12" y2="12"/>
					<line x1="12" y1="16" x2="12.01" y2="16"/>
				</svg>
				<span>{error}</span>
				<button class="excel-retry-btn" on:click={loadWorkbook}>Retry</button>
			</div>
		{/if}
		
		<!-- Univer container (Always rendered so bind:this works) -->
		<div 
			class="excel-univer-container" 
			bind:this={containerElement}
			style:visibility={loading || error ? 'hidden' : 'visible'}
			style:display={error ? 'none' : 'flex'}
		></div>
	</div>
</div>

<style>
	.excel-viewer-wrapper {
		width: 100%;
		height: 100%;
		display: flex;
		flex-direction: column;
		background: #f8f9fa;
	}

	:global(.excel-viewer-wrapper:fullscreen) {
		background: white;
		height: 100vh !important;
		width: 100vw !important;
		position: fixed !important;
		top: 0 !important;
		left: 0 !important;
		z-index: 999999 !important;
	}

	:global(.excel-viewer-wrapper:fullscreen) .excel-viewer {
		height: 100vh;
	}

	:global(.excel-viewer-wrapper:fullscreen) .excel-univer-container {
		height: 100%;
	}
	
	.excel-viewer {
		display: flex;
		flex-direction: column;
		height: 100%;
		width: 100%;
		overflow: hidden;
	}

	.excel-header {
		display: flex;
		justify-content: space-between;
		align-items: center;
		padding: 8px 12px;
		background: white;
		border-bottom: 1px solid #e5e7eb;
		gap: 8px;
		flex-shrink: 0;
		z-index: 100;
	}

	.excel-header-left {
		display: flex;
		align-items: center;
		gap: 4px;
		min-width: 0;
		flex: 1;
	}

	.excel-filename {
		font-weight: 600;
		font-size: 14px;
		color: #1f2937;
		white-space: nowrap;
		overflow: hidden;
		text-overflow: ellipsis;
	}

	.excel-unsaved-indicator {
		color: #f59e0b;
		font-size: 24px;
		line-height: 1;
	}

	.excel-header-right {
		display: flex;
		align-items: center;
		gap: 8px;
		flex-shrink: 0;
	}

	.excel-changes-badge {
		padding: 4px 8px;
		background: #fef3c7;
		color: #92400e;
		border-radius: 4px;
		font-size: 12px;
		font-weight: 500;
	}

	.excel-toolbar-btn {
		display: flex;
		align-items: center;
		justify-content: center;
		padding: 6px;
		border: 1px solid #e5e7eb;
		border-radius: 6px;
		background: white;
		color: #4b5563;
		cursor: pointer;
		transition: all 0.15s;
	}

	.excel-toolbar-btn:hover:not(:disabled) {
		background: #f3f4f6;
		border-color: #d1d5db;
	}

	.excel-toolbar-btn:disabled {
		opacity: 0.5;
		cursor: not-allowed;
	}

	.excel-save-btn {
		display: flex;
		align-items: center;
		gap: 6px;
		padding: 6px 12px;
		border: none;
		border-radius: 6px;
		background: #2563eb;
		color: white;
		font-size: 13px;
		font-weight: 500;
		cursor: pointer;
		transition: all 0.15s;
	}

	.excel-save-btn:hover:not(:disabled) {
		background: #1d4ed8;
	}

	.excel-save-btn:disabled {
		background: #9ca3af;
		cursor: not-allowed;
	}

	.excel-message {
		padding: 8px 12px;
		background: #fee2e2;
		color: #991b1b;
		font-size: 13px;
		border-bottom: 1px solid #fecaca;
	}

	.excel-message-success {
		background: #dcfce7;
		color: #166534;
		border-bottom-color: #bbf7d0;
	}

	.excel-loading,
	.excel-error {
		flex: 1;
		display: flex;
		flex-direction: column;
		align-items: center;
		justify-content: center;
		gap: 12px;
		color: #6b7280;
		padding: 32px;
	}

	.excel-loading-spinner {
		width: 32px;
		height: 32px;
		border: 3px solid #e5e7eb;
		border-top-color: #2563eb;
		border-radius: 50%;
		animation: spin 1s linear infinite;
	}

	@keyframes spin {
		to {
			transform: rotate(360deg);
		}
	}

	.excel-error {
		color: #dc2626;
	}

	.excel-error svg {
		color: #dc2626;
	}

	.excel-retry-btn {
		padding: 8px 16px;
		background: #2563eb;
		color: white;
		border: none;
		border-radius: 6px;
		font-size: 13px;
		font-weight: 500;
		cursor: pointer;
		margin-top: 8px;
	}

	.excel-retry-btn:hover {
		background: #1d4ed8;
	}

	.excel-univer-container {
		flex: 1;
		width: 100%;
		height: 0;
		min-height: 0;
		overflow: hidden;
	}

	/* Override Univer's default styles to fit our container */
	.excel-univer-container :global(.univer-app) {
		height: 100% !important;
		width: 100% !important;
		max-width: 100% !important;
	}

	.excel-univer-container :global(.univer-container) {
		height: 100% !important;
		width: 100% !important;
		max-width: 100% !important;
	}

	.excel-univer-container :global(.univer-workbench),
	.excel-univer-container :global(.univer-workbench-container),
	.excel-univer-container :global(.univer-sheet-container) {
		width: 100% !important;
		max-width: 100% !important;
	}

	/* Fullscreen-specific overrides for Univer width */
	:global(.excel-viewer-wrapper:fullscreen) .excel-univer-container :global(.univer-app),
	:global(.excel-viewer-wrapper:fullscreen) .excel-univer-container :global(.univer-container),
	:global(.excel-viewer-wrapper:fullscreen) .excel-univer-container :global(.univer-workbench),
	:global(.excel-viewer-wrapper:fullscreen) .excel-univer-container :global(.univer-workbench-container),
	:global(.excel-viewer-wrapper:fullscreen) .excel-univer-container :global(.univer-sheet-container) {
		width: 100% !important;
		max-width: 100% !important;
	}

	/* Dark mode support */
	:global(.dark) .excel-viewer-wrapper {
		background: #1f2937;
	}

	:global(.dark) .excel-header {
		background: #374151;
		border-bottom-color: #4b5563;
	}

	:global(.dark) .excel-filename {
		color: #f3f4f6;
	}

	:global(.dark) .excel-toolbar-btn {
		background: #374151;
		border-color: #4b5563;
		color: #d1d5db;
	}

	:global(.dark) .excel-toolbar-btn:hover:not(:disabled) {
		background: #4b5563;
	}

	:global(.dark) .excel-changes-badge {
		background: #78350f;
		color: #fef3c7;
	}

	:global(.dark) .excel-loading,
	:global(.dark) .excel-error {
		color: #9ca3af;
	}

	/* Animate spin for save button */
	.animate-spin {
		animation: spin 1s linear infinite;
	}

	.excel-chart-notice {
		display: flex;
		align-items: center;
		gap: 8px;
		padding: 8px 12px;
		background: #eff6ff;
		color: #1e40af;
		font-size: 13px;
		border-bottom: 1px solid #bfdbfe;
	}

	.excel-chart-notice svg {
		flex-shrink: 0;
		color: #3b82f6;
	}

	:global(.dark) .excel-chart-notice {
		background: #1e3a5f;
		color: #93c5fd;
		border-bottom-color: #1e40af;
	}

	:global(.dark) .excel-chart-notice svg {
		color: #60a5fa;
	}
</style>