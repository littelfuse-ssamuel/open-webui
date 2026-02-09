<script lang="ts">
	import { SvelteFlowProvider } from '@xyflow/svelte';
	import { slide } from 'svelte/transition';
	import { Pane, PaneResizer } from 'paneforge';

	import { onDestroy, onMount, tick } from 'svelte';
	import {
		mobile,
		showControls,
		showCallOverlay,
		showOverview,
		showArtifacts,
		showEmbeds
	} from '$lib/stores';

	import Controls from './Controls/Controls.svelte';
	import CallOverlay from './MessageInput/CallOverlay.svelte';
	import Drawer from '../common/Drawer.svelte';
	import Artifacts from './Artifacts.svelte';
	import Embeds from './ChatControls/Embeds.svelte';
	const UNIVER_TRACE_KEY = 'lfrag.univer.trace';

	export let history;
	export let models = [];

	export let chatId = null;

	export let chatFiles = [];
	export let params = {};

	export let eventTarget: EventTarget;
	export let submitPrompt: Function;
	export let stopResponse: Function;
	export let showMessage: Function;
	export let files;
	export let modelId;

	export let pane;

	let mediaQuery;
	let largeScreen = false;
	let dragged = false;

	let minSize = 0;

	function isTraceEnabled() {
		if (typeof window === 'undefined') return false;
		try {
			if (localStorage.getItem(UNIVER_TRACE_KEY) === '1') return true;
			return new URLSearchParams(window.location.search).get('univerTrace') === '1';
		} catch {
			return false;
		}
	}

	function trace(event: string, data: Record<string, any> = {}) {
		if (!isTraceEnabled()) return;
		console.debug('[UNIVER_TRACE][ChatControls]', {
			event,
			at: new Date().toISOString(),
			showControls: $showControls,
			showArtifacts: $showArtifacts,
			minSize,
			...data
		});
	}

	export const openPane = () => {
		if (parseInt(localStorage?.chatControlsSize)) {
			const container = document.getElementById('chat-container');
			let size = Math.floor(
				(parseInt(localStorage?.chatControlsSize) / container.clientWidth) * 100
			);
			trace('openPane:storedSize', {
				containerWidth: container?.clientWidth ?? null,
				storedPixels: parseInt(localStorage?.chatControlsSize),
				calculatedPercent: size
			});
			pane.resize(size);
		} else {
			trace('openPane:minSize', { targetPercent: minSize });
			pane.resize(minSize);
		}
	};

	const handleMediaQuery = async (e) => {
		trace('handleMediaQuery', { matches: e.matches });
		if (e.matches) {
			largeScreen = true;

			if ($showCallOverlay) {
				showCallOverlay.set(false);
				await tick();
				showCallOverlay.set(true);
			}
		} else {
			largeScreen = false;

			if ($showCallOverlay) {
				showCallOverlay.set(false);
				await tick();
				showCallOverlay.set(true);
			}
			pane = null;
		}
	};

	const onMouseDown = (event) => {
		if (event.target.closest('#controls-resizer')) {
			dragged = true;
		}
	};

	const onMouseUp = (event) => {
		dragged = false;
	};

	onMount(() => {
		trace('onMount');
		// listen to resize 1024px
		mediaQuery = window.matchMedia('(min-width: 1024px)');

		mediaQuery.addEventListener('change', handleMediaQuery);
		handleMediaQuery(mediaQuery);

		// Select the container element you want to observe
		const container = document.getElementById('chat-container');

		// initialize the minSize based on the container width
		minSize = Math.floor((350 / container.clientWidth) * 100);

		// Create a new ResizeObserver instance
		const resizeObserver = new ResizeObserver((entries) => {
			for (let entry of entries) {
				const width = entry.contentRect.width;
				// calculate the percentage of 350px
				const percentage = (350 / width) * 100;
				// set the minSize to the percentage, must be an integer
				minSize = Math.floor(percentage);
				trace('chatContainer:resizeObserver', {
					containerWidth: width,
					calculatedMinSize: minSize,
					paneSize: pane?.getSize?.() ?? null
				});

				if ($showControls) {
					if (pane && pane.isExpanded() && pane.getSize() < minSize) {
						pane.resize(minSize);
					} else {
						let size = Math.floor(
							(parseInt(localStorage?.chatControlsSize) / container.clientWidth) * 100
						);
						if (size < minSize) {
							pane.resize(minSize);
						}
					}
				}
			}
		});

		// Start observing the container's size changes
		resizeObserver.observe(container);

		document.addEventListener('mousedown', onMouseDown);
		document.addEventListener('mouseup', onMouseUp);
	});

	onDestroy(() => {
		trace('onDestroy');
		showControls.set(false);

		mediaQuery.removeEventListener('change', handleMediaQuery);
		document.removeEventListener('mousedown', onMouseDown);
		document.removeEventListener('mouseup', onMouseUp);
	});

	const closeHandler = () => {
		showControls.set(false);
		showOverview.set(false);
		showArtifacts.set(false);
		showEmbeds.set(false);

		if ($showCallOverlay) {
			showCallOverlay.set(false);
		}
	};

	$: if (!chatId) {
		closeHandler();
	}
</script>

{#if !largeScreen}
	{#if $showControls}
		<Drawer
			show={$showControls}
			onClose={() => {
				showControls.set(false);
			}}
		>
			<div
				class=" {$showCallOverlay || $showOverview || $showArtifacts || $showEmbeds
					? ' h-screen  w-full'
					: 'px-4 py-3'} h-full"
			>
				{#if $showCallOverlay}
					<div
						class=" h-full max-h-[100dvh] bg-white text-gray-700 dark:bg-black dark:text-gray-300 flex justify-center"
					>
						<CallOverlay
							bind:files
							{submitPrompt}
							{stopResponse}
							{modelId}
							{chatId}
							{eventTarget}
							on:close={() => {
								showControls.set(false);
							}}
						/>
					</div>
				{:else if $showEmbeds}
					<Embeds />
				{:else if $showArtifacts}
					<Artifacts {history} />
				{:else if $showOverview}
					{#await import('./Overview.svelte') then { default: Overview }}
						<Overview
							{history}
							onNodeClick={(e) => {
								const node = e.node;
								showMessage(node.data.message, true);
							}}
							onClose={() => {
								showControls.set(false);
							}}
						/>
					{/await}
				{:else}
					<Controls
						on:close={() => {
							showControls.set(false);
						}}
						{models}
						bind:chatFiles
						bind:params
					/>
				{/if}
			</div>
		</Drawer>
	{/if}
{:else}
	<!-- if $showControls -->

	{#if $showControls}
		<PaneResizer
			class="relative flex items-center justify-center group border-l border-gray-50 dark:border-gray-850/30 hover:border-gray-200 dark:hover:border-gray-800  transition z-20"
			id="controls-resizer"
		>
			<div
				class=" absolute -left-1.5 -right-1.5 -top-0 -bottom-0 z-20 cursor-col-resize bg-transparent"
			/>
		</PaneResizer>
	{/if}

	<Pane
		bind:pane
		defaultSize={0}
		onResize={(size) => {
			trace('pane:onResize', { size });
			if ($showControls && pane.isExpanded()) {
				if (size < minSize) {
					trace('pane:onResize:enforceMin', { size, minSize });
					pane.resize(minSize);
				}

				if (size < minSize) {
					localStorage.chatControlsSize = 0;
				} else {
					// save the size in  pixels to localStorage
					const container = document.getElementById('chat-container');
					localStorage.chatControlsSize = Math.floor((size / 100) * container.clientWidth);
					trace('pane:onResize:save', {
						sizePercent: size,
						containerWidth: container?.clientWidth ?? null,
						savedPixels: localStorage.chatControlsSize
					});
				}
			}
		}}
		onCollapse={() => {
			trace('pane:onCollapse');
			showControls.set(false);
		}}
		collapsible={true}
		class=" z-10 bg-white dark:bg-gray-850"
	>
		{#if $showControls}
			<div class="flex max-h-full min-h-full">
				<div
					class="w-full {($showOverview || $showArtifacts || $showEmbeds) && !$showCallOverlay
						? ' '
						: 'px-4 py-3 bg-white dark:shadow-lg dark:bg-gray-850 '} z-40 pointer-events-auto overflow-y-auto scrollbar-hidden"
					id="controls-container"
				>
					{#if $showCallOverlay}
						<div class="w-full h-full flex justify-center">
							<CallOverlay
								bind:files
								{submitPrompt}
								{stopResponse}
								{modelId}
								{chatId}
								{eventTarget}
								on:close={() => {
									showControls.set(false);
								}}
							/>
						</div>
					{:else if $showEmbeds}
						<Embeds overlay={dragged} />
					{:else if $showArtifacts}
						<Artifacts {history} overlay={dragged} />
					{:else if $showOverview}
						{#await import('./Overview.svelte') then { default: Overview }}
							<Overview
								{history}
								onNodeClick={(e) => {
									const node = e.node;
									if (node?.data?.message?.favorite) {
										history.messages[node.data.message.id].favorite = true;
									} else {
										history.messages[node.data.message.id].favorite = null;
									}

									showMessage(node.data.message, true);
								}}
								onClose={() => {
									showControls.set(false);
								}}
							/>
						{/await}
					{:else}
						<Controls
							on:close={() => {
								showControls.set(false);
							}}
							{models}
							bind:chatFiles
							bind:params
						/>
					{/if}
				</div>
			</div>
		{/if}
	</Pane>
{/if}
