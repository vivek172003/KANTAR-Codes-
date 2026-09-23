/*
 * Power BI Custom Visual - AI Insights
 * Refactored: Hierarchical Chunking + Aggregation Architecture
 *
 * Fixes applied on top of chunking refactor:
 *
 * FIX 1 - System prompt is now call-type aware. JSON extraction calls use a
 *          JSON-only system prompt. Only the final insight call uses the
 *          "headings and bullet points" system prompt. This was the primary
 *          cause of Azure returning empty output on chunk calls.
 *
 * FIX 2 - temperature set to 0 for all JSON extraction and aggregation calls.
 *          temperature: 1 on JSON calls causes the model to produce malformed
 *          or empty output. Only the final insight generation uses temperature: 1.
 *
 * FIX 3 - max_completion_tokens raised appropriately per call type:
 *          chunk extraction: 1200, aggregation: 1200, executive: 1500, final: 4000.
 *          800 tokens was not enough for chunk extraction causing truncated/empty JSON.
 *
 * FIX 4 - analyseChunk no longer silently swallows errors. If Azure returns
 *          empty JSON for a chunk after all retries, the chunk is retried once
 *          more with a simplified prompt before falling back to empty arrays.
 *
 * FIX 5 - callAzureOpenAIWithRetry now accepts a callType parameter so the
 *          correct system prompt and temperature are used per call type without
 *          any changes to the outer pipeline logic.
 *
 * FIX 6 - Empty chunk summaries (all arrays empty) are detected and logged
 *          as warnings. The pipeline continues rather than crashing, but the
 *          issue is visible in the browser console for debugging.
 *
 * Nothing else (UI, formatting, settings, caching, event handlers,
 * pivot logic, or chunking structure) has been changed.
 */

"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

import { VisualFormattingSettingsModel } from "./settings";
import * as config from "../config.json";

// PPT export feature - run `npm install pptxgenjs` and add it as a dependency.
// pptxgenjs builds the .pptx entirely client-side (no server round-trip, no win32com,
// no filesystem access needed), which is what a Power BI custom visual's sandbox allows.
import pptxgen from "pptxgenjs";
import { jsPDF } from "jspdf";
import { buildFullPivotedCsv } from "./csvDataService";

// ---------------------------------------------------------------------------
// Chunking architecture - type definitions
// ---------------------------------------------------------------------------

interface MetricEvidence {
    chunkIndex: number;
    rowRange: string;
    entity: string;
    metric: string;
    measure: string;
    score: number | null;
    scoreText: string;
    change: number | null;
    changeText: string;
    significance: string;
    Conversion: number | null;
    ConversionText: string;
    dimensions: { [key: string]: string };
}

interface ChunkSummary {
    chunkIndex: number;
    totalChunks: number;
    rowRange: string;
    keyTrends: string[];
    anomalies: string[];
    risks: string[];
    opportunities: string[];
    kpiObservations: string[];
    businessSignals: string[];
    metricEvidence: MetricEvidence[];
}

interface GroupSummary {
    groupIndex: number;
    chunksCovered: string;
    recurringPatterns: string[];
    mergedFindings: string[];
    prioritisedSignals: string[];
    metricEvidence: MetricEvidence[];
}

// ---------------------------------------------------------------------------
// FIX 5 - Call type enum so callAzureOpenAIWithRetry uses the right
//          system prompt and temperature per pipeline stage.
// ---------------------------------------------------------------------------
type AzureCallType = "json_extraction" | "json_aggregation" | "text_executive" | "text_final";
type UiMessageKind = "info" | "selection" | "error";

interface PersistedInsightCacheEntry {
    response: string;
    savedAt: number;
}

interface PptDownloadResult {
    fileName: string;
    confirmed: boolean;
    method: "server" | "powerbi" | "browser";
    fallbackBlob?: Blob;
    downloadUrl?: string;
}

interface PptServerUploadResponse {
    fileName: string;
    downloadUrl: string;
    expiresAt?: string;
    sizeBytes?: number;
}

interface AnalyticalDimensionDescriptor {
    name: string;
    values: string[];
    isComparison: boolean;
}

interface PdfDownloadResult {
    fileName: string;
    confirmed: boolean;
    method: "powerbi" | "browser";
    fallbackBlob?: Blob;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const CHUNK_ROW_LIMIT        = 300;
const AGGREGATION_BATCH_SIZE = 5;
const CHUNK_TEXT_CHAR_LIMIT  = 16000;   // slightly reduced to keep chunk prompt under 20k total
const MAX_RETRY_ATTEMPTS     = 3;
const RETRY_BASE_DELAY_MS    = 1500;
const MAX_PARALLEL_AZURE_CALLS = 1;

// FIX 3 - token budgets per call type
const TOKENS_CHUNK_EXTRACTION = 1200;
const TOKENS_AGGREGATION      = 1200;
const TOKENS_EXECUTIVE        = 1500;
const TOKENS_FINAL_INSIGHT    = 8000;

// ---------------------------------------------------------------------------
// Visual class
// ---------------------------------------------------------------------------

export class Visual implements IVisual {
    private target: HTMLElement;
    private host: powerbi.extensibility.visual.IVisualHost;
    private downloadService: any;
    private storageV2Service: any;
    private formattingSettings: VisualFormattingSettingsModel = new VisualFormattingSettingsModel();
    private formattingSettingsService: FormattingSettingsService;
    private dataView: powerbi.DataView | undefined;

    // Role-driven metadata from the current Power BI field wells.
    // These keys ensure scores are extracted only from Measure Data fields,
    // while Category Data remains context even when values are numeric.
    private allowedScoreHeaderKeys: { [key: string]: boolean } = {};
    private preferredEntityFieldFromPrompt: string = "";
    private highlightedCategoryValuesFromPrompt: string[] = [];
    private highlightedCategoryFieldsFromPrompt: string[] = [];

    // Section headings parsed directly out of the user's own prompt text (e.g. the
    // labels under an "Output format:" block). Populated by buildFinalOutputFormatInstruction
    // and consumed by formatInsightTextAsHtml so the renderer recognises whatever heading
    // names the user actually asked for, instead of only a fixed static list.
    private lastKnownHeadingsFromPrompt: string[] = [];
    private lastKnownBulletLimitsFromPrompt: { [headingKey: string]: number } = {};
    private lastKnownListStylesFromPrompt: { [headingKey: string]: string } = {};

    private gptResponse: string = "";
    private responseCache: { [key: string]: string } = {};
    private persistedInsightCache: { [key: string]: PersistedInsightCacheEntry } = {};
    private persistentCacheReady: Promise<void> | null = null;
    private persistentCacheAllowed: boolean = false;
    private readonly PERSISTENT_CACHE_STORAGE_KEY: string = "aiInsightsResponseCacheV17";
    private readonly PERSISTENT_CACHE_TTL_MS: number = 24 * 60 * 60 * 1000;
    private readonly MAX_PERSISTED_CACHE_ENTRIES: number = 8;
    private readonly MAX_PERSISTED_CACHE_CHARACTERS: number = 90000;

    private latestRequestId: number = 0;
    private cacheRestoreRequestId: number = 0;
    private activeAzureControllers: Set<AbortController> = new Set<AbortController>();
    private isPanelOpen: boolean = true; // panel remains open when true
    private lastSelectionKey: string = "";
    private inFlightSelectionKey: string = "";
    private hasSelectionChanged: boolean = false;
    private isLoadingInsights: boolean = false;
    private currentProgressMessage: string = "Preparing dataset...";
    private currentUiMessage: string = "Click Show to generate AI insights for the current selection.";
    private currentUiMessageKind: UiMessageKind = "info";
    private lastShowActivationTime: number = 0;
    private lastPptDownloadTime: number = 0;
    private isDownloadingPpt: boolean = false;
    private pptStatusTimeoutId: number | undefined;
    private pendingPptBlob: Blob | null = null;
    private pendingPptFileName: string = "";
    private pendingPptObjectUrl: string = "";
    private pendingExportKind: "ppt" | "pdf" = "ppt";

    private readonly MAX_DATA_ROWS: number = 200;
    private readonly MAX_METRICS: number = 12;
    private readonly MAX_CONTEXT_VALUES_PER_FIELD: number = 250;
    private readonly MAX_DATASET_TEXT_LENGTH: number = 22000;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.downloadService = options.host.downloadService;
        this.storageV2Service = (options.host as any).storageV2Service;
        this.persistentCacheReady = this.initialisePersistentCache();
        this.formattingSettingsService = new FormattingSettingsService();
        this.target.addEventListener("click",     (event: Event) => this.handleVisualAction(event), true);
        this.target.addEventListener("pointerup", (event: Event) => this.handleVisualAction(event), true);
    }

    // -------------------------------------------------------------------------
    // Event handling - unchanged
    // -------------------------------------------------------------------------

    private handleVisualAction(event: Event): void {
        const clickedElement = event.target as HTMLElement | null;
        if (!clickedElement) { return; }

        const toggleButton = clickedElement.closest(".aiInsightsToggle") as HTMLElement | null;
        if (toggleButton) {
            event.preventDefault();
            event.stopPropagation();
            this.isPanelOpen = !this.isPanelOpen;
            this.updateVisual();
            return;
        }

        // Footer actions are checked before Show because all three buttons share
        // the button_Run styling class. The old generic selector caused Download
        // and Copy clicks to launch the Azure analysis pipeline as well.
        const downloadPptButton = clickedElement.closest(".button_DownloadPpt") as HTMLButtonElement | null;
        if (downloadPptButton) {
            event.preventDefault();
            event.stopPropagation();
            const now = Date.now();
            if (now - this.lastPptDownloadTime < 700) { return; }
            this.lastPptDownloadTime = now;

            // Power BI Desktop hosts custom visuals in a sandbox. PPT generation
            // is asynchronous, so the browser's original user activation can expire
            // before the file is ready. The first click prepares the PPT; when the
            // host does not confirm a download, the next direct click saves it.
            if (this.pendingPptBlob && this.pendingPptFileName) {
                void this.savePendingPptFromUserGesture();
            } else {
                void this.downloadInsightsAsPpt();
            }
            return;
        }

        const copyButton = clickedElement.closest(".button_Copy") as HTMLButtonElement | null;
        if (copyButton) {
            event.preventDefault();
            event.stopPropagation();

            const text = String(this.gptResponse || "").trim();
            if (!text || !this.isValidAiResponse(text)) {
                this.setCopyButtonFeedback(copyButton, "No insights", false);
                return;
            }

            const doFallbackCopy = function(copyText: string): boolean {
                try {
                    const ta = document.createElement("textarea");
                    ta.style.position = "fixed";
                    ta.style.left = "-9999px";
                    ta.value = copyText;
                    document.body.appendChild(ta);
                    ta.select();
                    const copied = document.execCommand("copy");
                    document.body.removeChild(ta);
                    return copied;
                } catch (error) {
                    return false;
                }
            };

            const finishCopy = (success: boolean): void => {
                this.setCopyButtonFeedback(copyButton, success ? "Copied" : "Copy failed", success);
            };

            if ((navigator as any).clipboard && (navigator as any).clipboard.writeText) {
                (navigator as any).clipboard.writeText(text)
                    .then(function() { finishCopy(true); })
                    .catch(function() { finishCopy(doFallbackCopy(text)); });
            } else {
                finishCopy(doFallbackCopy(text));
            }
            return;
        }

        const showButton = clickedElement.closest(".button_ShowInsights") as HTMLButtonElement | null;
        if (showButton) {
            event.preventDefault();
            event.stopPropagation();
            const now = Date.now();
            if (now - this.lastShowActivationTime < 350) { return; }
            this.lastShowActivationTime = now;
            void this.submitPrompt();
            return;
        }
    }

    private setCopyButtonFeedback(button: HTMLButtonElement, label: string, success: boolean): void {
        button.textContent = label;
        button.classList.toggle("copied", success);
        button.setAttribute("aria-label", label);
        window.setTimeout(function() {
            button.textContent = "Copy";
            button.classList.remove("copied");
            button.setAttribute("aria-label", "Copy AI insights");
        }, 1300);
    }

    // -------------------------------------------------------------------------
    // Power BI update lifecycle - unchanged
    // -------------------------------------------------------------------------

    public update(options: VisualUpdateOptions): void {
        try {
            this.dataView = options.dataViews && options.dataViews.length > 0
                ? options.dataViews[0]
                : undefined;

            if (options.dataViews && options.dataViews.length > 0) {
                this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    options.dataViews
                );
            }

            if (!this.dataView || !this.dataView.categorical) {
                this.cancelActiveAnalysis();
                this.gptResponse = "";
                this.currentUiMessage = "No data available. Please check your data source.";
                this.currentUiMessageKind = "error";
                this.lastSelectionKey = "";
                this.updateVisual();
                return;
            }

            const categories = this.dataView.categorical.categories || [];
            const values = this.dataView.categorical.values;

            if (categories.length === 0 || !values || values.length === 0) {
                this.cancelActiveAnalysis();
                this.gptResponse = "";
                this.currentUiMessage = "Please provide both Category Data and Measure Data.";
                this.currentUiMessageKind = "error";
                this.lastSelectionKey = "";
                this.updateVisual();
                return;
            }

            const promptInfo = this.buildPromptFromData();
            const nextSelectionKey = promptInfo.selectionKey;
            const isFirstSelection = !this.lastSelectionKey;
            const selectionChanged = Boolean(
                this.lastSelectionKey && this.lastSelectionKey !== nextSelectionKey
            );

            if (selectionChanged) {
                // Cancel the UI ownership of the previous request. The network promise
                // may finish, but requestId checks prevent stale data from being shown.
                this.cancelActiveAnalysis();
                this.clearPendingPptDownload();
                this.gptResponse = "";
                this.hasSelectionChanged = true;
                this.currentUiMessage = "Selection changed. Click Show to generate insights for this selection.";
                this.currentUiMessageKind = "selection";
            }

            this.lastSelectionKey = nextSelectionKey;
            const memoryCachedResponse = this.getMemoryCachedResponse(nextSelectionKey);

            if (memoryCachedResponse) {
                this.gptResponse = memoryCachedResponse;
                this.currentUiMessage = "";
                this.hasSelectionChanged = false;
                this.isLoadingInsights = false;
                this.inFlightSelectionKey = "";
            } else if (isFirstSelection) {
                this.gptResponse = "";
                this.hasSelectionChanged = false;
                this.currentUiMessage = "Click Show to generate AI insights for the current selection.";
                this.currentUiMessageKind = "info";
            } else if (!selectionChanged && !this.isLoadingInsights && !this.isValidAiResponse(this.gptResponse)) {
                this.gptResponse = "";
                this.currentUiMessage = this.hasSelectionChanged
                    ? "Selection changed. Click Show to generate insights for this selection."
                    : "Click Show to generate AI insights for the current selection.";
                this.currentUiMessageKind = this.hasSelectionChanged ? "selection" : "info";
            }

            this.updateVisual();

            // This lookup does not call Azure. It only checks Power BI's local-storage
            // service and falls back to the in-memory cache when storage is unavailable.
            if (!memoryCachedResponse && !this.isLoadingInsights) {
                void this.restoreCachedResponseForSelection(nextSelectionKey);
            }
        } catch (error) {
            console.error("Error in update method:", error);
            this.cancelActiveAnalysis();
            this.gptResponse = "";
            this.currentUiMessage = this.getCleanErrorMessage(error);
            this.currentUiMessageKind = "error";
            this.updateVisual();
        }
    }

    // -------------------------------------------------------------------------
    // UI rendering
    // -------------------------------------------------------------------------

    // public updateVisual(): void {
    //     this.target.replaceChildren();

    //     const root = document.createElement("div");
public updateVisual(): void {
 
    // If the visual already exists, don't destroy
    // and recreate the entire DOM.
    const existingResponse =
        this.target.querySelector(
            "#gptResponse"
        ) as HTMLElement | null;
 
    if (existingResponse) {
 
        this.renderCurrentResponse();
 
        this.updateShowButtonState(
            this.isLoadingInsights
        );
 
        this.setDownloadPptButtonState(
            this.isDownloadingPpt
        );
 
        return;
    }
 
    // First render only.
    this.target.replaceChildren();
 
    const root = document.createElement("div");
 

        root.className = "popup_bg ai-root";

        const wrapper = document.createElement("div");
        wrapper.className = "ai-wrapper";

        const panel = document.createElement("div");
        panel.className = "popupDV ai-panel-shell";

        const header = document.createElement("div");
        header.className = "Popup_head ai-panel-header";

        const buttonWrap = document.createElement("div");
        buttonWrap.className = "popup_Button";

        const showButton = document.createElement("button");
        showButton.type = "button";
        showButton.className = "button_Run button_ShowInsights";
        showButton.setAttribute("aria-label", "Show AI insights");
        showButton.textContent = "Show";
        buttonWrap.appendChild(showButton);

        const title = document.createElement("h3");
        title.className = "ai-panel-title";
        title.textContent = "AI Powered Summary";

        header.appendChild(buttonWrap);
        header.appendChild(title);

        const contentArea = document.createElement("div");
        contentArea.className = "popup_ContentArea ai-content-area";

        const response = document.createElement("div");
        response.id = "gptResponse";
        response.className = "ai-response";
        response.setAttribute("role", "region");
        response.setAttribute("aria-live", "polite");
        contentArea.appendChild(response);

        const footer = document.createElement("div");
        footer.className = "ai-panel-footer";

        const footerButtons = document.createElement("div");
        footerButtons.className = "ai-footer-actions";

        const copyButton = document.createElement("button");
        copyButton.type = "button";
        copyButton.className = "button_Run button_Copy";
        copyButton.setAttribute("aria-label", "Copy AI insights");
        copyButton.textContent = "Copy";

        // const downloadButton = document.createElement("button");
        // downloadButton.type = "button";
        // downloadButton.className = "button_Run button_DownloadPpt";
        // downloadButton.setAttribute("aria-label", "Download AI insights as PowerPoint");
        // downloadButton.setAttribute(
        //     "title",
        //     "Downloads to the default Power BI or browser download location"
        // );
        // this.renderDownloadButtonContent(downloadButton, false);

        footerButtons.appendChild(copyButton);
        // footerButtons.appendChild(downloadButton);
        footer.appendChild(footerButtons);

        panel.appendChild(header);
        panel.appendChild(contentArea);
        panel.appendChild(footer);
        wrapper.appendChild(panel);
        root.appendChild(wrapper);
        this.target.appendChild(root);

        // Power BI can rebuild the DOM during resize/filter updates. Render from the
        // stored state so the loading spinner and progress text never disappear.
        this.renderCurrentResponse();
        this.updateShowButtonState(this.isLoadingInsights);
        this.setDownloadPptButtonState(this.isDownloadingPpt);
    }

    private createDownloadIcon(): SVGSVGElement {
        const namespace = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(namespace, "svg");
        svg.setAttribute("class", "ai-download-icon");
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("aria-hidden", "true");
        svg.setAttribute("focusable", "false");
        svg.setAttribute("fill", "none");
        svg.setAttribute("stroke", "currentColor");
        svg.setAttribute("stroke-width", "2");
        svg.setAttribute("stroke-linecap", "round");
        svg.setAttribute("stroke-linejoin", "round");

        const arrow = document.createElementNS(namespace, "path");
        arrow.setAttribute("d", "M12 3v11m0 0 4-4m-4 4-4-4");
        const tray = document.createElementNS(namespace, "path");
        tray.setAttribute("d", "M5 17v3h14v-3");
        svg.appendChild(arrow);
        svg.appendChild(tray);
        return svg;
    }

    private renderDownloadButtonContent(button: HTMLButtonElement, isBusy: boolean): void {
        button.replaceChildren();
        if (isBusy) {
            const spinner = document.createElement("span");
            spinner.className = "ai-button-spinner";
            spinner.setAttribute("aria-hidden", "true");
            button.appendChild(spinner);
        } else {
            button.appendChild(this.createDownloadIcon());
        }

        const label = document.createElement("span");
        label.className = "ai-download-label";
        label.textContent = isBusy
            ? "Preparing PPT..."
            : (this.pendingPptBlob && this.pendingPptFileName
                ? (this.pendingExportKind === "pdf" ? "Save PDF" : "Save PPT")
                : "Download PPT");
        button.appendChild(label);
    }

    private renderCurrentResponse(): void {
        if (this.isLoadingInsights) {
            this.renderMessageCard(this.currentProgressMessage || "Preparing dataset...", "loading");
            return;
        }

        if (this.currentUiMessage) {
            this.renderMessageCard(this.currentUiMessage, this.currentUiMessageKind);
            return;
        }

        if (this.gptResponse && this.isValidAiResponse(this.gptResponse)) {
            this.updateGptResponse(this.gptResponse);
            return;
        }

        this.renderMessageCard(
            "Click Show to generate AI insights for the current selection.",
            "info"
        );
    }

    private renderMessageCard(message: string, kind: UiMessageKind | "loading"): void {
        const responseDisplay = this.target.querySelector("#gptResponse") as HTMLElement | null;
        if (!responseDisplay) { return; }

        responseDisplay.replaceChildren();
        responseDisplay.setAttribute("aria-busy", kind === "loading" ? "true" : "false");

        const card = document.createElement("div");
        card.className = "ai-message-card" +
            (kind === "loading" ? " ai-message-card--loading" : "") +
            (kind === "error" ? " ai-message-card--error" : "");
        card.setAttribute("role", kind === "error" ? "alert" : "status");
        card.setAttribute("aria-live", "polite");

        if (kind === "loading") {
            const spinner = document.createElement("span");
            spinner.className = "ai-loading-spinner";
            spinner.setAttribute("aria-hidden", "true");
            card.appendChild(spinner);
        } else {
            const icon = document.createElement("span");
            icon.className = "ai-message-icon";
            icon.setAttribute("aria-hidden", "true");
            icon.textContent = kind === "selection" ? "?" : kind === "error" ? "!" : "i";
            card.appendChild(icon);
        }

        const copy = document.createElement("div");
        copy.className = "ai-message-copy";

        const title = document.createElement("div");
        title.className = "ai-message-title";
        title.textContent = kind === "loading"
            ? "Generating insights"
            : kind === "selection"
                ? "Selection changed"
                : kind === "error"
                    ? "Unable to continue"
                    : "Ready";

        const primary = document.createElement("p");
        primary.className = "ai-message-primary";
        primary.textContent = String(message || "").trim();

        if (kind === "loading") {
            const dots = document.createElement("span");
            dots.className = "ai-loading-dots";
            dots.setAttribute("aria-hidden", "true");
            for (let i = 0; i < 3; i++) {
                dots.appendChild(document.createElement("span"));
            }
            primary.appendChild(dots);
        }

        copy.appendChild(title);
        copy.appendChild(primary);
        card.appendChild(copy);

        if (kind === "loading") {
            const track = document.createElement("div");
            track.className = "ai-progress-track";
            track.setAttribute("aria-hidden", "true");
            track.appendChild(document.createElement("span"));
            card.appendChild(track);
        }

        responseDisplay.appendChild(card);
    }

    // private cancelActiveAnalysis(): void {
    //     this.latestRequestId += 1;
    //     this.cacheRestoreRequestId += 1;
    //     this.isLoadingInsights = false;
    //     this.inFlightSelectionKey = "";
    //     this.currentProgressMessage = "Preparing dataset...";
    //     this.updateShowButtonState(false);
    // }

    // -------------------------------------------------------------------------
    // submitPrompt - entry point
    // -------------------------------------------------------------------------

    private cancelActiveAnalysis(): void {
 
    this.latestRequestId += 1;
    this.cacheRestoreRequestId += 1;
 
    // Stop Azure calls belonging to the old selection.
    this.activeAzureControllers.forEach(
        function(controller: AbortController) {
            try {
                controller.abort();
            } catch (error) {
                // Ignore abort errors.
            }
        }
    );
 
    this.activeAzureControllers.clear();
 
    this.isLoadingInsights = false;
    this.inFlightSelectionKey = "";
    this.currentProgressMessage = "Preparing dataset...";
    this.updateShowButtonState(false);
}
 

    private async submitPrompt(): Promise<void> {
        try {
            if (!this.dataView || !this.dataView.categorical) {
                this.gptResponse = "";
                this.currentUiMessage = "No data available. Please check your data source.";
                this.currentUiMessageKind = "error";
                this.renderCurrentResponse();
                return;
            }

            const promptInfo = this.buildPromptFromData();
            const selectionKey = promptInfo.selectionKey;
            this.lastSelectionKey = selectionKey;

            if (this.isLoadingInsights && this.inFlightSelectionKey === selectionKey) {
                this.renderCurrentResponse();
                return;
            }

            const cachedResponse = await this.getCachedResponse(selectionKey);
            if (selectionKey !== this.lastSelectionKey) { return; }

            if (cachedResponse) {
                this.gptResponse = cachedResponse;
                this.currentUiMessage = "";
                this.hasSelectionChanged = false;
                this.isLoadingInsights = false;
                this.inFlightSelectionKey = "";
                this.renderCurrentResponse();
                console.log("[AI Insights Cache] Cache hit; Azure API was not called.");
                return;
            }

            if (this.isLoadingInsights && this.inFlightSelectionKey !== selectionKey) {
                this.cancelActiveAnalysis();
            }

            this.hasSelectionChanged = false;
            await this.runChunkedAnalysis(
                promptInfo.basePrompt,
                selectionKey,
                Boolean(promptInfo.basePromptIsDefault)
            );
        } catch (error) {
            console.error("Error in submitPrompt:", error);
            this.cancelActiveAnalysis();
            this.gptResponse = "";
            this.currentUiMessage = this.getCleanErrorMessage(error);
            this.currentUiMessageKind = "error";
            this.renderCurrentResponse();
        }
    }

    private async initialisePersistentCache(): Promise<void> {
        if (!this.storageV2Service || typeof this.storageV2Service.status !== "function") {
            return;
        }

        try {
            const status = await this.storageV2Service.status();
            this.persistentCacheAllowed = status === powerbi.PrivilegeStatus.Allowed;
            if (!this.persistentCacheAllowed || typeof this.storageV2Service.get !== "function") {
                return;
            }

            try {
                const rawValue = await this.storageV2Service.get(this.PERSISTENT_CACHE_STORAGE_KEY);
                const parsed = rawValue ? JSON.parse(String(rawValue)) : {};
                if (parsed && typeof parsed === "object") {
                    this.persistedInsightCache = parsed as { [key: string]: PersistedInsightCacheEntry };
                    this.removeExpiredPersistentCacheEntries();
                }
            } catch (readError) {
                console.info("[AI Insights Cache] No persisted cache was loaded.", readError);
            }
        } catch (statusError) {
            this.persistentCacheAllowed = false;
            console.info("[AI Insights Cache] LocalStorage privilege is unavailable; using memory cache only.", statusError);
        }
    }

    private getMemoryCachedResponse(selectionKey: string): string {
        const cached = String(this.responseCache[selectionKey] || "").trim();
        if (cached && this.isValidAiResponse(cached)) {
            return cached;
        }
        if (this.responseCache[selectionKey]) {
            delete this.responseCache[selectionKey];
        }
        return "";
    }

    private async getCachedResponse(selectionKey: string): Promise<string> {
        const memoryCached = this.getMemoryCachedResponse(selectionKey);
        if (memoryCached) { return memoryCached; }

        if (this.persistentCacheReady) {
            await this.persistentCacheReady;
        }

        const entry = this.persistedInsightCache[selectionKey];
        if (!entry) { return ""; }

        const expired = !entry.savedAt || Date.now() - entry.savedAt > this.PERSISTENT_CACHE_TTL_MS;
        const response = String(entry.response || "").trim();
        if (expired || !this.isValidAiResponse(response)) {
            delete this.persistedInsightCache[selectionKey];
            void this.persistInsightCache();
            return "";
        }

        this.responseCache[selectionKey] = response;
        return response;
    }

    private async restoreCachedResponseForSelection(selectionKey: string): Promise<void> {
        const restoreRequestId = ++this.cacheRestoreRequestId;
        const cachedResponse = await this.getCachedResponse(selectionKey);

        if (
            restoreRequestId !== this.cacheRestoreRequestId ||
            selectionKey !== this.lastSelectionKey ||
            this.isLoadingInsights ||
            !cachedResponse
        ) {
            return;
        }

        this.gptResponse = cachedResponse;
        this.currentUiMessage = "";
        this.hasSelectionChanged = false;
        this.inFlightSelectionKey = "";
        this.renderCurrentResponse();
        console.log("[AI Insights Cache] Restored saved insights; Azure API was not called.");
    }

    private storeCachedResponse(selectionKey: string, response: string): void {
        const cleaned = String(response || "").trim();
        if (!selectionKey || !this.isValidAiResponse(cleaned)) { return; }

        this.responseCache[selectionKey] = cleaned;
        this.persistedInsightCache[selectionKey] = {
            response: cleaned,
            savedAt: Date.now()
        };
        this.trimPersistentCacheEntries();
        void this.persistInsightCache();
    }

    private removeExpiredPersistentCacheEntries(): void {
        const now = Date.now();
        const keys = Object.keys(this.persistedInsightCache);
        for (let i = 0; i < keys.length; i++) {
            const entry = this.persistedInsightCache[keys[i]];
            if (
                !entry ||
                !entry.savedAt ||
                now - entry.savedAt > this.PERSISTENT_CACHE_TTL_MS ||
                !this.isValidAiResponse(String(entry.response || ""))
            ) {
                delete this.persistedInsightCache[keys[i]];
            }
        }
        this.trimPersistentCacheEntries();
    }

    private trimPersistentCacheEntries(): void {
        let keys = Object.keys(this.persistedInsightCache).sort((a: string, b: string) => {
            return (this.persistedInsightCache[b].savedAt || 0) - (this.persistedInsightCache[a].savedAt || 0);
        });

        for (let i = this.MAX_PERSISTED_CACHE_ENTRIES; i < keys.length; i++) {
            delete this.persistedInsightCache[keys[i]];
        }

        keys = Object.keys(this.persistedInsightCache).sort((a: string, b: string) => {
            return (this.persistedInsightCache[b].savedAt || 0) - (this.persistedInsightCache[a].savedAt || 0);
        });

        // Power BI local storage is deliberately small. Keep newest entries and
        // discard the oldest until the JSON payload is safely below the limit.
        while (
            keys.length > 1 &&
            JSON.stringify(this.persistedInsightCache).length > this.MAX_PERSISTED_CACHE_CHARACTERS
        ) {
            const oldestKey = keys.pop();
            if (oldestKey) { delete this.persistedInsightCache[oldestKey]; }
        }
    }

    private async persistInsightCache(): Promise<void> {
        if (this.persistentCacheReady) {
            await this.persistentCacheReady;
        }
        if (
            !this.persistentCacheAllowed ||
            !this.storageV2Service ||
            typeof this.storageV2Service.set !== "function"
        ) {
            return;
        }

        try {
            await this.storageV2Service.set(
                this.PERSISTENT_CACHE_STORAGE_KEY,
                JSON.stringify(this.persistedInsightCache)
            );
        } catch (writeError) {
            console.info("[AI Insights Cache] Unable to persist the cache; memory cache remains active.", writeError);
        }
    }

    // -------------------------------------------------------------------------
    // CHUNKING LAYER
    // -------------------------------------------------------------------------

    private splitCsvIntoChunks(fullCsvText: string): string[] {
        const lines = fullCsvText
            .split("\n")
            .map(function(l) { return l.replace(/\s+$/, ""); })
            .filter(function(l) { return l.length > 0; });

        if (lines.length === 0) { return []; }

        const noteLines: string[]  = [];
        let   headerLineIndex      = -1;

        for (let i = 0; i < lines.length; i++) {
            if (lines[i].indexOf("NOTE:") === 0) {
                noteLines.push(lines[i]);
            } else {
                headerLineIndex = i;
                break;
            }
        }

        if (headerLineIndex === -1) { return [fullCsvText]; }

        const headerLine  = lines[headerLineIndex];
        const dataLines   = lines.slice(headerLineIndex + 1);

        if (dataLines.length === 0) { return [fullCsvText]; }

        const chunks: string[]  = [];
        const notesPrefix       = noteLines.length > 0 ? noteLines.join("\n") + "\n" : "";

        for (let start = 0; start < dataLines.length; start += CHUNK_ROW_LIMIT) {
            const chunkRows = dataLines.slice(start, start + CHUNK_ROW_LIMIT);
            let   chunkText = notesPrefix + headerLine + "\n" + chunkRows.join("\n");

            if (chunkText.length > CHUNK_TEXT_CHAR_LIMIT) {
                chunkText = chunkText.substring(0, CHUNK_TEXT_CHAR_LIMIT) + "\n... chunk trimmed";
            }

            chunks.push(chunkText);
        }

        return chunks;
    }

    // -------------------------------------------------------------------------
    // CHUNK ANALYSIS LAYER
    // FIX 1 + FIX 2 + FIX 3 + FIX 4 applied here.
    // -------------------------------------------------------------------------

    private async analyseChunk(
        chunkText: string,
        chunkIndex: number,
        totalChunks: number,
        firstDataRow: number
    ): Promise<ChunkSummary> {

        const nonNoteLines = chunkText.split("\n").filter(function(l) {
            return l.trim().length > 0 && l.indexOf("NOTE:") !== 0;
        });
        // subtract 1 for header line
        const lastDataRow = firstDataRow + Math.max(0, nonNoteLines.length - 2);

        // FIX 1 - extraction prompt explicitly forbids prose; uses json_extraction call type
        // FIX 2 - temperature 0 enforced via callType
        // FIX 3 - TOKENS_CHUNK_EXTRACTION = 1200
        const extractionPrompt =
            "Extract structured data observations from the CSV chunk below.\n" +
            "You MUST return a single valid JSON object - nothing else.\n" +
            "No markdown code fences. No prose. No explanation before or after the JSON.\n" +
            "JSON schema (arrays must contain short strings only; metricEvidence is filled deterministically by code):\n" +
            "{\n" +
            "  \"keyTrends\": [],\n" +
            "  \"anomalies\": [],\n" +
            "  \"risks\": [],\n" +
            "  \"opportunities\": [],\n" +
            "  \"kpiObservations\": [],\n" +
            "  \"businessSignals\": []\n" +
            "}\n" +
            "Limits: keyTrends ?10, anomalies ?5, risks ?5, opportunities ?5, kpiObservations ?12, businessSignals ?8.\n" +
            "Important: preserve exact numeric values from all measure columns when available; never replace visible numbers with 'not visible'.\n" +
            "Important: use all significance columns/values dynamically. Treat Significantly Higher/Lower as significant changes.\n" +
            "Important: when describing entities, use the actual entity/category value from the data, not row numbers. Use row numbers only if no entity/category value exists.\n" +
            "Do not depend on fixed column names; infer dimensions, metrics, scores/measures, and significance from headers and values.\n" +
            "Keep each string under 28 words.\n\n" +
            "Chunk " + (chunkIndex + 1) + " of " + totalChunks +
            " (rows " + firstDataRow + "-" + lastDataRow + "):\n\n" +
            chunkText +
            "\n\nJSON:";

        let rawJson = "";

        try {
            rawJson = await this.callAzureOpenAIWithRetry(
                extractionPrompt,
                TOKENS_CHUNK_EXTRACTION,
                1,
                "json_extraction"
            );
        } catch (firstError) {
            // FIX 4 - retry once with a simpler prompt before giving up
            console.warn(
                "[ChunkAnalysis] Chunk " + (chunkIndex + 1) + " failed on first attempt. " +
                "Retrying with simplified prompt.", firstError
            );

            const simplifiedPrompt =
                "Return ONLY this JSON object filled with short observations from the data below. " +
                "No other text.\n" +
                "{\"keyTrends\":[],\"anomalies\":[],\"risks\":[],\"opportunities\":[]," +
                "\"kpiObservations\":[],\"businessSignals\":[]}\n\n" +
                chunkText.substring(0, 6000) +
                "\n\nJSON:";

            try {
                rawJson = await this.callAzureOpenAIWithRetry(
                    simplifiedPrompt,
                    TOKENS_CHUNK_EXTRACTION,
                    1,
                    "json_extraction"
                );
            } catch (secondError) {
                console.error(
                    "[ChunkAnalysis] Chunk " + (chunkIndex + 1) +
                    " failed after simplified retry. Using empty summary.", secondError
                );
                rawJson = "{}";
            }
        }

        let parsed: Partial<ChunkSummary> = {};

        try {
            // Strip any accidental fences a misbehaving model version might emit
            const cleaned = rawJson
                .replace(/```json/g, "")
                .replace(/```/g, "")
                .trim();

            // Find the first { and last } to handle any leading/trailing garbage text
            const firstBrace = cleaned.indexOf("{");
            const lastBrace  = cleaned.lastIndexOf("}");

            if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
                parsed = JSON.parse(cleaned.substring(firstBrace, lastBrace + 1)) as Partial<ChunkSummary>;
            }
        } catch (parseError) {
            console.warn(
                "[ChunkAnalysis] JSON parse failed for chunk " + (chunkIndex + 1) +
                ". Raw response was: " + rawJson.substring(0, 300), parseError
            );
        }

        const ensureArray = function(value: unknown): string[] {
            if (Array.isArray(value)) {
                return (value as unknown[])
                    .map(function(item) { return String(item); })
                    .filter(function(s) { return s.trim().length > 0; });
            }
            return [];
        };

        const rowRangeText = "rows " + firstDataRow + "-" + lastDataRow;


        const deterministicEvidence = this.extractMetricEvidenceFromChunk(
            chunkText,
            chunkIndex,
            totalChunks,
            rowRangeText
        );

        const summary: ChunkSummary = {
            chunkIndex,
            totalChunks,
            rowRange:         rowRangeText,
            keyTrends:        ensureArray(parsed.keyTrends),
            anomalies:        ensureArray(parsed.anomalies),
            risks:            ensureArray(parsed.risks),
            opportunities:    ensureArray(parsed.opportunities),
            kpiObservations:  ensureArray(parsed.kpiObservations),
            businessSignals:  ensureArray(parsed.businessSignals),
            metricEvidence:   deterministicEvidence
        };

        // FIX 6 - warn when a chunk produced no data at all
        const totalItems =
            summary.keyTrends.length +
            summary.anomalies.length +
            summary.risks.length +
            summary.opportunities.length +
            summary.kpiObservations.length +
            summary.businessSignals.length +
            summary.metricEvidence.length;

        if (totalItems === 0) {
            console.warn(
                "[ChunkAnalysis] Chunk " + (chunkIndex + 1) +
                " produced an empty summary. Check the raw response above."
            );
        }

        return summary;
    }

    // -------------------------------------------------------------------------
    // MEMORY LAYER
    // -------------------------------------------------------------------------

    private storeChunkSummary(memory: ChunkSummary[], summary: ChunkSummary): void {
        memory.push(summary);
        console.log(
            "[MemoryLayer] Stored chunk summary " +
            (summary.chunkIndex + 1) + "/" + summary.totalChunks +
            " (" + summary.rowRange + ")",
            summary
        );
    }

    // -------------------------------------------------------------------------
    // AGGREGATION LAYER
    // FIX 1 + FIX 2 + FIX 3 applied here.
    // -------------------------------------------------------------------------

    private serialiseChunkBatch(batch: ChunkSummary[]): string {
        return batch.map((s) => {
            return (
                "Chunk " + (s.chunkIndex + 1) + " (" + s.rowRange + "):\n" +
                "  Key Trends:        " + s.keyTrends.join(" | ") + "\n" +
                "  Anomalies:         " + s.anomalies.join(" | ") + "\n" +
                "  Risks:             " + s.risks.join(" | ") + "\n" +
                "  Opportunities:     " + s.opportunities.join(" | ") + "\n" +
                "  KPI Observations:  " + s.kpiObservations.join(" | ") + "\n" +
                "  Business Signals:  " + s.businessSignals.join(" | ") + "\n" +
                "  Metric Evidence:   " + this.buildCompactEvidenceLine(s.metricEvidence, 12)
            );
        }).join("\n\n");
    }

    private async aggregateChunkBatch(batch: ChunkSummary[], groupIndex: number): Promise<GroupSummary> {
        const firstChunk = batch[0].chunkIndex + 1;
        const lastChunk  = batch[batch.length - 1].chunkIndex + 1;

        // FIX 1 - aggregation also uses json_aggregation call type (temperature 0, JSON system prompt)
        // FIX 3 - TOKENS_AGGREGATION = 1200
        const aggregationPrompt =
            "Merge the chunk summaries below into one GroupSummary JSON object.\n" +
            "You MUST return a single valid JSON object - nothing else.\n" +
            "No markdown fences. No prose before or after.\n" +
            "JSON schema:\n" +
            "{\n" +
            "  \"recurringPatterns\": [],\n" +
            "  \"mergedFindings\": [],\n" +
            "  \"prioritisedSignals\": []\n" +
            "}\n" +
            "Rules:\n" +
            "  - recurringPatterns ?10: patterns seen in more than one chunk.\n" +
            "  - mergedFindings ?15: de-duplicated findings across all chunks.\n" +
            "  - prioritisedSignals ?8: most significant signals by business impact.\n" +
            "  - Keep each string under 25 words.\n\n" +
            "Chunk summaries (chunks " + firstChunk + "-" + lastChunk + "):\n\n" +
            this.serialiseChunkBatch(batch) +
            "\n\nJSON:";

        let rawJson = "";

        try {
            rawJson = await this.callAzureOpenAIWithRetry(
                aggregationPrompt,
                TOKENS_AGGREGATION,
                1,
                "json_aggregation"
            );
        } catch (error) {
            console.warn(
                "[AggregationLayer] Group " + (groupIndex + 1) +
                " aggregation failed. Using empty group summary.", error
            );
            rawJson = "{}";
        }

        let parsed: Partial<GroupSummary> = {};

        try {
            const cleaned    = rawJson.replace(/```json/g, "").replace(/```/g, "").trim();
            const firstBrace = cleaned.indexOf("{");
            const lastBrace  = cleaned.lastIndexOf("}");

            if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
                parsed = JSON.parse(cleaned.substring(firstBrace, lastBrace + 1)) as Partial<GroupSummary>;
            }
        } catch (parseError) {
            console.warn(
                "[AggregationLayer] JSON parse failed for group " + (groupIndex + 1) +
                ". Raw: " + rawJson.substring(0, 200), parseError
            );
        }

        const ensureArray = function(value: unknown): string[] {
            if (Array.isArray(value)) {
                return (value as unknown[])
                    .map(function(item) { return String(item); })
                    .filter(function(s) { return s.trim().length > 0; });
            }
            return [];
        };

        const mergedMetricEvidence: MetricEvidence[] = [];
        for (let i = 0; i < batch.length; i++) {
            mergedMetricEvidence.push.apply(mergedMetricEvidence, batch[i].metricEvidence || []);
        }

        return {
            groupIndex,
            chunksCovered:      "chunks " + firstChunk + "-" + lastChunk,
            recurringPatterns:  ensureArray(parsed.recurringPatterns),
            mergedFindings:     ensureArray(parsed.mergedFindings),
            prioritisedSignals: ensureArray(parsed.prioritisedSignals),
            metricEvidence:     mergedMetricEvidence
        };
    }

    // private async runAggregationLayer(memory: ChunkSummary[]): Promise<GroupSummary[]> {
    //     const groupSummaries: GroupSummary[] = [];

    //     for (let start = 0; start < memory.length; start += AGGREGATION_BATCH_SIZE) {
    //         const batch      = memory.slice(start, start + AGGREGATION_BATCH_SIZE);
    //         const groupIndex = Math.floor(start / AGGREGATION_BATCH_SIZE);
    //         const totalGroups = Math.ceil(memory.length / AGGREGATION_BATCH_SIZE);

    //         this.updateProgressMessage(
    //             // "Aggregating findings (group " + (groupIndex + 1) + " of " + totalGroups + ")..."
    //             //  "Aggregating findings..."
    //             "Summarizing data..."
    //         );

    //         const groupSummary = await this.aggregateChunkBatch(batch, groupIndex);
    //         groupSummaries.push(groupSummary);

    //         console.log("[AggregationLayer] Group " + (groupIndex + 1) + " complete.", groupSummary);
    //     }

    //     return groupSummaries;
    // }
private async runAggregationLayer(memory: ChunkSummary[]): Promise<GroupSummary[]> {
 
    const jobs: Array<{
        batch: ChunkSummary[];
        groupIndex: number;
    }> = [];
 
    // Build the exact same aggregation batches as before.
    for (
        let start = 0;
        start < memory.length;
        start += AGGREGATION_BATCH_SIZE
    ) {
        jobs.push({
            batch: memory.slice(
                start,
                start + AGGREGATION_BATCH_SIZE
            ),
            groupIndex: Math.floor(
                start / AGGREGATION_BATCH_SIZE
            )
        });
    }
 
    this.updateProgressMessage("Summarizing data...");
 
    // Keep output in exactly the same group order.
    const groupSummaries: GroupSummary[] =
        new Array(jobs.length);
 
    let nextJob = 0;
 
    const worker = async (): Promise<void> => {
 
        while (true) {
 
            const jobIndex = nextJob++;
 
            if (jobIndex >= jobs.length) {
                return;
            }
 
            const job = jobs[jobIndex];
 
            const groupSummary =
                await this.aggregateChunkBatch(
                    job.batch,
                    job.groupIndex
                );
 
            // Store in original position,
            // regardless of which request finishes first.
            groupSummaries[jobIndex] = groupSummary;
 
            console.log(
                "[AggregationLayer] Group " +
                (job.groupIndex + 1) +
                " complete.",
                groupSummary
            );
        }
    };
 
    const workerCount = Math.min(
        MAX_PARALLEL_AZURE_CALLS,
        jobs.length
    );
 
    await Promise.all(
        Array.from(
            { length: workerCount },
            () => worker()
        )
    );
 
    return groupSummaries;
}
 

    // -------------------------------------------------------------------------
    // EXECUTIVE SUMMARY LAYER (Level 3)
    // FIX 1 + FIX 3 applied here - text_executive call type, 1500 tokens.
    // -------------------------------------------------------------------------

    private async buildExecutiveSummary(
        groupSummaries: GroupSummary[],
        totalRows: number,
        totalChunks: number
    ): Promise<string> {
        const allEvidence = this.collectMetricEvidence(groupSummaries);
        const evidenceSummary = this.buildMetricEvidenceSummary(allEvidence, 18, 5);

        const groupText = groupSummaries.map((g) => {
            return (
                "Group " + (g.groupIndex + 1) + " (" + g.chunksCovered + "):\n" +
                "  Recurring Patterns:   " + g.recurringPatterns.join(" | ") + "\n" +
                "  Merged Findings:      " + g.mergedFindings.join(" | ") + "\n" +
                "  Prioritised Signals:  " + g.prioritisedSignals.join(" | ") + "\n" +
                "  Evidence:             " + this.buildCompactEvidenceLine(g.metricEvidence || [], 10)
            );
        }).join("\n\n");

        const executivePrompt =
            "You are a senior analyst. Write a concise Executive Summary based on the group summaries below.\n" +
            "Dataset: " + totalRows + " rows across " + totalChunks + " chunks.\n" +
            "Structure your output with these labelled sections (plain text, no JSON):\n" +
            "Overall Patterns:\n" +
            "Top KPI Signals:\n" +
            "Key Risks:\n" +
            "Key Opportunities:\n" +
            "Critical Anomalies:\n\n" +
            "Rules:\n" +
            "  - Under 500 words total.\n" +
            "  - Each section: 2-4 bullet points.\n" +
            "  - Do NOT produce final user-facing insight formatting yet.\n" +
            "  - When numeric evidence exists, include actual scores beside entities.\n" +
            "  - Report all visible Significantly Higher/Lower metrics, not just one metric.\n\n" +
            "Deterministic numeric/significance evidence extracted from all chunks:\n" +
            evidenceSummary +
            "\n\nGroup summaries:\n\n" +
            groupText +
            "\n\nExecutive Summary:";

        // FIX 1 - text_executive uses the analyst system prompt, not JSON system prompt
        // FIX 3 - TOKENS_EXECUTIVE = 1500
        const executiveSummary = await this.callAzureOpenAIWithRetry(
            executivePrompt,
            TOKENS_EXECUTIVE,
            1,
            "text_executive"
        );

        const safeExecutiveSummary = this.containsMissingDataHallucination(executiveSummary)
            ? this.buildDeterministicExecutiveSummary(groupSummaries, totalRows, totalChunks)
            : executiveSummary;

        console.log("[ExecutiveSummaryLayer] Complete.", safeExecutiveSummary);
        return safeExecutiveSummary;
    }

    // -------------------------------------------------------------------------
    // FINAL INSIGHT GENERATION LAYER (Level 4)
    // FIX 1 + FIX 3 applied - text_final call type, 4000 tokens.
    // -------------------------------------------------------------------------

    // Scans the user's prompt for an explicit "Output format:" / "Structure:" / "Sections:"
    // block and pulls out the literal heading labels the user wrote (e.g. "Summary",
    // "Key findings", or any custom names). This replaces guessing via keyword regex -
    // the actual heading text is read directly from the prompt so any naming works,
    // not just a hardcoded list.
    private extractHeadingsFromPrompt(userPrompt: string): string[] {
        const text  = String(userPrompt || "");
        const lines = text.split("\n");

        const blockMarkerPattern =/^(output format|output structure|format|structure|sections|response format|final output format|final format|return only the final output(?: in this exact format)?|return the final output(?: in this exact format)?)\s*:?\s*$/i;
        const stopMarkerPattern  = /^(style requirements?|analysis requirements?|important instructions?|notes?|instructions?)\s*:?\s*$/i;
        const bulletLine   = /^[-*-]\s*/;
        const numberedLine = /^\d+[.)]\s+/;
        const placeholderBullet = /^\[.*\]$/;

        let startIndex = -1;
        for (let i = 0; i < lines.length; i++) {
            if (blockMarkerPattern.test(lines[i].trim())) { startIndex = i + 1; break; }
        }
        if (startIndex === -1) { return []; }

        const headings: string[] = [];
        const seen: { [key: string]: boolean } = {};

        for (let i = startIndex; i < lines.length; i++) {
            const line = String(lines[i] || "").trim();
            if (!line) { continue; }
            if (stopMarkerPattern.test(line)) { break; }
            if (bulletLine.test(line) || numberedLine.test(line)) { continue; }
            if (placeholderBullet.test(line)) { continue; }
            if (line.length > 60) { continue; } // heading labels are short; long lines are body text

            const cleaned = line.replace(/:$/, "").replace(/^\*\*(.*?)\*\*$/, "$1").trim();
            if (!cleaned) { continue; }

            const key = cleaned.toLowerCase();
            if (seen[key]) { continue; }
            seen[key] = true;
            headings.push(cleaned);

            if (headings.length >= 8) { break; }
        }

        return headings;
    }

    private extractBulletLimitsFromPrompt(
        userPrompt: string,
        headings: string[]
    ): { [headingKey: string]: number } {
        const limits: { [headingKey: string]: number } = {};
        const styles: { [headingKey: string]: string } = {};
        if (!headings || headings.length === 0) {
            this.lastKnownListStylesFromPrompt = styles;
            return limits;
        }

        const lines = String(userPrompt || "").split("\n");
        const blockMarkerPattern =/^(output format|output structure|format|structure|sections|response format|final output format|final format|return only the final output(?: in this exact format)?|return the final output(?: in this exact format)?)\s*:?\s*$/i;
        const stopMarkerPattern = /^(style requirements?|analysis requirements?|important instructions?|notes?|instructions?)\s*:?\s*$/i;
        const headingLookup: { [key: string]: boolean } = {};

        for (let i = 0; i < headings.length; i++) {
            headingLookup[this.normalisePromptFieldToken(headings[i])] = true;
        }

        let insideBlock = false;
        let currentHeadingKey = "";
        for (let i = 0; i < lines.length; i++) {
            const line = String(lines[i] || "").trim();
            if (!insideBlock) {
                if (blockMarkerPattern.test(line)) { insideBlock = true; }
                continue;
            }
            if (stopMarkerPattern.test(line)) { break; }
            if (!line) { continue; }

            const headingKey = this.normalisePromptFieldToken(
                line.replace(/:$/, "").replace(/^\*\*(.*?)\*\*$/, "$1")
            );
            if (headingLookup[headingKey]) {
                currentHeadingKey = headingKey;
                if (limits[currentHeadingKey] === undefined) {
                    limits[currentHeadingKey] = 0;
                }
                continue;
            }

            if (!currentHeadingKey) { continue; }

            let detectedStyle = "";
            if (/^\d+\.\s+/.test(line)) {
                detectedStyle = "numbered-dot";
            } else if (/^\d+\)\s+/.test(line)) {
                detectedStyle = "numbered-paren";
            } else if (/^-\s+/.test(line)) {
                detectedStyle = "dash";
            } else if (/^\*\s+/.test(line)) {
                detectedStyle = "asterisk";
            } else if (/^•\s+/.test(line)) {
                detectedStyle = "bullet";
            }

            if (detectedStyle) {
                limits[currentHeadingKey] += 1;
                if (!styles[currentHeadingKey]) {
                    styles[currentHeadingKey] = detectedStyle;
                }
            }
        }

        const keys = Object.keys(limits);
        for (let i = 0; i < keys.length; i++) {
            if (limits[keys[i]] <= 0) { delete limits[keys[i]]; }
        }

        this.lastKnownListStylesFromPrompt = styles;
        return limits;
    }
    private syncPromptOutputFormatting(
        userPrompt: string
    ): void {
        const headings =
            this.extractHeadingsFromPrompt(
                userPrompt
            );
 
        this.lastKnownHeadingsFromPrompt =
            headings;
 
        this.lastKnownBulletLimitsFromPrompt =
            this.extractBulletLimitsFromPrompt(
                userPrompt,
                headings
            );
    }
 


    private getPromptListMarker(heading: string, itemIndex: number): string {
        const key = this.normalisePromptFieldToken(heading);
        const style = (this.lastKnownListStylesFromPrompt || {})[key] || "dash";

        if (style === "numbered-dot") { return String(itemIndex + 1) + "."; }
        if (style === "numbered-paren") { return String(itemIndex + 1) + ")"; }
        if (style === "asterisk") { return "*"; }
        if (style === "bullet") { return "•"; }
        return "-";
    }

    private buildFinalOutputFormatInstruction(userPrompt: string): { instruction: string; headings: string[] } {
        const headings = this.extractHeadingsFromPrompt(userPrompt);
        this.lastKnownBulletLimitsFromPrompt = this.extractBulletLimitsFromPrompt(userPrompt, headings);

        if (headings.length > 0) {
            const quotedHeadings = headings.map(function(h) { return "\"" + h + "\""; }).join(", ");
            const countRules: string[] = [];
            for (let i = 0; i < headings.length; i++) {
                const key = this.normalisePromptFieldToken(headings[i]);
                const limit = this.lastKnownBulletLimitsFromPrompt[key];
                if (limit && limit > 0) {
                    countRules.push(
                        "Section '" + headings[i] + "' must contain exactly " + limit +
                        " bullet" + (limit === 1 ? "" : "s") + "."
                    );
                }
            }

            return {
                headings: headings,
                instruction:
                    "Use exactly these section headings, in this order, and no others: " + quotedHeadings + ". " +
                    "Reproduce each heading exactly as written above, each on its own line with nothing else on that line. " +
                    "Do not rename, merge, reorder, translate, or add sections beyond these. " +
                    (countRules.length > 0 ? countRules.join(" ") + " " : "") +
                    headings.map((heading) => {
                        const key = this.normalisePromptFieldToken(heading);
                        const style = (this.lastKnownListStylesFromPrompt || {})[key] || "dash";
                        if (style === "numbered-dot") {
                            return "Section '" + heading + "' must use sequential numbering in the form 1., 2., 3., and so on.";
                        }
                        if (style === "numbered-paren") {
                            return "Section '" + heading + "' must use sequential numbering in the form 1), 2), 3), and so on.";
                        }
                        if (style === "asterisk") { return "Section '" + heading + "' must use '* ' list markers."; }
                        if (style === "bullet") { return "Section '" + heading + "' must use '• ' list markers."; }
                        return "Section '" + heading + "' must use '- ' list markers.";
                    }).join(" ") + " " +
                    "Preserve the list style requested by the user for each section. Do not convert numbered lists to bullet lists or bullet lists to numbered lists. " +
                    "Do not write body paragraphs, unlisted sentences, or tables."
            };
        }

        const explicitFormatSignal = /\b(output format|use the following format|format:|sections:|section headings|headings:|bullet points only|bullet points|no tables|numbered list|table)\b/i.test(userPrompt);

        if (explicitFormatSignal) {
            return {
                headings: [],
                instruction:
                    "Follow the user's requested structure exactly. If the user specified headings, sections, bullet counts, or bullet formatting, use those exactly. " +
                    "Do not add extra sections beyond what the user requested."
            };
        }

        return {
            headings: [],
            instruction:
                "The user did not provide an explicit output-format block. Do not impose predefined section names or a fixed section count. " +
                "Follow any formatting or wording requirements present elsewhere in the user's instruction."
        };
    }

    private buildVisibleDataContextSection(): string {
        if (!this.dataView || !this.dataView.categorical || !this.dataView.categorical.categories) {
            return "";
        }

        const descriptors = this.getAnalyticalDimensionDescriptors(
            this.dataView.categorical.categories || []
        );
        if (descriptors.length === 0) { return ""; }

        const lines = descriptors.map(function(item) {
            const role = item.isComparison ? "comparison dimension" : "context/filter dimension";
            const values = item.values.length > 0 ? item.values.join(", ") : "(blank)";
            return "- " + item.name + " (" + role + "): " + values;
        });

        return "\n\nVisible analytical dimensions from the current filtered data:\n" +
            lines.join("\n") + "\n\n";
    }

    private enforcePromptBulletLimits(content: string): string {
        const limits = this.lastKnownBulletLimitsFromPrompt || {};
        const headings = this.lastKnownHeadingsFromPrompt || [];
        if (headings.length === 0 || Object.keys(limits).length === 0) {
            return String(content || "").trim();
        }

        const parsed = this.parseInsightIntoSections(String(content || ""));
        if (parsed.length === 0) { return String(content || "").trim(); }

        const byHeading: { [key: string]: string[] } = {};
        for (let i = 0; i < parsed.length; i++) {
            const key = this.normalisePromptFieldToken(parsed[i].heading);
            if (!byHeading[key]) { byHeading[key] = []; }
            byHeading[key].push.apply(byHeading[key], parsed[i].bullets || []);
        }

        const output: string[] = [];
        for (let i = 0; i < headings.length; i++) {
            const heading = headings[i];
            const key = this.normalisePromptFieldToken(heading);
            const bullets = (byHeading[key] || []).filter(function(value) {
                return String(value || "").trim().length > 0;
            });
            const limit = limits[key];
            const selected = limit && limit > 0 ? bullets.slice(0, limit) : bullets;
            if (selected.length === 0) { continue; }

            output.push(heading);
            for (let bulletIndex = 0; bulletIndex < selected.length; bulletIndex++) {
                output.push(
                    this.getPromptListMarker(heading, bulletIndex) + " " +
                    String(selected[bulletIndex] || "")
                        .replace(/^[-*•]\s+/, "")
                        .replace(/^\d+[.)]\s+/, "")
                        .trim()
                );
            }
            output.push("");
        }

        return output.length > 0
            ? output.join("\n").trim()
            : String(content || "").trim();
    }

    private isLikelyIncompleteInsightBullet(value: string): boolean {
        const text = String(value || "").trim();
        if (!text) { return true; }

        // A clearly dangling conjunction/preposition is a strong signal that the
        // model stopped because of its output budget rather than finishing the sentence.
        if (/\b(and|or|but|with|for|to|of|in|on|at|by|also|was|were|is|are|has|have|the|a|an|that|which|while|whereas|then|where|than|because|although)\s*$/i.test(text)) {
            return true;
        }

        // Client-facing bullets are expected to be complete sentences. Accept common
        // closing punctuation after a quote/bracket as complete.
        return !/[.!?][\"')\]]?$/.test(text);
    }

    private validateFinalInsightCompleteness(content: string): { valid: boolean; reason: string } {
        const text = String(content || "").trim();
        if (!text) {
            return { valid: false, reason: "The final response is empty." };
        }

        const parsed = this.parseInsightIntoSections(text);
        const headings = this.lastKnownHeadingsFromPrompt || [];
        const limits = this.lastKnownBulletLimitsFromPrompt || {};

        // When the user supplied an Output format block, verify every requested heading
        // and every requested bullet count dynamically. Nothing here hard-codes section names.
        for (let i = 0; i < headings.length; i++) {
            const heading = headings[i];
            const headingKey = this.normalisePromptFieldToken(heading);
            const expectedCount = limits[headingKey];

            const matchingSections = parsed.filter((section) =>
                this.normalisePromptFieldToken(section.heading) === headingKey
            );

            if (matchingSections.length === 0) {
                return { valid: false, reason: "Missing requested section: " + heading };
            }

            const bullets: string[] = [];
            for (let sectionIndex = 0; sectionIndex < matchingSections.length; sectionIndex++) {
                bullets.push.apply(bullets, matchingSections[sectionIndex].bullets || []);
            }

            if (expectedCount && expectedCount > 0 && bullets.length < expectedCount) {
                return {
                    valid: false,
                    reason: "Section '" + heading + "' returned " + bullets.length +
                        " bullet(s), but the prompt requested " + expectedCount + "."
                };
            }

            const bulletsToCheck = expectedCount && expectedCount > 0
                ? bullets.slice(0, expectedCount)
                : bullets;

            for (let bulletIndex = 0; bulletIndex < bulletsToCheck.length; bulletIndex++) {
                if (this.isLikelyIncompleteInsightBullet(bulletsToCheck[bulletIndex])) {
                    return {
                        valid: false,
                        reason: "An insight bullet appears to end mid-sentence in section '" + heading + "'."
                    };
                }
            }
        }

        // Even when no explicit output block exists, prevent a visibly truncated final line.
        if (parsed.length > 0) {
            const lastSection = parsed[parsed.length - 1];
            if (lastSection.bullets && lastSection.bullets.length > 0) {
                const lastBullet = lastSection.bullets[lastSection.bullets.length - 1];
                if (this.isLikelyIncompleteInsightBullet(lastBullet)) {
                    return { valid: false, reason: "The last insight bullet appears to end mid-sentence." };
                }
            }
        }

        return { valid: true, reason: "" };
    }

    private validateDistinctMetricInsightItems(
    content: string,
    evidence: MetricEvidence[]
        ): {
            valid: boolean;
            reason: string;
        } {
            const parsed =
                this.parseInsightIntoSections(content);
        
            const knownMetrics: string[] = [];
        
            const seenMetricLabels: {
                [key: string]: boolean;
            } = {};
        
            for (let i = 0; i < evidence.length; i++) {
                const metric =
                    String(evidence[i].metric || "").trim();
            
                if (!metric) {
                    continue;
                }
            
                const key =
                    this.normaliseEvidenceKey(metric);
            
                if (!seenMetricLabels[key]) {
                    seenMetricLabels[key] = true;
                    knownMetrics.push(metric);
                }
            }
        
            for (
                let sectionIndex = 0;
                sectionIndex < parsed.length;
                sectionIndex++
            ) {
                const bullets =
                    parsed[sectionIndex].bullets || [];
            
                const usedMetricKeys: {
                    [key: string]: boolean;
                } = {};
            
                for (
                    let bulletIndex = 0;
                    bulletIndex < bullets.length;
                    bulletIndex++
                ) {
                    const bullet =
                        String(bullets[bulletIndex] || "");
                
                    let foundMetric = "";
                
                    for (
                        let metricIndex = 0;
                        metricIndex < knownMetrics.length;
                        metricIndex++
                    ) {
                        const metric =
                            knownMetrics[metricIndex];
                    
                        if (
                            bullet.toLowerCase().indexOf(
                                metric.toLowerCase()
                            ) !== -1
                        ) {
                            foundMetric = metric;
                            break;
                        }
                    }
                
                    if (!foundMetric) {
                        continue;
                    }
                
                    const metricKey =
                        this.normaliseEvidenceKey(
                            foundMetric
                        );
                    
                    if (usedMetricKeys[metricKey]) {
                        return {
                            valid: false,
                            reason:
                                "The same Metric was split across multiple list items: " +
                                foundMetric +
                                "."
                        };
                    }
                
                    usedMetricKeys[metricKey] = true;
                }
            }
        
            return {
                valid: true,
                reason: ""
            };
        }


    private repairUnavailableDimensionPlaceholders(content: string): string {
        let repaired = String(content || "");
        const categories = this.dataView && this.dataView.categorical
            ? (this.dataView.categorical.categories || [])
            : [];
        const descriptors = this.getAnalyticalDimensionDescriptors(categories);

        for (let i = 0; i < descriptors.length; i++) {
            const descriptor = descriptors[i];
            if (descriptor.values.length !== 1) { continue; }
            const actualValue = descriptor.values[0];
            if (!actualValue) { continue; }

            const unavailablePattern = new RegExp(
                "\\b" + this.escapeRegExp(descriptor.name) +
                "\\s*(?:[:=]\\s*)?(?:unavailable|not available|not provided|missing)\\b",
                "gi"
            );
            repaired = repaired.replace(unavailablePattern, actualValue);
        }

        return repaired;
    }

    private async generateFinalInsight(
        userPrompt: string,
        executiveSummary: string,
        evidenceSummary: string,
        deterministicEvidence: MetricEvidence[],
        basePromptIsDefault: boolean
    ): Promise<string> {
        const deterministicClientInsight = this.buildClientReadyEvidenceInsight(deterministicEvidence);
        const roleAwareComparisonEvidence = this.buildRoleAwareComparisonEvidence(
            deterministicEvidence,
            userPrompt
        );
        const visibleDataContextSection = this.buildVisibleDataContextSection();
        const categories = this.dataView && this.dataView.categorical
            ? (this.dataView.categorical.categories || [])
            : [];
        const descriptors = this.getAnalyticalDimensionDescriptors(categories);
        const grainInstruction = this.buildAnalyticalGrainInstruction(descriptors, userPrompt);

        const requestedDimensions = this.getPromptRequestedDimensionNames(userPrompt, descriptors);
        const formatResult = this.buildFinalOutputFormatInstruction(userPrompt);
        const formatInstruction = formatResult.instruction;

        this.lastKnownHeadingsFromPrompt = formatResult.headings;

        // const dimensionRequirement = requestedDimensions.length > 0
        //     ? "The user explicitly requested these dimensions: " + requestedDimensions.join(", ") +
        //       ". Every applicable high, low, and comparison bullet must state their exact values.\n"
        //     : "Use the exact values of all relevant visible analytical dimensions in each observation.\n";

        // Dynamically collect every analytical Category Data field
        // currently available to the visual.
        // Nothing such as Market, Brand, Age or Time Period is hard-coded.
        
        const visibleDimensionNames: string[] = [];
        const seenVisibleDimensions: {
            [key: string]: boolean;
        } = {};
 
        for (
            let dimensionIndex = 0;
            dimensionIndex < descriptors.length;
            dimensionIndex++
        ) {
            const dimensionName =
                String(
                 descriptors[dimensionIndex].name || ""
                ).trim();
 
            if (!dimensionName) {
                continue;
            }
 
         const dimensionKey =
                this.normalisePromptFieldToken(
                 dimensionName
                );
 
            if (
                dimensionKey &&
                !seenVisibleDimensions[dimensionKey]
            ) {
             seenVisibleDimensions[dimensionKey] = true;
                visibleDimensionNames.push(
                 dimensionName
             );
            }
        }
 
        const dimensionRequirement =
            visibleDimensionNames.length > 0
            ?
            "For every highest, lowest, positive increase, decline, " +
            "subgroup comparison, and other row-specific observation, " +
            "include the exact nonblank value of EVERY visible analytical " +
            "Category Data dimension belonging to that SAME evidence row. " +
            "The visible dimensions in the current selection are: " +
            visibleDimensionNames.join(", ") +
            ". Do not omit a visible dimension merely because the user did " +
            "not explicitly name it in the prompt. " +
            "Do not borrow a dimension value from another row. " +
            "Do not invent a value when the evidence row is blank for that dimension.\n"
            :
            "Use only dimensions actually present in the evidence row.\n";
        

        const hasConversionEvidance =
            deterministicEvidence.some(function(iteam){
                return(
                    iteam.Conversion !== null &&
                    iteam.Conversion !== undefined &&
                    isFinite(iteam.Conversion)
                );
            });
        const conversionInstruction = 
            hasConversionEvidance ? 
            "-Conversion comes only from the dedicated Conversion role. " +
            "When reported, use the Conversion value from the same evidence row as the Metric, Score and dimensions. " +
            "Do not invent Conversion.\n"
            : "";

        const systemBlock =
            "You are a concise market-research analyst for Power BI. Use the user instruction as the primary instruction.\n" +
            "The user's requested output headings, section count, bullet count, wording, and structure are presentation requirements. Do not invent or hard-code additional headings when the user supplied a format.\n" +
            "If the user did not request a specific output format, do not force predefined section names.\n" +
            formatInstruction + "\n" +
            "AUTHORITATIVE DATA SEMANTICS:\n" +
            "- The role-aware evidence block is authoritative for Metric, Score/Value, Change, Significance, and row dimensions. If any later narrative summary conflicts with it, ignore the conflicting narrative.\n" +
            "- Metric labels must be copied exactly from the Metric/KPI evidence. Never create a new Metric name and never append a measure field, Change field, Significance field, or another Metric to the Metric label.\n" +
            "- Score comes only from the score/value Measure Data evidence.\n" +
            "- Change comes only from the dedicated Change role and must never be reused as Score.\n" +
            conversionInstruction +
            "- Significance is text from the dedicated Significance role. Preserve text such as High Sig, Low Sig, No Significance, or any other provided label exactly; do not infer its wording.\n" +
            "- The direction of an increase or decline comes from the numeric Change sign/value, not from the Significance text label.\n" +
            "- When a Change is reported, its Score, Significance, Metric, and dimensions must come from the same evidence row.\n" +
            "- Do not call a Score or Change highest/lowest unless the authoritative comparison block confirms the comparison across all relevant visible rows.\n" +
            "- If a requested subgroup value is present in the prompt, use the subgroup comparison facts for that exact visible dimension value. Never substitute another entity.\n" +
            "- If no numeric Change exists for a Metric, do not invent Change language. Use Score/Value evidence only unless the user explicitly asks for something else.\n" +
            "- If the user instruction says to use Change and/or Significance whenever they are present, then any nonblank dedicated-role Change/Significance evidence means that branch applies. Do not fall back to a score-only sentence when those role values exist for the compared rows.\n" +
            "- Never invent a Metric, Score, Change, Significance label, dimension, dimension value, market, category, demographic, brand, or time period.\n" +
            "- Do not give suggestions, recommendations, actions, next steps, opportunities, or strategic advice unless the user explicitly asks for them.\n" +
            "Highlight Entity is presentation-only. It controls red/bold emphasis and must never alter analytical scope.\n" +
            "A field assigned to both Category Data and Highlight Entity remains a complete analytical dimension.\n" +
            "Treat every distinct Metric label as a separate metric. Never merge differently named metrics.\n" +
            "INSIGHT UNIT RULE:\n" +
            "- Each requested list item must represent ONE complete Metric insight, not one isolated fact.\n" +
            "- Do not split the highest positive Change, highest decline, requested subgroup comparison, or related score details for the SAME Metric across separate list items.\n" +
            "- When Change exists for a Metric, combine that Metric's applicable overall highest positive Change, overall highest decline, and any user-requested subgroup high/decline comparison into the SAME list item.\n" +
            "- Move to the next list item only when moving to a different Metric.\n" +
            "- The requested list-item count refers to the number of complete Metric insights, not the number of individual comparison facts.\n" +
            "- Never use multiple list items merely to break apart one Metric's required template.\n" +
            grainInstruction +
            dimensionRequirement +
            "Dimension completeness is mandatory: if an evidence row contains a nonblank " +
            "dimension value, preserve it in the final observation. Never shorten an " +
            "observation by dropping one of its row dimensions.\n" +
 
            "When comparing a demographic value such as Female versus Male, hold Metric and every other analytical dimension constant when that is the comparison requested by the user.\n" +
            "Never replace separate demographic values with their mean unless the user explicitly asks for a combined average or an explicit aggregate row exists.\n" +
            "Do not output an unavailable/missing placeholder when the current data contains an explicit value.\n" +
            "Use only values present in the authoritative evidence.\n\n";

        const userBlock =
            "User task:\n" +
            userPrompt +
            visibleDataContextSection +
            "\nAUTHORITATIVE role-aware complete-dataset comparison facts:\n" +
            roleAwareComparisonEvidence +
            "\n\nAUTHORITATIVE row-level numeric / Change / Significance evidence:\n" +
            evidenceSummary +
            "\n\nSecondary client-ready narrative evidence (use only when consistent with the authoritative blocks above):\n" +
            deterministicClientInsight +
            "\n\nSecondary executive evidence (use only when consistent with the authoritative blocks above and never copy suggestions the user did not request):\n" +
            executiveSummary +
            "\n\n";

        const finalPrompt = systemBlock + userBlock + "Write the insight now.";
        let finalInsight = await this.callAzureOpenAIWithRetry(
            finalPrompt,
            TOKENS_FINAL_INSIGHT,
            1,
            "text_final",
            !basePromptIsDefault
        );

        // The model can occasionally stop at its output limit while still returning non-empty
        // text. Non-empty text used to be accepted, displayed, and cached even when the last
        // sentence was cut off or the requested bullet count was not reached. Validate the
        // user-requested structure and completeness, then regenerate a complete replacement.
        let completionCheck = this.validateFinalInsightCompleteness(finalInsight);
        // Make sure one Metric has not been split
        // across multiple numbered items.
        if (completionCheck.valid) {
            const metricStructureCheck =
                this.validateDistinctMetricInsightItems(
                    finalInsight,
                    deterministicEvidence
                );
            
            if (!metricStructureCheck.valid) {
                completionCheck =
                    metricStructureCheck;
            }
        }
 
        if (!completionCheck.valid) {
            console.warn("[FinalInsightLayer] Incomplete response detected: " + completionCheck.reason);

            const completionRepairPrompt =
                finalPrompt +
                "\n\nCRITICAL COMPLETION RETRY:\n" +
                "The previous answer was incomplete or did not satisfy the user's requested bullet count. " +
                "Return a COMPLETE REPLACEMENT answer from the first requested heading through the last requested bullet. " +
                "Follow the user's headings and bullet counts exactly. " +
                "Each requested list item must remain one complete Metric insight. " +
                "Do not satisfy the required item count by splitting one Metric's increase, decline, subgroup comparison, or score facts into separate items. " +
                "Use a different Metric for each list item whenever enough distinct comparable Metrics are available. " +
                "Do not continue the previous answer. Do not leave any sentence unfinished. " +
                "Keep every bullet concise enough to finish within the available output budget while preserving all required factual fields. " +
                "Do not omit requested bullets merely to stay concise.\n" +
                "Validation failure: " + completionCheck.reason + "\n";

            finalInsight = await this.callAzureOpenAIWithRetry(
                completionRepairPrompt,
                TOKENS_FINAL_INSIGHT,
                1,
                "text_final",
                !basePromptIsDefault
            );
            completionCheck = this.validateFinalInsightCompleteness(finalInsight);

            if (!completionCheck.valid) {
                console.warn("[FinalInsightLayer] Completion retry still failed validation: " + completionCheck.reason);
            }
        }

        const safeFinalInsight = this.containsMissingDataHallucination(finalInsight)
            ? this.buildDeterministicFinalInsight(executiveSummary)
            : finalInsight;
        const finalText = String(safeFinalInsight || "").trim();
        const fallbackText = finalText || this.buildDeterministicFinalInsight(executiveSummary);
        const repairedText = this.repairUnavailableDimensionPlaceholders(fallbackText);
        const constrainedText = this.enforcePromptBulletLimits(repairedText);

        console.log("[FinalInsightLayer] Complete.");
        return constrainedText;
    }


    // -------------------------------------------------------------------------
    // Output safety helpers - prevent "pending dataset" hallucinations after
    // chunking has already completed.
    // -------------------------------------------------------------------------

    private containsMissingDataHallucination(text: string): boolean {
        const lower = String(text || "").toLowerCase();
        return (
            lower.indexOf("no data provided") !== -1 ||
            lower.indexOf("pending dataset") !== -1 ||
            lower.indexOf("please upload") !== -1 ||
            lower.indexOf("upload or paste") !== -1 ||
            lower.indexOf("provide the dataset") !== -1 ||
            lower.indexOf("data required") !== -1 ||
            lower.indexOf("once data is available") !== -1
        );
    }

    private takeUnique(items: string[], limit: number): string[] {
        const seen: { [key: string]: boolean } = {};
        const output: string[] = [];

        for (let i = 0; i < items.length; i++) {
            const item = String(items[i] || "").trim();
            if (!item) { continue; }

            const key = item.toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();
            if (seen[key]) { continue; }

            seen[key] = true;
            output.push(item);
            if (output.length >= limit) { break; }
        }

        return output;
    }


    private parseDelimitedLine(line: string): string[] {
        const values: string[] = [];
        let current = "";
        let inQuotes = false

        for (let i = 0; i < line.length; i++) {
            const ch = line.charAt(i);
            const next = i + 1 < line.length ? line.charAt(i + 1) : "";

            if (ch === '"') {
                if (inQuotes && next === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch === "," && !inQuotes) {
                values.push(current.trim());
                current = "";
            } else {
                current += ch;
            }
        }

        values.push(current.trim());
        return values;
    }

    private normaliseEvidenceKey(value: string): string {
        return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, "").trim();
    }

    private parseNumericValue(value: string): number | null {
        const raw = String(value == null ? "" : value).trim();
        if (!raw || raw.toLowerCase() === "n/a" || raw.toLowerCase() === "null") { return null; }

        // IMPORTANT: this must be strict.
        // The older implementation extracted the first number from any text, so
        // labels such as "Brand 7", "Country 1", and "Jun 2022" were treated as
        // numeric scores. That caused the evidence layer to choose Country as the
        // entity and to show impossible values such as 202200.0%.
        // Accept only cells that are numeric values themselves, with optional
        // formatting such as %, commas, currency symbols, parentheses, and K/M/B.
        let cleaned = raw.replace(/,/g, "").trim();
        cleaned = cleaned.replace(/^[$?£?¥]\s*/, "").replace(/\s*[$?£?¥]$/, "");

        let negative = false;
        const paren = cleaned.match(/^\((.*)\)$/);
        if (paren && paren[1]) {
            negative = true;
            cleaned = paren[1].trim();
        }

        let multiplier = 1;
        const suffix = cleaned.match(/([kmb])\s*%?$/i);
        if (suffix && suffix[1]) {
            const s = suffix[1].toLowerCase();
            if (s === "k") { multiplier = 1000; }
            if (s === "m") { multiplier = 1000000; }
            if (s === "b") { multiplier = 1000000000; }
            cleaned = cleaned.replace(/([kmb])\s*(%?)$/i, "$2").trim();
        }

        cleaned = cleaned.replace(/%$/, "").trim();
        if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(cleaned)) { return null; }

        let parsed = Number(cleaned) * multiplier;
        if (negative) { parsed = -parsed; }
        return isFinite(parsed) ? parsed : null;
    }

    private looksLikeSignificanceHeader(header: string): boolean {
        const text = String(header || "").toLowerCase();
        return text.indexOf("significance") !== -1 ||
            /(^|[^a-z])sig([^a-z]|$)/.test(text) ||
            text.indexOf("sig_") !== -1 ||
            text.indexOf("_sig") !== -1;
    }

    private looksLikeSignificanceText(value: string): boolean {
        const text = String(value || "").toLowerCase();
        return text.indexOf("significant") !== -1 ||
            text.indexOf("higher") !== -1 ||
            text.indexOf("lower") !== -1 ||
            text.indexOf("increase") !== -1 ||
            text.indexOf("decrease") !== -1 ||
            text.indexOf("above") !== -1 ||
            text.indexOf("below") !== -1;
    }

    private extractMetricNameFromSignificanceHeader(header: string): string {
        const raw = String(header || "").trim();

        // Dynamic Power BI pivot headers are often built as:
        //   <actual significance field display name> (<metric value>)
        // Example: "Significance_Text (Awareness)".
        // The old logic returned "Text (Awareness)", which did not match the
        // numeric metric header "Awareness" and therefore the significance value
        // was lost.  This is not a hardcoded column-name rule; it simply extracts
        // the metric label from the generated parenthesised header when present.
        const parenthesisedMetric = raw.match(/\(([^()]+)\)\s*$/);
        if (parenthesisedMetric && parenthesisedMetric[1]) {
            return parenthesisedMetric[1].trim();
        }

        let metric = raw;
        metric = metric.replace(/^significance[\s_:\-]*/i, "");
        metric = metric.replace(/[\s_:\-]*significance$/i, "");
        metric = metric.replace(/^sig[\s_:\-]*/i, "");
        metric = metric.replace(/[\s_:\-]*sig$/i, "");
        metric = metric.replace(/^text[\s_:\-]*/i, "");
        metric = metric.replace(/[\s_:\-]*text$/i, "");
        return metric.trim() || raw;
    }

    private splitMeasureMetricHeader(header: string, knownMetrics: string[]): { measure: string; metric: string } {
        const raw = String(header || "").trim();
        const normalisedHeader = this.normaliseEvidenceKey(raw);

        let bestMetric = "";
        for (let i = 0; i < knownMetrics.length; i++) {
            const metric = knownMetrics[i];
            const metricKey = this.normaliseEvidenceKey(metric);
            if (!metricKey) { continue; }
            if (normalisedHeader === metricKey || normalisedHeader.lastIndexOf(metricKey) === normalisedHeader.length - metricKey.length) {
                if (metric.length > bestMetric.length) { bestMetric = metric; }
            }
        }

        if (bestMetric) {
            const escapedMetric = bestMetric.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const measure = raw.replace(new RegExp("[\\s_:\\-]*" + escapedMetric + "$", "i"), "").trim();
            return { measure: measure || "Score", metric: bestMetric };
        }

        const separators = ["_", " - ", " | ", ":"];
        for (let i = 0; i < separators.length; i++) {
            const sep = separators[i];
            const idx = raw.lastIndexOf(sep);
            if (idx > 0 && idx < raw.length - sep.length) {
                return {
                    measure: raw.substring(0, idx).trim() || "Score",
                    metric: raw.substring(idx + sep.length).trim() || raw
                };
            }
        }

        return { measure: raw || "Score", metric: raw || "Metric" };
    }


    private looksLikeDateOrPeriodValue(value: string): boolean {
        const text = String(value || "").trim().toLowerCase();
        if (!text) { return false; }
        if (/^\d{4}$/.test(text)) { return true; }
        if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(text)) { return true; }
        if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{2,4}$/.test(text)) { return true; }
        if (/^\d{4}[\/\-]\d{1,2}([\/\-]\d{1,2})?$/.test(text)) { return true; }
        return false;
    }

    private isMostlySameValue(values: string[]): boolean {
        const nonBlank = values.filter(function(v) { return String(v || "").trim().length > 0; });
        if (nonBlank.length === 0) { return true; }
        const unique: { [key: string]: boolean } = {};
        for (let i = 0; i < nonBlank.length; i++) {
            unique[String(nonBlank[i]).trim().toLowerCase()] = true;
        }
        return Object.keys(unique).length <= 1;
    }

    private looksLikeEntityOrDimensionHeader(header: string): boolean {
        const h = String(header || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
        if (!h) { return false; }

        // Generic semantic hints only. These are not project-specific field names;
        // they protect categorical IDs such as brand/product/customer codes from
        // being misclassified as numeric measures.
        return /\b(brand|entity|customer|client|account|product|sku|item|segment|category|group|market|country|region|location|geo|channel|store|retailer|supplier|vendor|campaign|page|period|date|time|year|month|week|quarter|name|label|code|id)\b/.test(h);
    }

    private looksLikeMeasureOrScoreHeader(header: string): boolean {
        const h = String(header || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
        if (!h) { return false; }

        return /\b(score|value|actual|target|measure|metric value|percentage|percent|pct|rate|ratio|index|share|sales|revenue|volume|units|count|amount|total|avg|average|sum|mean|median|kpi|rolling|baseline|base|latest|current)\b/.test(h);
    }

    private isWeakEntityLabel(value: string): boolean {
        const text = String(value || "").trim();
        if (!text) { return true; }
        if (/^row\s*\d+$/i.test(text)) { return true; }
        if (/^\d+(?:\.\d+)?$/.test(text)) { return true; }
        if (text.toLowerCase() === "n/a") { return true; }
        return false;
    }

    private selectPrimaryEntityDimensionIndex(
        headers: string[],
        rows: string[][],
        dimensionIndexes: number[]
    ): number {
        if (!dimensionIndexes || dimensionIndexes.length === 0) { return -1; }

        if (this.preferredEntityFieldFromPrompt) {
            const preferredKey = this.normaliseEvidenceKey(this.preferredEntityFieldFromPrompt);
            for (let p = 0; p < dimensionIndexes.length; p++) {
                const headerKey = this.normaliseEvidenceKey(headers[dimensionIndexes[p]] || "");
                if (headerKey && headerKey === preferredKey) {
                    return dimensionIndexes[p];
                }
            }
        }

        let bestIndex = dimensionIndexes[0];
        let bestScore = -Infinity;
        const sampleSize = Math.min(rows.length, 200);

        for (let d = 0; d < dimensionIndexes.length; d++) {
            const index = dimensionIndexes[d];
            const header = String(headers[index] || "");
            const headerLower = header.toLowerCase();
            const values: string[] = [];

            for (let r = 0; r < sampleSize; r++) {
                values.push(String((rows[r] && rows[r][index] != null) ? rows[r][index] : "").trim());
            }

            const nonBlank = values.filter(function(v) { return v.length > 0 && v.toLowerCase() !== "n/a"; });
            if (nonBlank.length === 0) { continue; }

            const unique: { [key: string]: boolean } = {};
            let totalLength = 0;
            let numericCount = 0;
            let dateCount = 0;
            let longTextCount = 0;

            for (let i = 0; i < nonBlank.length; i++) {
                const value = nonBlank[i];
                unique[value.toLowerCase()] = true;
                totalLength += value.length;
                if (this.parseNumericValue(value) !== null) { numericCount++; }
                if (this.looksLikeDateOrPeriodValue(value)) { dateCount++; }
                if (value.length > 60) { longTextCount++; }
            }

            const uniqueCount = Object.keys(unique).length;
            const avgLength = totalLength / Math.max(1, nonBlank.length);
            const uniqueRatio = uniqueCount / Math.max(1, nonBlank.length);

            let score = 0;

            // Entity dimensions usually contain repeated-but-varied labels after Power BI filters.
            if (uniqueCount > 1) { score += 35; }
            score += Math.min(uniqueCount, 40);
            if (uniqueRatio > 0.02 && uniqueRatio < 0.95) { score += 12; }
            if (avgLength >= 2 && avgLength <= 45) { score += 10; }

            // Generic role hints. These are not project-specific column requirements;
            // they only help choose the best display label when multiple dimensions exist.
            if (/brand|entity|customer|account|product|sku|item|segment|category|name|label|group/.test(headerLower)) { score += 60; }
            if (/prompt|question|instruction|description|comment|note|text/.test(headerLower)) { score -= 45; }
            if (/period|date|time|year|month|week|quarter/.test(headerLower)) { score -= 45; }
            if (/page|kpi|metric|measure|significance/.test(headerLower)) { score -= 35; }
            if (/country|region|market|geo|location/.test(headerLower)) { score -= uniqueCount <= 3 ? 45 : 15; }

            if (numericCount / Math.max(1, nonBlank.length) > 0.5) { score -= 35; }
            if (dateCount / Math.max(1, nonBlank.length) > 0.5) { score -= 35; }
            if (longTextCount / Math.max(1, nonBlank.length) > 0.3) { score -= 35; }
            if (uniqueCount <= 1) { score -= 20; }

            if (score > bestScore) {
                bestScore = score;
                bestIndex = index;
            }
        }

        return bestIndex;
    }

    private buildEntityLabelFromDimensions(
        dimensions: { [key: string]: string },
        primaryHeader: string
    ): string {
        const cleanValue = (value: string): string => String(value || "").trim();
        const formatValue = (header: string, value: string): string => {
            const clean = cleanValue(value);
            if (!clean) { return ""; }

            // If the selected entity value is a numeric code, keep it understandable
            // by pairing it with the dynamic field display name instead of returning
            // a vague row number.
            if (/^\d+(?:\.\d+)?$/.test(clean)) {
                return String(header || "Entity").trim() + " " + clean;
            }
            return clean;
        };

        const keys = Object.keys(dimensions);

        // Highest priority: the business instruction from Prompt as text can choose
        // the analysis grain by mentioning an available Category Data field.
        if (this.preferredEntityFieldFromPrompt) {
            const preferredKey = this.normaliseEvidenceKey(this.preferredEntityFieldFromPrompt);
            for (let i = 0; i < keys.length; i++) {
                const header = keys[i];
                if (this.normaliseEvidenceKey(header) !== preferredKey) { continue; }
                const value = cleanValue(dimensions[header] || "");
                if (value && value.toLowerCase() !== "n/a" && value.length <= 80) {
                    return formatValue(header, value);
                }
            }
        }

        if (primaryHeader) {
            const primaryValue = cleanValue(dimensions[primaryHeader] || "");
            if (primaryValue && primaryValue.toLowerCase() !== "n/a" && primaryValue.length <= 80) {
                return formatValue(primaryHeader, primaryValue);
            }
        }

        const candidates: { header: string; value: string; score: number }[] = [];

        for (let i = 0; i < keys.length; i++) {
            const header = keys[i];
            const value = cleanValue(dimensions[header] || "");
            if (!value || value.toLowerCase() === "n/a") { continue; }
            if (value.length > 60) { continue; }
            if (this.looksLikeDateOrPeriodValue(value)) { continue; }

            let score = 0;
            const headerLower = header.toLowerCase();
            if (this.looksLikeEntityOrDimensionHeader(header)) { score += 20; }
            if (/brand|entity|customer|client|account|product|sku|item|segment|category|name|label|code|id/.test(headerLower)) { score += 20; }
            if (/country|region|market|geo|location|period|date|time|year|month|week|quarter|page/.test(headerLower)) { score -= 10; }
            if (!/^\d+(?:\.\d+)?$/.test(value)) { score += 10; }
            if (value.length >= 2 && value.length <= 40) { score += 5; }

            candidates.push({ header: header, value: value, score: score });
        }

        candidates.sort(function(a, b) { return b.score - a.score; });
        if (candidates.length > 0) {
            return formatValue(candidates[0].header, candidates[0].value);
        }

        return "";
    }

    private extractMetricEvidenceFromChunk(
        chunkText: string,
        chunkIndex: number,
        totalChunks: number,
        rowRange: string
    ): MetricEvidence[] {
        const rawLines = String(chunkText || "")
            .split("\n")
            .filter(function(line) {
                const trimmed = line.trim();
                return trimmed.length > 0 && trimmed.indexOf("NOTE:") !== 0 && trimmed.indexOf("...") !== 0;
            });

        if (rawLines.length < 2) { return []; }

        const headers = this.parseDelimitedLine(rawLines[0]);
        if (headers.length === 0) { return []; }

        const rows = rawLines.slice(1).map((line) => this.parseDelimitedLine(line));
        const significanceIndexes: number[] = [];
        const knownMetrics: string[] = [];

        for (let h = 0; h < headers.length; h++) {
            if (this.looksLikeSignificanceHeader(headers[h])) {
                significanceIndexes.push(h);
                const metricName = this.extractMetricNameFromSignificanceHeader(headers[h]);
                if (metricName && knownMetrics.indexOf(metricName) === -1) { knownMetrics.push(metricName); }
            }
        }

        // Some exports may not name the significance column clearly. Detect it from values.
        for (let h = 0; h < headers.length; h++) {
            if (significanceIndexes.indexOf(h) !== -1) { continue; }
            let sigValueCount = 0;
            const sample = Math.min(rows.length, 60);
            for (let r = 0; r < sample; r++) {
                if (this.looksLikeSignificanceText(rows[r][h] || "")) { sigValueCount++; }
            }
            if (sample > 0 && sigValueCount >= Math.max(2, Math.ceil(sample * 0.25))) {
                significanceIndexes.push(h);
                const metricName = this.extractMetricNameFromSignificanceHeader(headers[h]);
                if (metricName && knownMetrics.indexOf(metricName) === -1) { knownMetrics.push(metricName); }
            }
        }

        const numericIndexes: number[] = [];
        for (let h = 0; h < headers.length; h++) {
            if (significanceIndexes.indexOf(h) !== -1) { continue; }

            const header = headers[h] || "";
            const allowedScoreHeaderCount = Object.keys(this.allowedScoreHeaderKeys || {}).length;
            if (allowedScoreHeaderCount > 0 && !this.allowedScoreHeaderKeys[this.normaliseEvidenceKey(header)]) {
                continue;
            }
            const headerLooksLikeDimension = this.looksLikeEntityOrDimensionHeader(header);
            const headerLooksLikeMeasure   = this.looksLikeMeasureOrScoreHeader(header);
            const normalisedHeader = this.normaliseEvidenceKey(header);
            let headerHasKnownMetricSuffix = false;
            for (let km = 0; km < knownMetrics.length; km++) {
                const metricKey = this.normaliseEvidenceKey(knownMetrics[km]);
                if (metricKey && normalisedHeader !== metricKey && normalisedHeader.lastIndexOf(metricKey) === normalisedHeader.length - metricKey.length) {
                    headerHasKnownMetricSuffix = true;
                    break;
                }
            }

            // Inspect sample rows to see if the column is predominantly numeric
            let nonBlankCount = 0;
            let numericCount = 0;
            const sample = Math.min(rows.length, 60);
            for (let r = 0; r < sample; r++) {
                const cell = rows[r][h];
                if (String(cell).trim().length > 0 && String(cell).toLowerCase() !== "n/a") { nonBlankCount++; }
                if (this.parseNumericValue(cell) !== null) { numericCount++; }
            }
            if (nonBlankCount > 0 && numericCount >= Math.max(1, Math.ceil(nonBlankCount * 0.6))) {
                numericIndexes.push(h);
            }
        }

        const dimensionIndexes: number[] = [];
        for (let h = 0; h < headers.length; h++) {
            if (numericIndexes.indexOf(h) === -1 && significanceIndexes.indexOf(h) === -1) {
                dimensionIndexes.push(h);
            }
        }

        const primaryEntityDimensionIndex = this.selectPrimaryEntityDimensionIndex(headers, rows, dimensionIndexes);
        const primaryEntityHeader = primaryEntityDimensionIndex >= 0 ? (headers[primaryEntityDimensionIndex] || "") : "";

        const significanceByMetricKey: { [metricKey: string]: number } = {};
        for (let i = 0; i < significanceIndexes.length; i++) {
            const index = significanceIndexes[i];
            const metricName = this.extractMetricNameFromSignificanceHeader(headers[index]);
            significanceByMetricKey[this.normaliseEvidenceKey(metricName)] = index;
        }

        const evidence: MetricEvidence[] = [];

        for (let r = 0; r < rows.length; r++) {
            const row = rows[r];
            const dimensions: { [key: string]: string } = {};

            for (let d = 0; d < dimensionIndexes.length; d++) {
                const index = dimensionIndexes[d];
                const header = headers[index] || ("Dimension_" + (d + 1));
                const value = row[index] == null ? "" : String(row[index]).trim();
                dimensions[header] = value;
            }

            const entity = this.buildEntityLabelFromDimensions(dimensions, primaryEntityHeader) || ("Row " + (r + 1));

            for (let n = 0; n < numericIndexes.length; n++) {
                const index = numericIndexes[n];
                const score = this.parseNumericValue(row[index] || "");
                if (score === null) { continue; }

                const headerInfo = this.splitMeasureMetricHeader(headers[index], knownMetrics);
                const metricKey = this.normaliseEvidenceKey(headerInfo.metric);
                let significance = "";

                if (significanceByMetricKey[metricKey] !== undefined) {
                    significance = row[significanceByMetricKey[metricKey]] || "";
                } else if (significanceIndexes.length === 1) {
                    significance = row[significanceIndexes[0]] || "";
                } else {
                    // Last-resort fuzzy match between score header and significance header.
                    for (let s = 0; s < significanceIndexes.length; s++) {
                        const sigMetric = this.extractMetricNameFromSignificanceHeader(headers[significanceIndexes[s]]);
                        const sigKey = this.normaliseEvidenceKey(sigMetric);
                        if (sigKey && (metricKey.indexOf(sigKey) !== -1 || sigKey.indexOf(metricKey) !== -1)) {
                            significance = row[significanceIndexes[s]] || "";
                            break;
                        }
                    }
                }

                evidence.push({
                    chunkIndex: chunkIndex,
                    rowRange: rowRange,
                    entity: entity,
                    metric: headerInfo.metric || headers[index] || "Metric",
                    measure: headerInfo.measure || "Score",
                    score: score,
                    scoreText: row[index] || String(score),
                    change: null,
                    changeText: "",
                    Conversion: null,
                    ConversionText:"",
                    significance: String(significance || "").trim(),
                    dimensions: dimensions
                });
            }
        }

        console.log(
            "[EvidenceLayer] Extracted " + evidence.length +
            " numeric evidence rows from chunk " + (chunkIndex + 1) + "/" + totalChunks + "."
        );

        return evidence;
    }

    private collectMetricEvidence(groupSummaries: GroupSummary[]): MetricEvidence[] {
        const evidence: MetricEvidence[] = [];
        for (let i = 0; i < groupSummaries.length; i++) {
            evidence.push.apply(evidence, groupSummaries[i].metricEvidence || []);
        }
        return evidence;
    }

    private getValueColumnsByRole(
        values: powerbi.DataViewValueColumns | undefined,
        roleName: string
    ): powerbi.DataViewValueColumn[] {
        if (!values) { return []; }
        return (Array.from(values) as powerbi.DataViewValueColumn[]).filter((valueColumn) => {
            return !!valueColumn && this.sourceHasRole(valueColumn.source, roleName);
        });
    }

    private getCategoryColumnsByRole(
        categories: powerbi.DataViewCategoryColumn[] | undefined,
        roleName: string
    ): powerbi.DataViewCategoryColumn[] {
        if (!categories) { return []; }
        return categories.filter((category) => {
            return !!category && this.sourceHasRole(category.source, roleName);
        });
    }

    private getFirstNonBlankCategoryRoleTextAtRow(
        columns: powerbi.DataViewCategoryColumn[],
        rowIndex: number
    ): string {
        for (let i = 0; i < columns.length; i++) {
            const values = columns[i].values || [];
            if (rowIndex >= values.length) { continue; }
            const text = String(values[rowIndex] == null ? "" : values[rowIndex]).trim();
            if (text && text.toLowerCase() !== "n/a") { return text; }
        }
        return "";
    }

    private getFirstNonBlankRoleTextAtRow(
        columns: powerbi.DataViewValueColumn[],
        rowIndex: number
    ): string {
        for (let i = 0; i < columns.length; i++) {
            const values = columns[i].values || [];
            if (rowIndex >= values.length) { continue; }
            const text = String(values[rowIndex] == null ? "" : values[rowIndex]).trim();
            if (text && text.toLowerCase() !== "n/a") { return text; }
        }
        return "";
    }

    private getFirstNumericRoleValueAtRow(
        columns: powerbi.DataViewValueColumn[],
        rowIndex: number
    ): { value: number | null; text: string } {
        for (let i = 0; i < columns.length; i++) {
            const values = columns[i].values || [];
            if (rowIndex >= values.length) { continue; }
            const text = String(values[rowIndex] == null ? "" : values[rowIndex]).trim();
            if (!text || text.toLowerCase() === "n/a") { continue; }
            const parsed = this.parseNumericValue(text);
            if (parsed !== null) { return { value: parsed, text: text }; }
        }
        return { value: null, text: "" };
    }

    private extractRoleAwareMetricEvidenceFromDataView(): MetricEvidence[] {
        if (!this.dataView || !this.dataView.categorical) { return []; }

        const categorical = this.dataView.categorical;
        const categories = categorical.categories || [];
        const values = categorical.values;
        if (!values || values.length === 0) { return []; }

        const split = this.splitCategoriesByRoles(categories);
        const metricCategory = split.metricCategories.length > 0
            ? split.metricCategories[0]
            : null;
        const dimensionCategories = this.getAnalyticalCategoryColumns(categories);

        // These are schema-role identifiers from capabilities.json, not business
        // field display names or data values. The actual field names and values stay dynamic.
        const changeColumns = this.getValueColumnsByRole(values, "change");
        const significanceValueColumns = this.getValueColumnsByRole(values, "significance");
        const significanceCategoryColumns = this.getCategoryColumnsByRole(categories, "significance");
        const conversionColumns = this.getValueColumnsByRole(values, "Conversion");

        // Every remaining mostly-numeric Measure Data column is a score/value source.
        // Dedicated Change and Significance roles are excluded, so Change can never
        // accidentally become the metric score and text Significance stays text.
        const classified = this.classifyMeasureColumns(values);
        const scoreColumns = classified.numericMeasures.filter((valueColumn) => {
            return !this.sourceHasRole(valueColumn.source, "change") &&
                !this.sourceHasRole(valueColumn.source, "significance") &&
                !this.sourceHasRole(valueColumn.source,"Conversion");
        });

        if (scoreColumns.length === 0) {
            console.warn("[RoleAwareEvidence] No numeric score/value Measure Data column was detected.");
            return [];
        }

        const evidence: MetricEvidence[] = [];

        for (let scoreIndex = 0; scoreIndex < scoreColumns.length; scoreIndex++) {
            const scoreColumn = scoreColumns[scoreIndex];
            const scoreValues = scoreColumn.values || [];
            const scoreMeasureName = this.getColumnDisplayName(scoreColumn.source, "Measure");

            for (let rowIndex = 0; rowIndex < scoreValues.length; rowIndex++) {
                const scoreText = String(scoreValues[rowIndex] == null ? "" : scoreValues[rowIndex]).trim();
                const score = this.parseNumericValue(scoreText);
                if (score === null) { continue; }

                const dimensions: { [key: string]: string } = {};
                for (let dimensionIndex = 0; dimensionIndex < dimensionCategories.length; dimensionIndex++) {
                    const dimension = dimensionCategories[dimensionIndex];
                    const header = this.getColumnDisplayName(
                        dimension.source,
                        "Dimension_" + (dimensionIndex + 1)
                    );
                    const dimensionValues = dimension.values || [];
                    dimensions[header] = rowIndex < dimensionValues.length
                        ? String(dimensionValues[rowIndex] == null ? "" : dimensionValues[rowIndex]).trim()
                        : "";
                }

                const metric = metricCategory
                    ? this.getMetricValue(metricCategory, rowIndex, scoreMeasureName)
                    : scoreMeasureName;

                const changeValue = this.getFirstNumericRoleValueAtRow(changeColumns, rowIndex);
                const conversionValue = this.getFirstNumericRoleValueAtRow(conversionColumns, rowIndex);
                const significance =
                    this.getFirstNonBlankRoleTextAtRow(significanceValueColumns, rowIndex) ||
                    this.getFirstNonBlankCategoryRoleTextAtRow(significanceCategoryColumns, rowIndex);
                const entity = this.buildEntityLabelFromDimensions(dimensions, "") || metric;

                evidence.push({
                    chunkIndex: 0,
                    rowRange: "DataView row " + (rowIndex + 1),
                    entity: entity,
                    metric: String(metric || scoreMeasureName).trim(),
                    measure: scoreMeasureName,
                    score: score,
                    scoreText: scoreText,
                    change: changeValue.value,
                    changeText: changeValue.text,
                    Conversion:conversionValue.value,
                    ConversionText:conversionValue.text,
                    significance: significance,
                    dimensions: dimensions
                });
            }
        }

        const deduped = this.dedupeEvidence(evidence);
        console.log(
            "[RoleAwareEvidence] Extracted " + deduped.length +
            " authoritative row-aligned evidence records. Change role value columns=" + changeColumns.length +
            ", Significance role value columns=" + significanceValueColumns.length +
            ", Significance role category columns=" + significanceCategoryColumns.length + "."
        );
        return deduped;
    }

    private formatEvidenceChangeForGroup(value: number | null, groupValues: MetricEvidence[]): string {
        if (value === null || value === undefined || !isFinite(value)) { return "N/A"; }

        let finiteCount = 0;
        let proportionLikeCount = 0;
        for (let i = 0; i < groupValues.length; i++) {
            const change = groupValues[i].change;
            if (change === null || change === undefined || !isFinite(change)) { continue; }
            finiteCount++;
            if (Math.abs(change) <= 1.5) { proportionLikeCount++; }
        }

        const usePercent = finiteCount > 0 && proportionLikeCount >= Math.ceil(finiteCount * 0.75);
        if (usePercent) {
            return (Math.round(value * 1000) / 10).toFixed(1) + "%";
        }

        const rounded = Math.round(value * 100) / 100;
        return String(rounded);
    }

    private getPromptMentionedDimensionValueSpecs(
        userPrompt: string,
        evidence: MetricEvidence[]
    ): Array<{ header: string; value: string }> {
        const promptToken = " " + this.normalisePromptFieldToken(userPrompt) + " ";
        const available: { [key: string]: { header: string; value: string } } = {};

        for (let i = 0; i < evidence.length; i++) {
            const dimensions = evidence[i].dimensions || {};
            const headers = Object.keys(dimensions);
            for (let h = 0; h < headers.length; h++) {
                const header = headers[h];
                const value = String(dimensions[header] == null ? "" : dimensions[header]).trim();
                if (!value || value.toLowerCase() === "n/a") { continue; }
                const valueToken = this.normalisePromptFieldToken(value);
                if (!valueToken) { continue; }
                if (promptToken.indexOf(" " + valueToken + " ") === -1) { continue; }
                const key = this.normaliseEvidenceKey(header + "__" + value);
                if (!available[key]) { available[key] = { header: header, value: value }; }
            }
        }

        return Object.keys(available).map(function(key) { return available[key]; });
    }

    private buildRoleAwareComparisonEvidence(
        evidence: MetricEvidence[],
        userPrompt: string
    ): string {
        const cleanEvidence = this.dedupeEvidence(evidence || []);
        if (cleanEvidence.length === 0) {
            return "No authoritative role-aware evidence was available.";
        }

        const groups: { [key: string]: MetricEvidence[] } = {};
        const groupLabels: { [key: string]: { metric: string; measure: string } } = {};
        for (let i = 0; i < cleanEvidence.length; i++) {
            const item = cleanEvidence[i];
            const key = this.normaliseEvidenceKey(item.metric + "__" + item.measure);
            if (!groups[key]) {
                groups[key] = [];
                groupLabels[key] = { metric: item.metric, measure: item.measure };
            }
            groups[key].push(item);
        }

        const requestedSubgroups = this.getPromptMentionedDimensionValueSpecs(userPrompt, cleanEvidence);
        const getMetricChangeStrength = function(
            items: MetricEvidence[]
            ): number {
                let strongest = 0;
            
                for (let i = 0; i < items.length; i++) {
                    const change = items[i].change;
                
                    if (
                        change === null ||
                        change === undefined ||
                        !isFinite(change)
                    ) {
                        continue;
                    }
                
                    const magnitude =
                        Math.abs(change);
                
                    if (magnitude > strongest) {
                        strongest = magnitude;
                    }
                }
            
                return strongest;
            };

            const keys =
                Object.keys(groups).sort(
                    function(left, right) {
                        const rightStrength =
                            getMetricChangeStrength(
                                groups[right]
                            );
                        
                        const leftStrength =
                            getMetricChangeStrength(
                                groups[left]
                            );
                        
                        return (
                            rightStrength -
                            leftStrength
                        );
                    }
                );
            
        // const keys = Object.keys(groups).sort(function(left, right) {
        //     return groups[right].length - groups[left].length;
        // });
        const sections: string[] = [];

        const describe = (item: MetricEvidence, group: MetricEvidence[]): string => {
            const score = this.formatEvidenceScoreForGroup(item.score, group);
            const change = item.change !== null && item.change !== undefined && isFinite(item.change)
                ? this.formatEvidenceChangeForGroup(item.change, group)
                : "";
            const Conversion = item.Conversion !== null && item.Conversion !== undefined && isFinite(item.Conversion) ?
            this.formatEvidenceScoreForGroup(item.Conversion,[item]):"";
            const significance = String(item.significance || "").trim();
            let line = this.getEvidenceDisplayEntity(item) +
                "; Metric=" + String(item.metric || "").trim() +
                "; Score=" + score;
            if (change) { line += "; Change=" + change; }
            if (Conversion) { line += "; Conversion=" + Conversion; }
            if (significance) { line += "; Significance=" + significance; }
            return line;
        };

        const buildFacts = (
            label: string,
            items: MetricEvidence[]
        ): string[] => {
            const facts: string[] = [];
            const scoreItems = items.filter(function(item) {
                return item.score !== null && item.score !== undefined && isFinite(item.score as number);
            });
            if (scoreItems.length > 0) {
                const sortedScore = scoreItems.slice().sort(function(a, b) {
                    return (b.score as number) - (a.score as number);
                });
                facts.push(label + " comparable score observations=" + sortedScore.length + ".");
                if (sortedScore.length >= 2) {
                    facts.push(label + " score highest: " + describe(sortedScore[0], scoreItems));
                    facts.push(label + " score lowest: " + describe(sortedScore[sortedScore.length - 1], scoreItems));
                } else {
                    facts.push(label + " has only one score observation, so a score high/low comparison is not valid.");
                }
            }

            const changeItems = items.filter(function(item) {
                return item.change !== null && item.change !== undefined && isFinite(item.change as number);
            });
            if (changeItems.length > 0) {
                facts.push(label + " comparable Change observations=" + changeItems.length + ".");
                const positive = changeItems.filter(function(item) {
                    return (item.change as number) > 0;
                }).sort(function(a, b) {
                    return (b.change as number) - (a.change as number);
                });
                const negative = changeItems.filter(function(item) {
                    return (item.change as number) < 0;
                }).sort(function(a, b) {
                    return (a.change as number) - (b.change as number);
                });
                if (positive.length > 0) {
                    facts.push(label + " highest positive Change: " + describe(positive[0], changeItems));
                }
                if (negative.length > 0) {
                    facts.push(label + " highest decline / most negative Change: " + describe(negative[0], changeItems));
                }
            }
            return facts;
        };

        for (let k = 0; k < keys.length; k++) {
            const key = keys[k];
            const items = groups[key];
            const labelInfo = groupLabels[key];
            const lines: string[] = [];
            lines.push("Exact Metric label: " + labelInfo.metric);
            lines.push("Score/Value Measure Data field: " + labelInfo.measure);
            lines.push.apply(lines, buildFacts("All visible rows for this Metric", items));

            for (let s = 0; s < requestedSubgroups.length; s++) {
                const subgroup = requestedSubgroups[s];
                const subgroupItems = items.filter(function(item) {
                    const dimensions = item.dimensions || {};
                    return String(dimensions[subgroup.header] == null ? "" : dimensions[subgroup.header]).trim() === subgroup.value;
                });
                if (subgroupItems.length === 0) { continue; }
                lines.push.apply(
                    lines,
                    buildFacts(subgroup.header + "=" + subgroup.value, subgroupItems)
                );
            }

            sections.push(lines.join("\n"));
        }

        return sections.join("\n\n");
    }

    private formatScore(value: number | null): string {
        if (value === null || value === undefined || !isFinite(value)) { return "N/A"; }
        const rounded = Math.round(value * 100) / 100;
        return String(rounded);
    }

    private compactEvidenceItem(item: MetricEvidence): string {
        const metric = String(item.metric || "Metric").trim();
        const measure = String(item.measure || "Measure").trim();
        const score = item.score !== null && item.score !== undefined && isFinite(item.score)
            ? this.formatEvidenceScoreForGroup(item.score, [item])
            : "N/A";
        const change = item.change !== null && item.change !== undefined && isFinite(item.change)
            ? this.formatEvidenceChangeForGroup(item.change, [item])
            : "";
        const Conversion = 
            item.Conversion !== null && item.Conversion !== undefined && isFinite(item.Conversion) ?
            this.formatEvidenceScoreForGroup(item.Conversion, [item]) : "";
        const significance = String(item.significance || "").trim();

        let output = this.getEvidenceDisplayEntity(item) +
            " - Metric: " + metric +
            "; Score/Value field: " + measure +
            "; Score: " + score;
        if (change) { output += "; Change: " + change; }
        if (Conversion) { output += "; Conversion: " + Conversion; }
        if (significance) { output += "; Significance: " + significance; }
        return output;
    }

    private buildCompactEvidenceLine(evidence: MetricEvidence[], limit: number): string {
        if (!evidence || evidence.length === 0) { return "No numeric evidence extracted"; }
        const items = evidence.slice(0, limit).map((item) => this.compactEvidenceItem(item));
        const suffix = evidence.length > limit ? " | +" + (evidence.length - limit) + " more" : "";
        return items.join(" | ") + suffix;
    }

    private buildMetricEvidenceSummary(evidence: MetricEvidence[], maxGroups: number, perGroupLimit: number): string {
        if (!evidence || evidence.length === 0) {
            return "No deterministic numeric, Change, or Significance evidence was extracted from the current dataset.";
        }

        const groups: { [key: string]: MetricEvidence[] } = {};
        const groupLabels: { [key: string]: { metric: string; measure: string } } = {};

        for (let i = 0; i < evidence.length; i++) {
            const item = evidence[i];
            if (item.score === null || item.score === undefined || !isFinite(item.score)) { continue; }
            const key = this.normaliseEvidenceKey(item.metric + "__" + item.measure);
            if (!groups[key]) {
                groups[key] = [];
                groupLabels[key] = { metric: item.metric, measure: item.measure };
            }
            groups[key].push(item);
        }

        const keys = Object.keys(groups).sort(function(a, b) {
            return groups[b].length - groups[a].length;
        }).slice(0, maxGroups);

        const sections: string[] = [];

        for (let k = 0; k < keys.length; k++) {
            const key = keys[k];
            const items = groups[key].slice();
            const label = groupLabels[key];
            const scoreItems = items.filter(function(item) {
                return item.score !== null && item.score !== undefined && isFinite(item.score as number);
            }).sort(function(a, b) {
                return (b.score as number) - (a.score as number);
            });
            const changeItems = items.filter(function(item) {
                return item.change !== null && item.change !== undefined && isFinite(item.change as number);
            });

            const lines: string[] = [];
            lines.push("Exact Metric label: " + label.metric);
            lines.push("Score/Value Measure Data field: " + label.measure);
            lines.push("Comparable visible score observations: " + scoreItems.length);

            if (scoreItems.length >= 2) {
                lines.push("Highest score row: " + this.compactEvidenceItem(scoreItems[0]));
                lines.push("Lowest score row: " + this.compactEvidenceItem(scoreItems[scoreItems.length - 1]));
            } else if (scoreItems.length === 1) {
                lines.push("Only score observation: " + this.compactEvidenceItem(scoreItems[0]));
                lines.push("A score high/low comparison is not valid because only one comparable observation is visible.");
            }

            if (changeItems.length > 0) {
                lines.push("Comparable visible Change observations: " + changeItems.length);
                const positiveChanges = changeItems.filter(function(item) {
                    return (item.change as number) > 0;
                }).sort(function(a, b) {
                    return (b.change as number) - (a.change as number);
                });
                const negativeChanges = changeItems.filter(function(item) {
                    return (item.change as number) < 0;
                }).sort(function(a, b) {
                    return (a.change as number) - (b.change as number);
                });

                if (positiveChanges.length > 0) {
                    lines.push("Highest positive Change row: " + this.compactEvidenceItem(positiveChanges[0]));
                }
                if (negativeChanges.length > 0) {
                    lines.push("Highest decline / most negative Change row: " + this.compactEvidenceItem(negativeChanges[0]));
                }
            } else {
                lines.push("No numeric Change value is present for this Metric in the dedicated Change role.");
            }

            const sigRows = items.filter(function(item) {
                return String(item.significance || "").trim().length > 0;
            }).slice(0, Math.max(1, perGroupLimit));
            if (sigRows.length > 0) {
                lines.push(
                    "Examples of exact Significance text paired to the same rows: " +
                    sigRows.map((item) => this.compactEvidenceItem(item)).join("; ")
                );
            }

            sections.push(lines.join("\n"));
        }

        return sections.join("\n\n");
    }

    private removeRowLabelledLines(text: string): string {
        return String(text || "")
            .split("\n")
            .filter(function(line) {
                return !/\bRow\s+\d+\b/i.test(line);
            })
            .join("\n")
            .replace(/\n{3,}/g, "\n\n")
            .trim();
    }


    private isPositiveSignificance(value: string): boolean {
        const text = String(value || "").toLowerCase().trim();
        if (!text) { return false; }
        if (text.indexOf("not significant") !== -1 || text.indexOf("non significant") !== -1 || text.indexOf("insignificant") !== -1) { return false; }
        return text.indexOf("significantly higher") !== -1 ||
            text.indexOf("significant higher") !== -1 ||
            text.indexOf("higher") !== -1 ||
            text.indexOf("increase") !== -1 ||
            text.indexOf("above") !== -1;
    }

    private isNegativeSignificance(value: string): boolean {
        const text = String(value || "").toLowerCase().trim();
        if (!text) { return false; }
        if (text.indexOf("not significant") !== -1 || text.indexOf("non significant") !== -1 || text.indexOf("insignificant") !== -1) { return false; }
        return text.indexOf("significantly lower") !== -1 ||
            text.indexOf("significant lower") !== -1 ||
            text.indexOf("lower") !== -1 ||
            text.indexOf("decrease") !== -1 ||
            text.indexOf("below") !== -1;
    }

    private isAnySignificance(value: string): boolean {
        const text = String(value || "").toLowerCase().trim();
        if (!text) { return false; }
        if (text.indexOf("not significant") !== -1 || text.indexOf("non significant") !== -1 || text.indexOf("insignificant") !== -1) { return false; }
        return text.indexOf("significant") !== -1 ||
            text.indexOf("higher") !== -1 ||
            text.indexOf("lower") !== -1 ||
            text.indexOf("increase") !== -1 ||
            text.indexOf("decrease") !== -1 ||
            text.indexOf("above") !== -1 ||
            text.indexOf("below") !== -1;
    }

    private formatEvidenceScoreForGroup(value: number | null, groupValues: MetricEvidence[]): string {
        if (value === null || value === undefined || !isFinite(value)) { return "N/A"; }

        let finiteCount = 0;
        let proportionLikeCount = 0;
        for (let i = 0; i < groupValues.length; i++) {
            const score = groupValues[i].score;
            if (score === null || score === undefined || !isFinite(score)) { continue; }
            finiteCount++;
            if (Math.abs(score) <= 1.5) { proportionLikeCount++; }
        }

        const usePercent = finiteCount > 0 && proportionLikeCount >= Math.ceil(finiteCount * 0.75);
        if (usePercent) {
            return (Math.round(value * 1000) / 10).toFixed(1) + "%";
        }

        const rounded = Math.round(value * 100) / 100;
        return String(rounded);
    }

    private dedupeEvidence(evidence: MetricEvidence[]): MetricEvidence[] {
        const seen: { [key: string]: boolean } = {};
        const output: MetricEvidence[] = [];
        for (let i = 0; i < (evidence || []).length; i++) {
            const item = evidence[i];
            if (!item || item.score === null || item.score === undefined || !isFinite(item.score)) { continue; }
            const entity = String(item.entity || "").trim();
            const metric = String(item.metric || "").trim();
            if (!entity || !metric || /^row\s+\d+$/i.test(entity)) { continue; }
            const dimensionSignature = Object.keys(item.dimensions || {})
                .sort()
                .map(function(header) {
                    return header + "=" + String(item.dimensions[header] == null ? "" : item.dimensions[header]);
                })
                .join("||");
            const key = this.normaliseEvidenceKey(
                entity + "__" + dimensionSignature + "__" + metric + "__" +
                item.measure + "__" + item.score + "__" + item.change + "__" + item.significance
            );
            if (seen[key]) { continue; }
            seen[key] = true;
            output.push(item);
        }
        return output;
    }


    private getEvidenceContextValues(item: MetricEvidence, headerPattern: RegExp): string[] {
        const values: string[] = [];
        const seen: { [key: string]: boolean } = {};
        const dimensions = item && item.dimensions ? item.dimensions : {};
        const keys = Object.keys(dimensions);

        for (let i = 0; i < keys.length; i++) {
            const header = String(keys[i] || "");
            const value = String(dimensions[header] == null ? "" : dimensions[header]).trim();
            if (!value || value.toLowerCase() === "n/a") { continue; }
            if (!headerPattern.test(header.toLowerCase())) { continue; }
            const key = this.normaliseEvidenceKey(header + "__" + value);
            if (!seen[key]) {
                seen[key] = true;
                values.push(value);
            }
        }

        return values;
    }

    private getEvidenceDisplayEntity(item: MetricEvidence): string {
    const base = String(item && item.entity ? item.entity : "Entity").trim();
    const dimensions = item && item.dimensions ? item.dimensions : {};
    const suffixParts: string[] = [];
    const seen: { [key: string]: boolean } = {};
    const keys = Object.keys(dimensions);
 
    for (let i = 0; i < keys.length; i++) {
        const header = String(keys[i] || "").trim();
        const value = String(dimensions[header] == null ? "" : dimensions[header]).trim();
 
        if (!header || !value || value.toLowerCase() === "n/a") { continue; }
        if (/prompt|instruction/i.test(header)) { continue; }
 
        // Do not repeat the main display entity inside brackets.
        if (this.normaliseEvidenceKey(value) === this.normaliseEvidenceKey(base)) { continue; }
        if (this.normaliseEvidenceKey(header) === this.normaliseEvidenceKey(base)) { continue; }
 
        const part = header + ": " + value;
        const key = this.normaliseEvidenceKey(part);
 
        if (seen[key]) { continue; }
 
        seen[key] = true;
        suffixParts.push(part);
    }
 
    if (suffixParts.length > 0) {
        return base + " (" + suffixParts.join(", ") + ")";
    }
 
    return base;
    }
 

    private buildCurrentContextLine(evidence: MetricEvidence[]): string {
        const buckets: { [label: string]: { [value: string]: boolean } } = {};
        const addValue = function(label: string, value: string): void {
            const cleanLabel = String(label || "").trim();
            const cleanValue = String(value || "").trim();
            if (!cleanLabel || !cleanValue || cleanValue.toLowerCase() === "n/a") { return; }
            if (/prompt|instruction/i.test(cleanLabel)) { return; }
            if (!buckets[cleanLabel]) { buckets[cleanLabel] = {}; }
            buckets[cleanLabel][cleanValue] = true;
        };

        for (let i = 0; i < Math.min(evidence.length, 1000); i++) {
            const item = evidence[i];
            const dimensions = item.dimensions || {};
            const keys = Object.keys(dimensions);

            for (let keyIndex = 0; keyIndex < keys.length; keyIndex++) {
                const key = keys[keyIndex];
                addValue(key, String(dimensions[key] == null ? "" : dimensions[key]));
            }

            if (item.metric) { addValue("Metrics", item.metric); }
            if (item.measure) { addValue("Measures", item.measure); }
        }

        const labels = Object.keys(buckets);
        if (labels.length === 0) { return ""; }

        labels.sort((a: string, b: string) => {
            const preferred = this.normaliseEvidenceKey(this.preferredEntityFieldFromPrompt || "");
            const aPriority = this.normaliseEvidenceKey(a) === preferred ? -1 : 0;
            const bPriority = this.normaliseEvidenceKey(b) === preferred ? -1 : 0;
            if (aPriority !== bPriority) { return aPriority - bPriority; }
            return a.localeCompare(b);
        });

        const parts: string[] = [];
        for (let i = 0; i < labels.length; i++) {
            const label = labels[i];
            const values = Object.keys(buckets[label]);
            const shown = values.slice(0, 8).join(", ");
            const extra = values.length > 8 ? " +" + (values.length - 8) + " more" : "";
            parts.push(label + ": " + shown + extra);
        }

        return "Current selection context:\n" +
            parts.map(function(part) { return "- " + part; }).join("\n");
    }

    private buildClientReadyEvidenceInsight(evidence: MetricEvidence[]): string {
        const cleanEvidence = this.dedupeEvidence(evidence || []);
        if (!cleanEvidence.length) { return ""; }

        const metricGroups: { [key: string]: MetricEvidence[] } = {};
        const metricLabels: { [key: string]: string } = {};
        for (let i = 0; i < cleanEvidence.length; i++) {
            const item = cleanEvidence[i];
            const metric = String(item.metric || item.measure || "Metric").trim();
            const key = this.normaliseEvidenceKey(metric);
            if (!metricGroups[key]) {
                metricGroups[key] = [];
                metricLabels[key] = metric;
            }
            metricGroups[key].push(item);
        }

        // const metricKeys = Object.keys(metricGroups).sort(function(a, b) {
        //     return metricGroups[b].length - metricGroups[a].length;
        // });

        const getMetricPriority = function(
            items: MetricEvidence[]
        ): {
            hasChange: boolean;
            value: number;
        } {
            // ---------------------------------------------
            // 1. Use Change when valid Change data exists
            // ---------------------------------------------
            const validChanges = items
                .filter(function(item) {
                    return (
                        item.change !== null &&
                        item.change !== undefined &&
                        isFinite(item.change as number)
                    );
                })
                .map(function(item) {
                    return Math.abs(item.change as number);
                });
            
            if (validChanges.length > 0) {
                return {
                    hasChange: true,
                    value: Math.max.apply(null, validChanges)
                };
            }
        
            // ---------------------------------------------
            // 2. Fall back to Score when Change is absent
            // ---------------------------------------------
            const validScores = items
                .filter(function(item) {
                    return (
                        item.score !== null &&
                        item.score !== undefined &&
                        isFinite(item.score as number)
                    );
                })
                .map(function(item) {
                    return item.score as number;
                });
            
            if (validScores.length > 0) {
                return {
                    hasChange: false,
                    value: Math.max.apply(null, validScores)
                };
            }
        
            return {
                hasChange: false,
                value: -Infinity
            };
        };

            const metricKeys = Object.keys(metricGroups).sort(
                function(a, b) {
                    const priorityA =
                        getMetricPriority(metricGroups[a]);
                
                    const priorityB =
                        getMetricPriority(metricGroups[b]);
                
                    // If both have Change, strongest absolute Change first.
                    if (
                        priorityA.hasChange &&
                        priorityB.hasChange
                    ) {
                        return priorityB.value - priorityA.value;
                    }
                
                    // If neither has Change, highest Score first.
                    if (
                        !priorityA.hasChange &&
                        !priorityB.hasChange
                    ) {
                        return priorityB.value - priorityA.value;
                    }
                
                    // Metrics with valid Change come before score-only metrics.
                    return priorityA.hasChange ? -1 : 1;
                }
            );
 

        const formatItem = (item: MetricEvidence, group: MetricEvidence[]): string => {
            const score = this.formatEvidenceScoreForGroup(item.score, group);
            const sig = String(item.significance || "").trim();
            return this.getEvidenceDisplayEntity(item) + " - " + score + (sig ? " (" + sig + ")" : "");
        };

        const significantLines: string[] = [];
        for (let k = 0; k < metricKeys.length; k++) {
            const key = metricKeys[k];
            const group = metricGroups[key].slice().sort(function(a, b) {
                return (b.score == null ? -Infinity : b.score) - (a.score == null ? -Infinity : a.score);
            });
            const sigItems = group.filter((item) => this.isAnySignificance(item.significance));
            if (!sigItems.length) { continue; }
            const shown = sigItems.slice(0, 8).map((item) => formatItem(item, group));
            const extra = sigItems.length > shown.length ? " plus " + (sigItems.length - shown.length) + " more" : "";
            significantLines.push("- " + metricLabels[key] + ": " + shown.join("; ") + extra);
        }

        const highestLowestLines: string[] = [];
        for (let k = 0; k < metricKeys.length; k++) {
            const key = metricKeys[k];
            const group = metricGroups[key].slice().sort(function(a, b) {
                return (b.score == null ? -Infinity : b.score) - (a.score == null ? -Infinity : a.score);
            });
            if (!group.length) { continue; }
            if (group.length === 1) {
                const onlyItem = group[0];
                highestLowestLines.push(
                    "- " + metricLabels[key] + ": only observed " +
                    this.getEvidenceDisplayEntity(onlyItem) + " " +
                    this.formatEvidenceScoreForGroup(onlyItem.score, group) +
                    "; a high/low comparison is not available from one comparable observation"
                );
                continue;
            }
            const highest = group[0];
            const lowest = group[group.length - 1];
            const nextItems = group.slice(1, Math.min(group.length, 4)).map((item) => this.getEvidenceDisplayEntity(item) + " " + this.formatEvidenceScoreForGroup(item.score, group));
            highestLowestLines.push(
                "- " + metricLabels[key] + ": highest " + this.getEvidenceDisplayEntity(highest) + " " + this.formatEvidenceScoreForGroup(highest.score, group) +
                (nextItems.length ? "; next " + nextItems.join(", ") : "") +
                "; lowest " + this.getEvidenceDisplayEntity(lowest) + " " + this.formatEvidenceScoreForGroup(lowest.score, group)
            );
        }

        const entityGroups: { [key: string]: MetricEvidence[] } = {};
        const entityLabels: { [key: string]: string } = {};
        for (let i = 0; i < cleanEvidence.length; i++) {
            const item = cleanEvidence[i];
            const entityKey = this.normaliseEvidenceKey(this.getEvidenceDisplayEntity(item));
            if (!entityGroups[entityKey]) {
                entityGroups[entityKey] = [];
                entityLabels[entityKey] = this.getEvidenceDisplayEntity(item);
            }
            entityGroups[entityKey].push(item);
        }

        const entityKeys = Object.keys(entityGroups).sort(function(a, b) {
            const avg = function(items: MetricEvidence[]): number {
                let sum = 0;
                let count = 0;
                for (let i = 0; i < items.length; i++) {
                    if (items[i].score !== null && items[i].score !== undefined && isFinite(items[i].score as number)) {
                        sum += items[i].score as number;
                        count++;
                    }
                }
                return count ? sum / count : -Infinity;
            };
            return avg(entityGroups[b]) - avg(entityGroups[a]);
        }).slice(0, 6);

        const brandLines: string[] = [];
        for (let e = 0; e < entityKeys.length; e++) {
            const entityKey = entityKeys[e];
            const items = entityGroups[entityKey].slice().sort(function(a, b) {
                return String(a.metric).localeCompare(String(b.metric));
            });
            const parts: string[] = [];
            for (let i = 0; i < Math.min(items.length, 6); i++) {
                const item = items[i];
                const metricKey = this.normaliseEvidenceKey(item.metric);
                const group = metricGroups[metricKey] || items;
                parts.push(item.metric + " " + this.formatEvidenceScoreForGroup(item.score, group));
            }
            brandLines.push("- " + entityLabels[entityKey] + ": " + parts.join(", ") + ".");
        }

        const opportunityLines: string[] = [];
        for (let k = 0; k < metricKeys.length; k++) {
            const key = metricKeys[k];
            const group = metricGroups[key].slice().sort(function(a, b) {
                return (a.score == null ? Infinity : a.score) - (b.score == null ? Infinity : b.score);
            });
            if (group.length > 1) {
                opportunityLines.push("- Improve lower performers on " + metricLabels[key] + ", led by " + this.getEvidenceDisplayEntity(group[0]) + " (" + this.formatEvidenceScoreForGroup(group[0].score, group) + ").");
            }
            if (opportunityLines.length >= 4) { break; }
        }

        const riskLines: string[] = [];
        const entityDropLines: string[] = [];
        for (let e = 0; e < entityKeys.length; e++) {
            const entityKey = entityKeys[e];
            const items = entityGroups[entityKey].filter(function(item) { return item.score !== null && item.score !== undefined && isFinite(item.score as number); })
                .sort(function(a, b) { return (b.score as number) - (a.score as number); });
            if (items.length >= 2) {
                const high = items[0];
                const low = items[items.length - 1];
                const highGroup = metricGroups[this.normaliseEvidenceKey(high.metric)] || items;
                const lowGroup = metricGroups[this.normaliseEvidenceKey(low.metric)] || items;
                entityDropLines.push("- " + entityLabels[entityKey] + " shows a funnel gap from " + high.metric + " " + this.formatEvidenceScoreForGroup(high.score, highGroup) + " to " + low.metric + " " + this.formatEvidenceScoreForGroup(low.score, lowGroup) + ".");
            }
            if (entityDropLines.length >= 3) { break; }
        }
        riskLines.push.apply(riskLines, entityDropLines);
        if (!riskLines.length) { riskLines.push("- No major risk can be confirmed beyond the visible metric gaps."); }

        const hasSignificance = significantLines.length > 0;
        const contextLine = this.buildCurrentContextLine(cleanEvidence);

        return (
            (contextLine ? contextLine + "\n\n" : "") +
            "Summary of Significant Changes:\n" +
            (hasSignificance ? significantLines.join("\n") : "- No Significantly Higher/Lower items are visible in the current selection.") +
            "\n\nPerformance Analysis:\n" +
            (brandLines.length ? brandLines.join("\n") : "- Entity-level performance could not be grouped from the current fields.") +
            "\n\nHighest and Lowest Analysis:\n" +
            (highestLowestLines.length ? highestLowestLines.join("\n") : "- Numeric high/low evidence is not visible in the current selection.") +
            "\n\nKey Risks:\n" +
            riskLines.join("\n") +
            "\n\nKey Opportunities:\n" +
            (opportunityLines.length ? opportunityLines.join("\n") : "- No specific opportunity can be confirmed beyond the visible scores.") +
            "\n\nCritical Anomalies:\n" +
            "- No critical anomaly is confirmed from the deterministic evidence; review unusually large gaps between highest and lowest performers."
        );
    }

    private removeContradictoryEvidenceLines(text: string): string {
        const lower = String(text || "").toLowerCase();
        if (lower.indexOf("no significant") !== -1 ||
            lower.indexOf("not visible in the summarised evidence") !== -1 ||
            lower.indexOf("score not visible") !== -1 ||
            /\brow\s+\d+\b/i.test(text)) {
            return "";
        }
        return String(text || "").trim();
    }

    private ensureEvidenceSectionsInFinalInsight(finalInsight: string, evidenceSummary: string): string {
        return this.removeRowLabelledLines(String(finalInsight || "").trim());
    }

    private buildDeterministicExecutiveSummary(
        groupSummaries: GroupSummary[],
        totalRows: number,
        totalChunks: number
    ): string {
        const recurring = this.takeUnique(groupSummaries.reduce(function(all: string[], group: GroupSummary) {
            return all.concat(group.recurringPatterns || []);
        }, []), 6);
        const findings = this.takeUnique(groupSummaries.reduce(function(all: string[], group: GroupSummary) {
            return all.concat(group.mergedFindings || []);
        }, []), 8);
        const signals = this.takeUnique(groupSummaries.reduce(function(all: string[], group: GroupSummary) {
            return all.concat(group.prioritisedSignals || []);
        }, []), 6);

        const section = function(title: string, values: string[], fallback: string): string {
            const bullets = values.length > 0 ? values : [fallback];
            return title + ":\n" + bullets.map(function(v) { return "- " + v; }).join("\n");
        };

        return (
            "Dataset Coverage:\n" +
            "- Analysed " + totalRows + " rows across " + totalChunks + " chunks before final insight generation.\n\n" +
            section("Overall Patterns", recurring, "No recurring pattern was visible in the summarised evidence.") + "\n\n" +
            section("Top KPI Signals", signals, "No KPI signal was visible in the summarised evidence.") + "\n\n" +
            section("Key Risks", findings.slice(0, 4), "No material risk was visible in the summarised evidence.") + "\n\n" +
            section("Key Opportunities", findings.slice(4, 8), "No material opportunity was visible in the summarised evidence.") + "\n\n" +
            section("Critical Anomalies", [], "No critical anomaly was visible in the summarised evidence.")
        );
    }

    private buildDeterministicFinalInsight(executiveSummary: string): string {
        const evidenceBullets = String(executiveSummary || "")
            .replace(/\r\n/g, "\n")
            .replace(/\r/g, "\n")
            .split("\n")
            .map(function(line) {
                return line.replace(/^[-*-]\s+/, "").trim();
            })
            .filter(function(line) {
                return line.length > 0 && !/^[A-Za-z][A-Za-z0-9\s\-/&()]{2,}:$/.test(line);
            });

        const uniqueBullets = this.takeUnique(evidenceBullets, 12);
        if (uniqueBullets.length === 0) {
            uniqueBullets.push("No deterministic evidence was available for a fallback response.");
        }

        const headings = this.lastKnownHeadingsFromPrompt || [];
        if (headings.length === 0) {
            return uniqueBullets.map(function(value) { return "- " + value; }).join("\n");
        }

        const output: string[] = [];
        let cursor = 0;
        for (let headingIndex = 0; headingIndex < headings.length; headingIndex++) {
            const heading = headings[headingIndex];
            const key = this.normalisePromptFieldToken(heading);
            const requestedLimit = this.lastKnownBulletLimitsFromPrompt[key];
            const remainingHeadings = Math.max(1, headings.length - headingIndex);
            const remainingBullets = Math.max(0, uniqueBullets.length - cursor);
            const count = requestedLimit && requestedLimit > 0
                ? requestedLimit
                : Math.max(1, Math.ceil(remainingBullets / remainingHeadings));

            output.push(heading);
            for (let i = 0; i < count && cursor < uniqueBullets.length; i++, cursor++) {
                output.push(this.getPromptListMarker(heading, i) + " " + uniqueBullets[cursor]);
            }
            output.push("");
        }

        return output.join("\n").trim();
    }


    // -------------------------------------------------------------------------
    // ORCHESTRATOR - runChunkedAnalysis
    // -------------------------------------------------------------------------

    private async runChunkedAnalysis(userPrompt: string, selectionKey: string, basePromptIsDefault: boolean): Promise<void> {
        if (this.isLoadingInsights && this.inFlightSelectionKey === selectionKey) { return; }

        const requestId = ++this.latestRequestId;
        this.inFlightSelectionKey = selectionKey;
        this.isLoadingInsights = true;
        this.currentUiMessage = "";
        this.currentProgressMessage = "Preparing dataset...";
        this.updateShowButtonState(true);
        this.renderCurrentResponse();

        const chunkMemory: ChunkSummary[] = [];

        try {
            this.updateProgressMessage("Preparing dataset...");
            const fullCsvText = this.formatFullPivotedData();

            if (this.isStructuralMessage(fullCsvText)) {
                if (requestId !== this.latestRequestId) { return; }
                this.gptResponse = "";
                this.currentUiMessage = fullCsvText;
                this.currentUiMessageKind = "error";
                this.isLoadingInsights = false;
                this.inFlightSelectionKey = "";
                this.updateShowButtonState(false);
                this.renderCurrentResponse();
                return;
            }

            const chunks = this.splitCsvIntoChunks(fullCsvText);

            if (chunks.length === 0) {
                if (requestId !== this.latestRequestId) { return; }
                this.gptResponse = "";
                this.currentUiMessage = "No data rows were found for the current selection.";
                this.currentUiMessageKind = "error";
                this.isLoadingInsights = false;
                this.inFlightSelectionKey = "";
                this.updateShowButtonState(false);
                this.renderCurrentResponse();
                return;
            }

            // console.log("[ChunkingLayer] Total chunks: " + chunks.length);

            // let currentDataRow = 1;

            // for (let i = 0; i < chunks.length; i++) {
            //     if (requestId !== this.latestRequestId) { return; }

            //     this.updateProgressMessage(
            //         chunks.length > 1
            //             // ? "Analysing data " + (i + 1) + " of " + chunks.length + "..."
            //             ?"Summarizing Data..."
            //             : "Analysing selected data..."
            //     );

            //     const summary = await this.analyseChunk(chunks[i], i, chunks.length, currentDataRow);
            //     this.storeChunkSummary(chunkMemory, summary);

            //     const dataLinesInChunk = chunks[i]
            //         .split("\n")
            //         .filter(function(line) {
            //             return line.trim().length > 0 && line.indexOf("NOTE:") !== 0;
            //         }).length - 1;

            //     currentDataRow += Math.max(0, dataLinesInChunk);
            // }

            // if (requestId !== this.latestRequestId) { return; }

            // const groupSummaries = await this.runAggregationLayer(chunkMemory);

console.log(
    "[ChunkingLayer] Total chunks: " +
    chunks.length
);
 
let currentDataRow = 1;
 
// Build jobs first so every chunk keeps
// exactly the same row numbering as before.
const chunkJobs: Array<{
    chunkText: string;
    chunkIndex: number;
    firstDataRow: number;
}> = [];
 
for (let i = 0; i < chunks.length; i++) {
 
    chunkJobs.push({
        chunkText: chunks[i],
        chunkIndex: i,
        firstDataRow: currentDataRow
    });
 
    const dataLinesInChunk = chunks[i]
        .split("\n")
        .filter(function(line) {
            return (
                line.trim().length > 0 &&
                line.indexOf("NOTE:") !== 0
            );
        }).length - 1;
 
    currentDataRow += Math.max(
        0,
        dataLinesInChunk
    );
}
 
this.updateProgressMessage(
    chunks.length > 1
        ? "Summarizing Data..."
        : "Analysing selected data..."
);
 
// Results will still be stored
// in original chunk order.
const orderedChunkSummaries: ChunkSummary[] =
    new Array(chunkJobs.length);
 
let nextChunkJob = 0;
 
const chunkWorker =
    async (): Promise<void> => {
 
    while (true) {
 
        if (
            requestId !==
            this.latestRequestId
        ) {
            return;
        }
 
        const jobIndex =
            nextChunkJob++;
 
        if (
            jobIndex >=
            chunkJobs.length
        ) {
            return;
        }
 
        const job =
            chunkJobs[jobIndex];
 
        orderedChunkSummaries[jobIndex] =
            await this.analyseChunk(
                job.chunkText,
                job.chunkIndex,
                chunks.length,
                job.firstDataRow
            );
    }
};
 
const chunkWorkerCount = Math.min(
    MAX_PARALLEL_AZURE_CALLS,
    chunkJobs.length
);
 
await Promise.all(
    Array.from(
        { length: chunkWorkerCount },
        () => chunkWorker()
    )
);
 
if (
    requestId !==
    this.latestRequestId
) {
    return;
}
 
// Store summaries in exactly the same
// order as your old implementation.
for (
    let i = 0;
    i < orderedChunkSummaries.length;
    i++
) {
    this.storeChunkSummary(
        chunkMemory,
        orderedChunkSummaries[i]
    );
}
 
if (
    requestId !==
    this.latestRequestId
) {
    return;
}
 
const groupSummaries =
    await this.runAggregationLayer(
        chunkMemory
    );
 

            if (requestId !== this.latestRequestId) { return; }

            // this.updateProgressMessage("Building executive summary across all data...");
            this.updateProgressMessage("Summarizing the data.....");

            const totalDataRows = currentDataRow - 1;
            const executiveSummary = await this.buildExecutiveSummary(
                groupSummaries,
                totalDataRows,
                chunks.length
            );

            if (requestId !== this.latestRequestId) { return; }

            // this.updateProgressMessage("Generating final insights from the complete dataset...");
            this.updateProgressMessage("Summarizing the data.......");

            const userFacingPrompt = userPrompt.trim();
            const roleAwareMetricEvidence = this.extractRoleAwareMetricEvidenceFromDataView();
            const completeMetricEvidence = roleAwareMetricEvidence.length > 0
                ? roleAwareMetricEvidence
                : this.collectMetricEvidence(groupSummaries);
            const evidenceSummary = this.buildMetricEvidenceSummary(
                completeMetricEvidence,
                18,
                5
            );
            const finalInsight = await this.generateFinalInsight(
                userFacingPrompt,
                executiveSummary,
                evidenceSummary,
                completeMetricEvidence,
                basePromptIsDefault
            );

            if (requestId !== this.latestRequestId) { return; }

            this.gptResponse = finalInsight;
            this.currentUiMessage = "";

            if (this.isValidAiResponse(finalInsight)) {
                this.storeCachedResponse(selectionKey, finalInsight);
            }

            this.hasSelectionChanged = false;
            this.isLoadingInsights = false;
            this.inFlightSelectionKey = "";
            this.updateShowButtonState(false);
            this.renderCurrentResponse();
        } catch (error) {
            console.error("Error in runChunkedAnalysis:", error);
            if (requestId !== this.latestRequestId) { return; }
            this.isLoadingInsights = false;
            this.inFlightSelectionKey = "";
            this.updateShowButtonState(false);
            this.gptResponse = "";
            this.currentUiMessage = this.getCleanErrorMessage(error);
            this.currentUiMessageKind = "error";
            this.renderCurrentResponse();
        }
    }

    // -------------------------------------------------------------------------
    // AZURE OPENAI CALL WITH RETRY
    // FIX 1 - callType parameter selects system prompt and temperature.
    // FIX 2 - JSON call types use temperature 0; text types use temperature 1.
    // FIX 3 - maxTokens is now passed per call type by the caller.
    // -------------------------------------------------------------------------

    private async callAzureOpenAIWithRetry(
        prompt: string,
        maxTokens: number,
        attempt: number,
        callType: AzureCallType,
        flexibleTextFinal?: boolean
    ): Promise<string> {
        const apiKey = (config as any).OPENAI_API_KEY;
        if (!apiKey) { throw new Error("API key not found in config"); }

        const apiEndpoint    = "https://openai-coe-poc.openai.azure.com";
        const deploymentName = "gpt-5-mini";

        // Azure OpenAI GPT-5 deployments are more reliable through the v1 Responses API.
        // The previous Chat Completions call used max_completion_tokens and temperature.
        // For reasoning models, max_completion_tokens includes hidden reasoning tokens, so
        // small values can produce an apparently empty message even when the API succeeded.
        // Responses API uses max_output_tokens for visible output and supports reasoning effort.
        const apiUrl = apiEndpoint + "/openai/v1/responses";

        let instructions: string;
        let reasoningEffort: string;
        let verbosity: string;

        if (callType === "json_extraction" || callType === "json_aggregation") {
            instructions =
                "You are a deterministic JSON extraction engine. " +
                "Return ONLY one valid JSON object. No markdown fences. No prose. No explanation.";
            reasoningEffort = "minimal";
            verbosity       = "low";
        } else if (callType === "text_executive") {
            instructions =
                "You are a senior business analyst. Write concise structured summaries using plain labelled sections and bullet points.";
            reasoningEffort = "minimal";
            verbosity       = "low";
        } else if (callType === "text_final") {
            instructions =
                "You are a senior data analyst. Follow the user's prompt exactly. Do not impose predefined headings, section names, bullet counts, or a fixed structure unless the user's prompt asks for them. Keep the output concise and data-driven.";
            reasoningEffort = "low";
            verbosity       = "medium";
        } else {
            instructions =
                "You are a senior business analyst. Follow the requested task structure and keep the output concise and data-driven.";
            reasoningEffort = "low";
            verbosity       = "medium";
        }

        const payload: any = {
            model: deploymentName,
            instructions: instructions,
            input: prompt,
            max_output_tokens: maxTokens,
            reasoning: {
                effort: reasoningEffort
            },
            text: {
                verbosity: verbosity
            }
        };

        // let response: Response;

        // try {
        //     response = await fetch(apiUrl, {
        //         method: "POST",
        //         headers: {
        //             "Content-Type": "application/json",
        //             "api-key": apiKey
        //         },
        //         body: JSON.stringify(payload)
        //     });
        // } catch (networkError) {
        //     if (attempt < MAX_RETRY_ATTEMPTS) {
        //         const delay = RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);
        //         console.warn(
        //             "[AzureCall] Network error on attempt " + attempt +
        //             " (type=" + callType + "). Retrying in " + delay + "ms.", networkError
        //         );
        //         await this.sleep(delay);
        //         return this.callAzureOpenAIWithRetry(prompt, maxTokens, attempt + 1, callType, flexibleTextFinal);
        //     }
        //     throw networkError;
        // }

        let response: Response;
 
const controller = new AbortController();
 
this.activeAzureControllers.add(controller);
 
try {
 
    response = await fetch(apiUrl, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "api-key": apiKey
        },
        body: JSON.stringify(payload),
 
        // Allows requests from an old Power BI
        // selection to be stopped immediately.
        signal: controller.signal
    });
 
} catch (networkError) {
 
    const errorName =
        networkError &&
        (networkError as any).name
            ? String((networkError as any).name)
            : "";
 
    // IMPORTANT:
    // An intentional cancellation should NOT
    // enter the retry mechanism.
    if (errorName === "AbortError") {
        throw networkError;
    }
 
    if (attempt < MAX_RETRY_ATTEMPTS) {
 
        const delay =
            RETRY_BASE_DELAY_MS *
            Math.pow(2, attempt - 1);
 
        console.warn(
            "[AzureCall] Network error on attempt " +
            attempt +
            " (type=" +
            callType +
            "). Retrying in " +
            delay +
            "ms.",
            networkError
        );
 
        await this.sleep(delay);
 
        return this.callAzureOpenAIWithRetry(
            prompt,
            maxTokens,
            attempt + 1,
            callType,
            flexibleTextFinal
        );
    }
 
    throw networkError;
 
} finally {
 
    this.activeAzureControllers.delete(
        controller
    );
}
 

        const rawResponse = await response.text();

        if (!response.ok) {
            const isRetryable = response.status === 408 || response.status === 409 || response.status === 429 || response.status >= 500;

            if (isRetryable && attempt < MAX_RETRY_ATTEMPTS) {
                const retryAfterHeader  = response.headers.get("Retry-After");
                const retryAfterSeconds = retryAfterHeader ? parseInt(retryAfterHeader, 10) : 0;
                const delay             = retryAfterSeconds > 0
                    ? retryAfterSeconds * 1000
                    : RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);

                console.warn(
                    "[AzureCall] HTTP " + response.status +
                    " on attempt " + attempt + " (type=" + callType +
                    "). Retrying in " + delay + "ms. Response: " + rawResponse.substring(0, 500)
                );
                await this.sleep(delay);
                return this.callAzureOpenAIWithRetry(prompt, maxTokens, attempt + 1, callType, flexibleTextFinal);
            }

            // Helpful fallback: if this Azure resource has not enabled /openai/v1 yet,
            // retry the same call once through the dated Chat Completions endpoint, but
            // without unsupported temperature/top_p parameters and with reasoning_effort.
            if ((response.status === 404 || response.status === 400) && attempt === 1) {
                console.warn(
                    "[AzureCall] Responses API failed with HTTP " + response.status +
                    ". Trying legacy Chat Completions compatibility path once. Response: " + rawResponse.substring(0, 500)
                );
                return this.callAzureChatCompletionsFallback(prompt, maxTokens, callType);
            }

            throw new Error("HTTP error! Status: " + response.status + ", Response: " + rawResponse);
        }

        let data: any;

        try {
            data = JSON.parse(rawResponse);
        } catch (parseError) {
            throw new Error("Unable to parse AI response JSON");
        }

        console.log(
            "[AzureCall] Response received (attempt=" + attempt +
            ", type=" + callType + "): ", rawResponse.substring(0, 400)
        );

        const parsedText = this.extractTextFromAiResponse(data);

        const responseStatus = String(data && data.status ? data.status : "").toLowerCase();
        const incompleteReason = data && data.incomplete_details && data.incomplete_details.reason
            ? String(data.incomplete_details.reason)
            : "";

        if (responseStatus === "incomplete") {
            console.warn(
                "[AzureCall] Responses API returned incomplete output" +
                (incompleteReason ? " (reason=" + incompleteReason + ")" : "") +
                " on attempt " + attempt + "."
            );

            if (attempt < MAX_RETRY_ATTEMPTS && maxTokens < 8000) {
                const expandedTokenBudget = Math.min(Math.max(maxTokens * 2, 6000), 8000);
                const delay = RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);
                await this.sleep(delay);
                return this.callAzureOpenAIWithRetry(
                    prompt,
                    expandedTokenBudget,
                    attempt + 1,
                    callType,
                    flexibleTextFinal
                );
            }
            // If the service is already at the maximum budget, return the partial text
            // to the final-layer validator. That layer will issue a concise full replacement
            // request instead of displaying/caching a cut-off sentence.
        }

        if (!parsedText) {
            console.warn(
                "[AzureCall] Empty content returned (attempt=" + attempt +
                ", type=" + callType + "). Full response:", data
            );

            if (attempt < MAX_RETRY_ATTEMPTS) {
                const expandedTokenBudget = Math.min(maxTokens * 2, 8000);
                const delay = RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);
                console.warn(
                    "[AzureCall] Retrying empty output with larger visible output budget " +
                    expandedTokenBudget + " after " + delay + "ms."
                );
                await this.sleep(delay);
                return this.callAzureOpenAIWithRetry(prompt, expandedTokenBudget, attempt + 1, callType, flexibleTextFinal);
            }

            throw new Error("AI service returned an empty message");
        }

        return parsedText;
    }

    private async callAzureChatCompletionsFallback(
        prompt: string,
        maxTokens: number,
        callType: AzureCallType
    ): Promise<string> {
        const apiKey = (config as any).OPENAI_API_KEY;
        if (!apiKey) { throw new Error("API key not found in config"); }

        const apiEndpoint    = "https://openai-coe-poc.openai.azure.com";
        const apiVersion     = "2024-12-01-preview";
        const deploymentName = "gpt-5-mini";
        const apiUrl =
            apiEndpoint +
            "/openai/deployments/" +
            deploymentName +
            "/chat/completions?api-version=" +
            apiVersion;

        let systemPrompt: string;
        let reasoningEffort: string;

        if (callType === "json_extraction" || callType === "json_aggregation") {
            systemPrompt = "You are a JSON data extraction engine. Output ONLY valid JSON. No prose. No markdown fences.";
            reasoningEffort = "minimal";
        } else if (callType === "text_executive") {
            systemPrompt = "You are a senior business analyst. Write concise structured summaries using labelled sections and bullets.";
            reasoningEffort = "minimal";
        } else {
            systemPrompt = "Follow the user's requested output format exactly. Do not impose predefined headings or section names that the user did not request.";
            reasoningEffort = "low";
        }

        const payload: any = {
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: prompt }
            ],
            max_completion_tokens: Math.min(maxTokens * 2, 8000),
            reasoning_effort: reasoningEffort
        };

        const response = await fetch(apiUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "api-key": apiKey
            },
            body: JSON.stringify(payload)
        });

        const rawResponse = await response.text();
        if (!response.ok) {
            throw new Error("HTTP error! Status: " + response.status + ", Response: " + rawResponse);
        }

        let data: any;
        try {
            data = JSON.parse(rawResponse);
        } catch (parseError) {
            throw new Error("Unable to parse AI response JSON");
        }

        const parsedText = this.extractTextFromAiResponse(data);
        if (!parsedText) {
            console.warn("[AzureCallFallback] Empty content returned. Full response:", data);
            throw new Error("AI service returned an empty message");
        }

        return parsedText;
    }

    private sleep(ms: number): Promise<void> {
        return new Promise(function(resolve) { setTimeout(resolve, ms); });
    }


    // -------------------------------------------------------------------------
    // Role helpers - these keep the implementation dynamic and prevent numeric
    // Category Data such as AgeBands/Education/Gender from becoming scores.
    // -------------------------------------------------------------------------

    private sourceHasRole(source: powerbi.DataViewMetadataColumn | undefined, roleName: string): boolean {
        if (!source || !source.roles) { return false; }
        const roles: any = source.roles as any;
        if (roles[roleName] === true) { return true; }

        const requested = String(roleName || "").toLowerCase();
        const roleNames = Object.keys(roles);
        for (let i = 0; i < roleNames.length; i++) {
            const current = roleNames[i];
            if (roles[current] === true && String(current).toLowerCase() === requested) {
                return true;
            }
        }
        return false;
    }

    /**
     * Returns the analytical Category Data fields at the original Power BI row
     * grain. A field can be assigned to both Category Data and Highlight Entity.
     * Power BI may then expose duplicate category columns with the same queryName.
     * We keep the best analytical copy and ignore a Highlight-Entity-only copy.
     */
    private getAnalyticalCategoryColumns(
        categories: powerbi.DataViewCategoryColumn[]
    ): powerbi.DataViewCategoryColumn[] {
        const selected: {
            [key: string]: {
                column: powerbi.DataViewCategoryColumn;
                score: number;
                order: number;
            }
        } = {};
        const orderedKeys: string[] = [];

        for (let i = 0; i < (categories || []).length; i++) {
            const category = categories[i];
            if (!category || this.isLikelyPromptCategory(category) || this.isLikelyMetricCategory(category)) {
                continue;
            }

            // Dedicated semantic roles are evidence fields, not analytical dimensions.
            // Significance may arrive as a text category (for example High Sig / Low Sig / No Significance).
            if (this.sourceHasRole(category.source, "significance") ||
                this.sourceHasRole(category.source, "change")) {
                continue;
            }

            const source: any = category.source || {};
            const roles: any = source.roles || {};
            const roleNames = Object.keys(roles).filter(function(roleName) {
                return roles[roleName] === true;
            });
            const hasAnalyticalRole = roleNames.some(function(roleName) {
                return roleName !== "highlightEntity" && roleName !== "promptText" && roleName !== "metricCategory";
            });
            const isHighlightOnly = roleNames.length > 0 &&
                roleNames.indexOf("highlightEntity") !== -1 &&
                !hasAnalyticalRole;

            if (isHighlightOnly) { continue; }

            const key = this.getCategoryIdentityKey(category, i);
            const values = category.values || [];
            let nonBlankCount = 0;
            for (let valueIndex = 0; valueIndex < values.length; valueIndex++) {
                const text = String(values[valueIndex] == null ? "" : values[valueIndex]).trim();
                if (text) { nonBlankCount++; }
            }

            let score = 0;
            if (this.sourceHasRole(category.source, "categoryData")) { score += 1000; }
            if (hasAnalyticalRole) { score += 250; }
            if (!this.sourceHasRole(category.source, "highlightEntity")) { score += 25; }
            score += Math.min(nonBlankCount, 200);
            score += Math.min(values.length, 100);

            if (!selected[key]) {
                selected[key] = { column: category, score: score, order: i };
                orderedKeys.push(key);
            } else if (score > selected[key].score) {
                selected[key].column = category;
                selected[key].score = score;
            }
        }

        orderedKeys.sort(function(left, right) {
            return selected[left].order - selected[right].order;
        });

        return orderedKeys.map(function(key) {
            return selected[key].column;
        });
    }

    private getAnalyticalDimensionDescriptors(
        categories: powerbi.DataViewCategoryColumn[]
    ): AnalyticalDimensionDescriptor[] {
        const columns = this.getAnalyticalCategoryColumns(categories || []);
        const result: AnalyticalDimensionDescriptor[] = [];

        for (let i = 0; i < columns.length; i++) {
            const name = this.getColumnDisplayName(columns[i].source, "Dimension_" + (i + 1));
            const values = this.getUniqueNonBlankValues(
                columns[i].values,
                this.MAX_CONTEXT_VALUES_PER_FIELD
            ).filter(function(value) {
                return value !== "... more selected";
            });
            result.push({
                name: name,
                values: values,
                isComparison: values.length > 1
            });
        }

        return result;
    }

    private isDemographicDimensionName(fieldName: string): boolean {
        const token = this.normalisePromptFieldToken(fieldName);
        return /(^| )(gender|sex|age|age band|ageband|generation|income|education|race|ethnicity|sec|social class|demographic|household|marital|occupation)( |$)/.test(token);
    }

    private getPromptRequestedDimensionNames(
        promptInstruction: string,
        descriptors: AnalyticalDimensionDescriptor[]
    ): string[] {
        const requested: string[] = [];
        const seen: { [key: string]: boolean } = {};
        const promptToken = " " + this.normalisePromptFieldToken(promptInstruction) + " ";
        const asksForDemographics =
            promptToken.indexOf(" demographic ") !== -1 ||
            promptToken.indexOf(" demographics ") !== -1 ||
            promptToken.indexOf(" demographic breakdown ") !== -1;

        for (let i = 0; i < descriptors.length; i++) {
            const descriptor = descriptors[i];
            const mentioned = this.promptMentionsField(promptInstruction, descriptor.name) ||
                (asksForDemographics && this.isDemographicDimensionName(descriptor.name));
            if (!mentioned) { continue; }
            const key = this.normalisePromptFieldToken(descriptor.name);
            if (!seen[key]) {
                seen[key] = true;
                requested.push(descriptor.name);
            }
        }

        return requested;
    }

    private buildAnalyticalGrainInstruction(
        descriptors: AnalyticalDimensionDescriptor[],
        userPrompt: string
    ): string {
        const fields = descriptors.map(function(item) { return item.name; });
        const comparison = descriptors.filter(function(item) { return item.isComparison; });
        const context = descriptors.filter(function(item) { return !item.isComparison; });
        const requested = this.getPromptRequestedDimensionNames(userPrompt, descriptors);
        const formatItem = function(item: AnalyticalDimensionDescriptor): string {
            const shown = item.values.slice(0, 12).join(", ");
            const extra = item.values.length > 12 ? " +" + (item.values.length - 12) + " more" : "";
            return item.name + (shown ? " = " + shown + extra : "");
        };

        let text =
            "Analytical row grain: " +
            (fields.length > 0 ? fields.join(" + ") : "all visible Category Data fields") +
            " + Metric.\n" +
            "Rows were averaged only when every listed Category Data value and the Metric were identical.\n" +
            "Never average, merge, or collapse rows when ANY Category Data value differs. Male and Female, age groups, income groups, education groups, markets, categories, brands, and time periods are separate observations.\n";

        if (comparison.length > 0) {
            text += "Comparison dimensions with multiple visible values:\n" +
                comparison.map(function(item) { return "- " + formatItem(item); }).join("\n") + "\n";
        }
        if (context.length > 0) {
            text += "Context/filter dimensions with one visible value:\n" +
                context.map(function(item) { return "- " + formatItem(item); }).join("\n") + "\n";
        }
        if (requested.length > 0) {
            text += "Dimensions explicitly requested by the user: " + requested.join(", ") +
                ". Include their exact value in every applicable high, low, comparison, and finding.\n";
        }

        return text;
    }


    private normalisePromptFieldToken(value: string): string {
        return String(value || "")
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    // private getPromptFieldAliases(fieldName: string): string[] {
    //     const base = this.normalisePromptFieldToken(fieldName);
    //     if (!base) { return []; }

    //     const aliases: { [key: string]: boolean } = {};
    //     const add = (value: string): void => {
    //         const clean = this.normalisePromptFieldToken(value);
    //         if (clean) { aliases[clean] = true; }
    //     };

    //     add(base);
    //     add(base.replace(/s$/, ""));
    //     add(base + "s");
    //     add(base.replace(/y$/, "ies"));
    //     add(base.replace(/ies$/, "y"));

    //     // Token-level aliases help with display names such as "Time Period" or "Brand Name".
    //     const tokens = base.split(" ").filter(function(t) { return t.length > 1; });
    //     for (let i = 0; i < tokens.length; i++) {
    //         add(tokens[i]);
    //         add(tokens[i].replace(/s$/, ""));
    //         add(tokens[i] + "s");
    //         add(tokens[i].replace(/y$/, "ies"));
    //         add(tokens[i].replace(/ies$/, "y"));
    //     }

    //     // Generic language aliases. These are not project-specific values; they map common
    //     // business wording in prompts to whatever matching Category Data fields are present.
    //     if (base.indexOf("country") !== -1 || base.indexOf("market") !== -1 || base.indexOf("region") !== -1 || base.indexOf("geo") !== -1) {
    //         add("country"); add("countries"); add("market"); add("markets"); add("region"); add("regions"); add("geography"); add("geographies");
    //     }
    //     if (base.indexOf("brand") !== -1) { add("brand"); add("brands"); }
    //     if (base.indexOf("age") !== -1 || base.indexOf("demo") !== -1 || base.indexOf("gender") !== -1 || base.indexOf("education") !== -1) {
    //         add("demographic"); add("demographics"); add("demo"); add("demos"); add("age"); add("age band"); add("age bands"); add("gender"); add("education");
    //     }

    //     return Object.keys(aliases);
    // }

    private getPromptFieldAliases(fieldName: string): string[] {
        const base = this.normalisePromptFieldToken(fieldName);
        if (!base) { return []; }

        const aliases: { [key: string]: boolean } = {};
        const add = (value: string): void => {
            const clean = this.normalisePromptFieldToken(value);
            if (clean) { aliases[clean] = true; }
        };

        add(base);
        add(base.replace(/s$/, ""));
        add(base + "s");
        add(base.replace(/y$/, "ies"));
        add(base.replace(/ies$/, "y"));

        const tokens = base.split(" ").filter(function(token) { return token.length > 1; });
        for (let i = 0; i < tokens.length; i++) {
            add(tokens[i]);
            add(tokens[i].replace(/s$/, ""));
            add(tokens[i] + "s");
        }

        if (/market|country|region|geography|geo/.test(base)) {
            add("market"); add("markets"); add("country"); add("countries");
            add("region"); add("regions"); add("geography"); add("geographies");
        }
        if (/brand/.test(base)) { add("brand"); add("brands"); }
        if (/category/.test(base)) { add("category"); add("categories"); }
        if (/gender|sex|age|age band|ageband|generation|income|education|race|ethnicity|sec|social class|household|marital|occupation|demographic/.test(base)) {
            add("demographic"); add("demographics"); add("demographic breakdown");
            add("gender"); add("genders"); add("age"); add("age band"); add("age bands");
            add("income"); add("education"); add("race"); add("ethnicity");
        }
        if (/time|period|quarter|month|week|year|date/.test(base)) {
            add("time"); add("time period"); add("period"); add("periods");
            add("quarter"); add("quarters"); add("month"); add("months"); add("year"); add("years");
        }

        return Object.keys(aliases);
    }

    private inferPreferredEntityFieldFromPrompt(
        promptInstruction: string,
        dataCategories: powerbi.DataViewCategoryColumn[]
    ): string {
        const prompt = " " + this.normalisePromptFieldToken(promptInstruction) + " ";
        if (!prompt.trim() || !dataCategories || !dataCategories.length) { return ""; }

        let bestField = "";
        let bestScore = 0;

        for (let i = 0; i < dataCategories.length; i++) {
            const fieldName = this.getColumnDisplayName(dataCategories[i].source, "");
            const fieldKey = this.normalisePromptFieldToken(fieldName);
            if (!fieldKey) { continue; }
            if (/prompt|instruction|text|metric|measure|significance|period|date|time|year|month|week|quarter/.test(fieldKey)) { continue; }

            const aliases = this.getPromptFieldAliases(fieldName);
            let score = 0;
            for (let a = 0; a < aliases.length; a++) {
                const alias = aliases[a];
                if (!alias) { continue; }
                if (prompt.indexOf(" " + alias + " ") !== -1) {
                    score += alias.length > 4 ? 20 : 10;
                }
                if (prompt.indexOf("multiple " + alias) !== -1 || prompt.indexOf("by " + alias) !== -1 || prompt.indexOf("across " + alias) !== -1) {
                    score += 35;
                }
                if (prompt.indexOf("one " + alias) !== -1 || prompt.indexOf("single " + alias) !== -1) {
                    score -= 25;
                }
            }

            // Prefer fields that actually vary in the current filtered dataset.
            const uniqueValues = this.getUniqueNonBlankValues(dataCategories[i].values, 25);
            if (uniqueValues.length > 1) { score += 8; }
            if (uniqueValues.length <= 1) { score -= 8; }

            if (score > bestScore) {
                bestScore = score;
                bestField = fieldName;
            }
        }

        return bestScore > 0 ? bestField : "";
    }

    private promptMentionsField(promptInstruction: string, fieldName: string): boolean {
    const prompt = " " + this.normalisePromptFieldToken(promptInstruction) + " ";
    const aliases = this.getPromptFieldAliases(fieldName);
 
    for (let i = 0; i < aliases.length; i++) {
        const alias = aliases[i];
        if (!alias) { continue; }
 
        if (prompt.indexOf(" " + alias + " ") !== -1) {
            return true;
        }
 
        if (
            prompt.indexOf("multiple " + alias) !== -1 ||
            prompt.indexOf("by " + alias) !== -1 ||
            prompt.indexOf("across " + alias) !== -1 ||
            prompt.indexOf("over " + alias) !== -1
        ) {
            return true;
        }
    }
 
    return false;
}
 
private promptMentionsMetrics(promptInstruction: string): boolean {
    const prompt = " " + this.normalisePromptFieldToken(promptInstruction) + " ";
 
    return (
        prompt.indexOf(" metric ") !== -1 ||
        prompt.indexOf(" metrics ") !== -1 ||
        prompt.indexOf(" kpi ") !== -1 ||
        prompt.indexOf(" kpis ") !== -1 ||
        prompt.indexOf(" measure ") !== -1 ||
        prompt.indexOf(" measures ") !== -1
    );
}
 
// private collectHighlightValuesFromPrompt(
//     promptInstruction: string,
//     dataCategories: powerbi.DataViewCategoryColumn[],
//     metricCategories: powerbi.DataViewCategoryColumn[]
// ): { fields: string[]; values: string[] } {
//     const fields: string[] = [];
//     const values: string[] = [];
//     const seenField: { [key: string]: boolean } = {};
//     const seenValue: { [key: string]: boolean } = {};
 
//     const addFieldAndValues = (category: powerbi.DataViewCategoryColumn): void => {
//         const fieldName = this.getColumnDisplayName(category.source, "");
//         if (!fieldName) { return; }
 
//         const fieldKey = this.normalisePromptFieldToken(fieldName);
//         if (!seenField[fieldKey]) {
//             seenField[fieldKey] = true;
//             fields.push(fieldName);
//         }
 
//         const uniqueValues = this.getUniqueNonBlankValues(category.values, 100);
//         for (let i = 0; i < uniqueValues.length; i++) {
//             const value = String(uniqueValues[i] || "").trim();
//             if (!value || value === "... more selected") { continue; }
 
//             const valueKey = value.toLowerCase();
//             if (!seenValue[valueKey]) {
//                 seenValue[valueKey] = true;
//                 values.push(value);
//             }
//         }
//     };
 
//     const metricRequested = this.promptMentionsMetrics(promptInstruction);
 
//     if (metricRequested) {
//         for (let i = 0; i < metricCategories.length; i++) {
//             addFieldAndValues(metricCategories[i]);
//         }
//     }
 
//     for (let i = 0; i < dataCategories.length; i++) {
//         const fieldName = this.getColumnDisplayName(dataCategories[i].source, "");
//         if (!fieldName) { continue; }
 
//         if (this.promptMentionsField(promptInstruction, fieldName)) {
//             addFieldAndValues(dataCategories[i]);
//         }
//     }
 
//     if (fields.length === 0 && this.preferredEntityFieldFromPrompt) {
//         for (let i = 0; i < dataCategories.length; i++) {
//             const fieldName = this.getColumnDisplayName(dataCategories[i].source, "");
//             if (fieldName === this.preferredEntityFieldFromPrompt) {
//                 addFieldAndValues(dataCategories[i]);
//                 break;
//             }
//         }
//     }
 
//     return {
//         fields: fields,
//         values: values
//     };
// }

private collectHighlightValuesFromPrompt(
    promptInstruction: string,
    dataCategories: powerbi.DataViewCategoryColumn[],
    metricCategories: powerbi.DataViewCategoryColumn[]
): { fields: string[]; values: string[] } {
    const fields: string[] = [];
    const values: string[] = [];
    const seenValue: { [key: string]: boolean } = {};
    const seenField: { [key: string]: boolean } = {};
 
    const addValuesFromCategory = (category: powerbi.DataViewCategoryColumn): void => {
        const fieldName = this.getColumnDisplayName(category.source, "");
        if (!fieldName) {
            return;
        }
 
        const fieldKey = this.normalisePromptFieldToken(fieldName);
        if (!seenField[fieldKey]) {
            seenField[fieldKey] = true;
            fields.push(fieldName);
        }
 
        const uniqueValues = this.getUniqueNonBlankValues(category.values, 100);
 
        for (let i = 0; i < uniqueValues.length; i++) {
            const value = String(uniqueValues[i] || "").trim();
 
            if (!value || value === "... more selected") {
                continue;
            }
 
            const valueKey = value.toLowerCase();
 
            if (!seenValue[valueKey]) {
                seenValue[valueKey] = true;
                values.push(value);
            }
        }
    };
 
    const prompt = " " + this.normalisePromptFieldToken(promptInstruction) + " ";
 
    const wantsMetricLevel =
        prompt.indexOf(" metric ") !== -1 ||
        prompt.indexOf(" metrics ") !== -1 ||
        prompt.indexOf(" kpi ") !== -1 ||
        prompt.indexOf(" kpis ") !== -1 ||
        prompt.indexOf(" measure ") !== -1 ||
        prompt.indexOf(" measures ") !== -1 ||
        prompt.indexOf(" score ") !== -1 ||
        prompt.indexOf(" scores ") !== -1;
 
    const wantsBrandLevel =
        prompt.indexOf(" brand ") !== -1 ||
        prompt.indexOf(" brands ") !== -1;
 
    const wantsCountryLevel =
        prompt.indexOf(" country ") !== -1 ||
        prompt.indexOf(" countries ") !== -1 ||
        prompt.indexOf(" market ") !== -1 ||
        prompt.indexOf(" markets ") !== -1;
 
    const wantsDemographicLevel =
        prompt.indexOf(" demographic ") !== -1 ||
        prompt.indexOf(" demographics ") !== -1 ||
        prompt.indexOf(" ageband ") !== -1 ||
        prompt.indexOf(" agebands ") !== -1 ||
        prompt.indexOf(" age band ") !== -1 ||
        prompt.indexOf(" gender ") !== -1 ||
        prompt.indexOf(" race ") !== -1 ||
        prompt.indexOf(" income ") !== -1 ||
        prompt.indexOf(" education ") !== -1;
 
    /*
     * Priority order:
     * 1. Metrics/KPIs
     * 2. Demographics
     * 3. Countries/Markets
     * 4. Brands
     *
     * This prevents highlighting all context values together.
     */
 
    if (wantsMetricLevel && metricCategories.length > 0) {
        for (let i = 0; i < metricCategories.length; i++) {
            addValuesFromCategory(metricCategories[i]);
        }
 
        return {
            fields: fields,
            values: values
        };
    }
 
    for (let i = 0; i < dataCategories.length; i++) {
        const fieldName = this.getColumnDisplayName(dataCategories[i].source, "");
        const fieldToken = this.normalisePromptFieldToken(fieldName);
 
        if (!fieldName) {
            continue;
        }
 
        if (
            wantsDemographicLevel &&
            (
                fieldToken.indexOf("age") !== -1 ||
                fieldToken.indexOf("gender") !== -1 ||
                fieldToken.indexOf("race") !== -1 ||
                fieldToken.indexOf("income") !== -1 ||
                fieldToken.indexOf("education") !== -1 ||
                fieldToken.indexOf("sec") !== -1
            )
        ) {
            addValuesFromCategory(dataCategories[i]);
        }
    }
 
    if (values.length > 0) {
        return {
            fields: fields,
            values: values
        };
    }
 
    for (let i = 0; i < dataCategories.length; i++) {
        const fieldName = this.getColumnDisplayName(dataCategories[i].source, "");
        const fieldToken = this.normalisePromptFieldToken(fieldName);
 
        if (!fieldName) {
            continue;
        }
 
        if (
            wantsCountryLevel &&
            (
                fieldToken.indexOf("country") !== -1 ||
                fieldToken.indexOf("market") !== -1
            )
        ) {
            addValuesFromCategory(dataCategories[i]);
        }
    }
 
    if (values.length > 0) {
        return {
            fields: fields,
            values: values
        };
    }
 
    for (let i = 0; i < dataCategories.length; i++) {
        const fieldName = this.getColumnDisplayName(dataCategories[i].source, "");
        const fieldToken = this.normalisePromptFieldToken(fieldName);
 
        if (!fieldName) {
            continue;
        }
 
        if (
            wantsBrandLevel &&
            fieldToken.indexOf("brand") !== -1
        ) {
            addValuesFromCategory(dataCategories[i]);
        }
    }
 
    return {
        fields: fields,
        values: values
    };
}
 
 

    private getCategoryIdentityKey(category: powerbi.DataViewCategoryColumn, fallbackIndex: number = 0): string {
        const source: any = category && category.source ? category.source : {};
        const rawIdentity = String(source.queryName || source.displayName || ("category_" + fallbackIndex));
        return this.normalisePromptFieldToken(rawIdentity) || ("category_" + fallbackIndex);
    }

    private getCategorySemanticToken(category: powerbi.DataViewCategoryColumn): string {
        const source: any = category && category.source ? category.source : {};
        return this.normalisePromptFieldToken(
            [String(source.displayName || ""), String(source.queryName || "")].join(" ")
        );
    }

    private isLikelyPromptCategory(category: powerbi.DataViewCategoryColumn): boolean {
        if (this.sourceHasRole(category.source, "promptText")) { return true; }
        const token = this.getCategorySemanticToken(category);
        return /(^| )(prompt|instruction|user prompt|prompt text)( |$)/.test(token);
    }

    private isLikelyMetricCategory(category: powerbi.DataViewCategoryColumn): boolean {
        if (this.sourceHasRole(category.source, "metricCategory")) { return true; }

        const source: any = category && category.source ? category.source : {};
        const displayToken = this.normalisePromptFieldToken(String(source.displayName || ""));
        const semanticToken = this.getCategorySemanticToken(category);

        // Group/family fields are descriptive dimensions, not the row-level Metric.
        if (/(^| )(metric group|metric family|metric category|metric type|kpi group|measure group)( |$)/.test(semanticToken)) {
            return false;
        }

        return displayToken === "metric" ||
            displayToken === "metric name" ||
            displayToken === "kpi" ||
            displayToken === "kpi name" ||
            /(^| )(metric name|kpi name|measure name|funnel stage|journey stage)( |$)/.test(semanticToken);
    }

    private splitCategoriesByRoles(categories: powerbi.DataViewCategoryColumn[]): {
        promptCategories: powerbi.DataViewCategoryColumn[];
        dataCategories: powerbi.DataViewCategoryColumn[];
        metricCategories: powerbi.DataViewCategoryColumn[];
        highlightCategories: powerbi.DataViewCategoryColumn[];
    } {
        const promptCategories: powerbi.DataViewCategoryColumn[] = [];
        const dataCategories: powerbi.DataViewCategoryColumn[] = [];
        const metricCategories: powerbi.DataViewCategoryColumn[] = [];
        const highlightCategories: powerbi.DataViewCategoryColumn[] = [];

        const seenPrompt: { [key: string]: boolean } = {};
        const seenData: { [key: string]: boolean } = {};
        const seenMetric: { [key: string]: boolean } = {};
        const seenHighlight: { [key: string]: boolean } = {};

        const addUnique = (
            target: powerbi.DataViewCategoryColumn[],
            seen: { [key: string]: boolean },
            category: powerbi.DataViewCategoryColumn,
            index: number
        ): void => {
            const key = this.getCategoryIdentityKey(category, index);
            if (!seen[key]) {
                seen[key] = true;
                target.push(category);
            }
        };

        for (let i = 0; i < categories.length; i++) {
            const category = categories[i];
            const displayName = String(category.source && category.source.displayName ? category.source.displayName : "");
            const isHighlightColumn = this.sourceHasRole(category.source, "highlightEntity") || /highlight|focus/i.test(displayName);

            // Highlight is orthogonal metadata. A field can be both Highlight Entity
            // and a real Metric/Category dimension. Never remove it from the dataset.
            if (isHighlightColumn) {
                addUnique(highlightCategories, seenHighlight, category, i);
            }

            if (this.isLikelyPromptCategory(category)) {
                addUnique(promptCategories, seenPrompt, category, i);
            } else if (this.isLikelyMetricCategory(category)) {
                addUnique(metricCategories, seenMetric, category, i);
            } else {
                addUnique(dataCategories, seenData, category, i);
            }
        }

        return {
            promptCategories: promptCategories,
            dataCategories: dataCategories,
            metricCategories: metricCategories,
            highlightCategories: highlightCategories
        };
    }

    // private extractPromptInstructionFromPromptCategories(promptCategories: powerbi.DataViewCategoryColumn[]): string {
    //     const lines: string[] = [];
    //     const seen: { [key: string]: boolean } = {};

    //     for (let i = 0; i < promptCategories.length; i++) {
    //         const values = this.getUniqueNonBlankValues(promptCategories[i].values, 50);
    //         for (let j = 0; j < values.length; j++) {
    //             const value = String(values[j] || "").trim();
    //             if (!value || value === "... more selected") { continue; }
    //             const key = value.toLowerCase().replace(/\s+/g, " ");
    //             if (!seen[key]) {
    //                 seen[key] = true;
    //                 lines.push(value);
    //             }
    //         }
    //     }

    //     return lines.join("\n").trim();
    // }

    // private extractPromptInstructionFromPromptCategories(promptCategories: powerbi.DataViewCategoryColumn[]): string {
    // for (let i = 0; i < promptCategories.length; i++) {
    //     const values = this.getUniqueNonBlankValues(promptCategories[i].values, 50);
 
    //     for (let j = 0; j < values.length; j++) {
    //         const value = String(values[j] || "").trim();
 
    //         if (value && value !== "... more selected") {
    //             return value;
    //         }
    //     }
    // }
 
    // return "";
    // }
 
    // private extractPromptInstructionFromPromptCategories(promptCategories: powerbi.DataViewCategoryColumn[]): string {
    // for (let i = 0; i < promptCategories.length; i++) {
    //     const values = this.getUniqueNonBlankValues(promptCategories[i].values, 50);
 
    //     for (let j = 0; j < values.length; j++) {
    //         const value = String(values[j] || "").trim();
 
    //         if (value && value !== "... more selected") {
    //             return value;
    //         }
    //     }
    // }
 
    // return "";
    // }

    private extractPromptInstructionFromPromptCategories(promptCategories: powerbi.DataViewCategoryColumn[]): string {
    for (let i = 0; i < promptCategories.length; i++) {
        const values = this.getUniqueNonBlankValues(promptCategories[i].values, 50);
 
        for (let j = 0; j < values.length; j++) {
            const value = String(values[j] || "").trim();
 
            if (value && value !== "... more selected") {
                return value;
            }
        }
    }
 
    return "";
    }
 

    private classifyMeasureColumns(values: powerbi.DataViewValueColumns | undefined): {
        numericMeasures: powerbi.DataViewValueColumn[];
        textMeasures: powerbi.DataViewValueColumn[];
    } {
        const numericMeasures: powerbi.DataViewValueColumn[] = [];
        const textMeasures: powerbi.DataViewValueColumn[] = [];

        if (!values) {
            return { numericMeasures: numericMeasures, textMeasures: textMeasures };
        }

        Array.from(values).forEach((valueColumn: powerbi.DataViewValueColumn) => {
            const vals = valueColumn.values || [];
            let nonBlank = 0;
            let numeric = 0;
            let textLike = 0;
            const sample = Math.min(vals.length, 100);

            for (let i = 0; i < sample; i++) {
                const raw = vals[i];
                const text = String(raw == null ? "" : raw).trim();
                if (!text || text.toLowerCase() === "n/a") { continue; }
                nonBlank++;
                if (this.parseNumericValue(text) !== null) {
                    numeric++;
                } else {
                    textLike++;
                }
            }

            // Only Measure Data fields with mostly numeric values can be scores.
            // Text measure fields, for example Significance_Text, remain text evidence.
            if (nonBlank > 0 && numeric >= Math.max(1, Math.ceil(nonBlank * 0.6))) {
                numericMeasures.push(valueColumn);
            } else if (textLike > 0 || nonBlank > 0) {
                textMeasures.push(valueColumn);
            }
        });

        return { numericMeasures: numericMeasures, textMeasures: textMeasures };
    }

    private getColumnDisplayName(source: powerbi.DataViewMetadataColumn | undefined, fallback: string): string {
        return source && source.displayName ? source.displayName : fallback;
    }

    private getMetricValue(metricCategory: powerbi.DataViewCategoryColumn | null, rowIndex: number, fallback: string): string {
        if (!metricCategory || !metricCategory.values || metricCategory.values.length <= rowIndex) {
            return fallback;
        }
        const value = metricCategory.values[rowIndex];
        const text = String(value == null ? "" : value).trim();
        return text || fallback;
    }

    private rememberAllowedScoreHeader(header: string): void {
        this.allowedScoreHeaderKeys[this.normaliseEvidenceKey(header)] = true;
    }

    // -------------------------------------------------------------------------
    // formatFullPivotedData - returns ALL rows with no row cap
    // -------------------------------------------------------------------------

    private formatFullPivotedData(): string {
        if (!this.dataView || !this.dataView.categorical) { return "No data available"; }

        const categorical = this.dataView.categorical;
        const categories = categorical.categories || [];
        const values = categorical.values;
        if (!categories.length || !values || values.length === 0) {
            return "No data available";
        }

        const split = this.splitCategoriesByRoles(categories);
        const promptInstructionForData = this.extractPromptInstructionFromPromptCategories(
            split.promptCategories
        );
        const dimensionCategories = this.getAnalyticalCategoryColumns(categories);
        this.preferredEntityFieldFromPrompt = this.inferPreferredEntityFieldFromPrompt(
            promptInstructionForData || this.DEFAULT_ANALYST_PROMPT,
            dimensionCategories
        );

        const metricCategory = split.metricCategories.length > 0
            ? split.metricCategories[0]
            : null;
        const csvResult = buildFullPivotedCsv(
            this.dataView,
            dimensionCategories,
            metricCategory,
            this.MAX_CONTEXT_VALUES_PER_FIELD
        );

        this.allowedScoreHeaderKeys = {};
        for (let i = 0; i < csvResult.scoreHeaders.length; i++) {
            this.rememberAllowedScoreHeader(csvResult.scoreHeaders[i]);
        }

        return csvResult.csvText;
    }

    // -------------------------------------------------------------------------
    // Progress message helper
    // -------------------------------------------------------------------------

    private updateProgressMessage(message: string): void {
        this.currentProgressMessage = String(message || "Preparing dataset...").trim();
        if (this.isLoadingInsights) {
            this.renderCurrentResponse();
        }
    }

    // -------------------------------------------------------------------------
    // All methods below are unchanged from the original visual.ts
    // -------------------------------------------------------------------------

    private updateShowButtonState(isLoading: boolean): void {
        const showButton = this.target.querySelector(".button_ShowInsights") as HTMLButtonElement | null;
        if (!showButton) { return; }
        showButton.textContent = isLoading ? "Loading..." : "Show";
        showButton.classList.toggle("is-loading", isLoading);
        showButton.disabled = isLoading;
        showButton.setAttribute("aria-busy", isLoading ? "true" : "false");
    }

    private async getGptResponse(prompt: string): Promise<string> {
        return this.callAzureOpenAIWithRetry(prompt, 4000, 1, "text_final");
    }

    private updateGptResponse(content: string): void {
        const responseDisplay = this.target.querySelector("#gptResponse") as HTMLElement | null;
        if (!responseDisplay) { return; }

        responseDisplay.replaceChildren();
        responseDisplay.setAttribute("aria-busy", "false");

        const cleaned = String(content || "").trim();
        if (!cleaned) {
            this.renderMessageCard(
                "AI returned an empty response. Click Show to try again or reduce the selected fields.",
                "error"
            );
            return;
        }

        const sections = this.parseInsightIntoSections(cleaned);
        const renderSections = sections.length > 0
            ? sections
            : [{ heading: "Summary", bullets: [cleaned] }];

        for (let sectionIndex = 0; sectionIndex < renderSections.length; sectionIndex++) {
            const section = renderSections[sectionIndex];
            const bullets = section.bullets.filter(function(value) {
                return String(value || "").trim().length > 0;
            });
            if (bullets.length === 0) { continue; }

            const rawHeading = String(section.heading || "Summary").trim();
            const safeHeading = rawHeading.toLowerCase() === "insights" ? "Summary" : rawHeading;

            const heading = document.createElement("div");
            heading.className = "ai-section-title";
            const headingStrong = document.createElement("strong");
            headingStrong.textContent = safeHeading || "Summary";
            heading.appendChild(headingStrong);
            responseDisplay.appendChild(heading);

            const list = document.createElement("div");
            list.className = "ai-bullet-list";
            list.setAttribute("role", "list");

            for (let bulletIndex = 0; bulletIndex < bullets.length; bulletIndex++) {
                const bulletText = String(bullets[bulletIndex] || "").trim();
                if (!bulletText) { continue; }

                const item = document.createElement("div");
                item.className = "ai-bullet-item";
                item.setAttribute("role", "listitem");

                const marker = document.createElement("span");
                marker.className = "ai-bullet-marker";
                marker.setAttribute("aria-hidden", "true");
                marker.textContent = this.getPromptListMarker(section.heading, bulletIndex);

                const body = document.createElement("span");
                body.className = "ai-bullet-text";
                this.appendFormattedText(body, bulletText);

                item.appendChild(marker);
                item.appendChild(body);
                list.appendChild(item);
            }

            responseDisplay.appendChild(list);
        }
    }

    private appendFormattedText(target: HTMLElement, rawText: string): void {
        const text = String(rawText || "");

        // The visible renderer deliberately creates DOM nodes rather than assigning
        // innerHTML. Include every value supplied through the Highlight Entity field
        // well so Market, Category and Brand values are highlighted wherever they
        // occur in Summary or Key findings.
        
        // const highlightValues = (this.highlightedCategoryValuesFromPrompt || [])
        const highlightValues = ((this.highlightedCategoryFieldsFromPrompt || []) as string[])
            .concat(this.highlightedCategoryValuesFromPrompt || [])
            .map(function(value) { return String(value || "").trim(); })
            .filter(function(value, index, values) {
                if (!value) { return false; }
                const lower = value.toLowerCase();
                for (let i = 0; i < index; i++) {
                    if (values[i].toLowerCase() === lower) { return false; }
                }
                return true;
            })
            .sort(function(left, right) { return right.length - left.length; });

        const dynamicAlternation = highlightValues.length > 0
            ? "|" + highlightValues.map((value) => this.escapeRegExp(value)).join("|")
            : "";
        const tokenPattern = new RegExp(
            "(\\*\\*[^*]+\\*\\*|[+-]?\\d+(?:\\.\\d+)?%" + dynamicAlternation + ")",
            "gi"
        );

        let lastIndex = 0;
        let match: RegExpExecArray | null;

        while ((match = tokenPattern.exec(text)) !== null) {
            if (match.index > lastIndex) {
                target.appendChild(document.createTextNode(text.slice(lastIndex, match.index)));
            }

            const token = match[0];
            if (token.indexOf("**") === 0 && token.lastIndexOf("**") === token.length - 2) {
                const strong = document.createElement("strong");
                strong.textContent = token.slice(2, -2);
                target.appendChild(strong);
            } else if (/^[+-]?\d+(?:\.\d+)?%$/.test(token)) {
                const percentage = document.createElement("span");
                percentage.className = "ai-percent";
                percentage.textContent = token;
                target.appendChild(percentage);
            } else {
                const entity = document.createElement("strong");
                entity.className = "ai-dynamic-highlight";
                entity.textContent = token;
                target.appendChild(entity);
            }

            lastIndex = match.index + token.length;
        }

        if (lastIndex < text.length) {
            target.appendChild(document.createTextNode(text.slice(lastIndex)));
        }
    }

    private isValidAiResponse(response: string | undefined): boolean {
        const text      = String(response || "").trim();
        if (!text) { return false; }
        const lowerText = text.toLowerCase();

        return (
            lowerText.indexOf("ai service returned an empty message") === -1 &&
            lowerText.indexOf("ai returned an empty response")        === -1 &&
            lowerText.indexOf("unable to generate insights")          === -1 &&
            lowerText.indexOf("please try again")                     === -1 &&
            lowerText.indexOf("authentication error")                 === -1 &&
            lowerText.indexOf("too many requests")                    === -1 &&
            lowerText.indexOf("server error")                         === -1
        );
    }

    private formatInsightTextAsHtml(content: string): string {
        const cleaned = String(content || "").trim();

        if (!cleaned) {
            return "<div class=\"ai-message-card\">AI returned an empty response. Try clicking Show again or reduce the selected fields.</div>";
        }

        const normalized = cleaned
            .replace(/\r\n/g, "\n")
            .replace(/\r/g, "\n")
            .replace(/\n{3,}/g, "\n\n");

        const lines              = normalized.split("\n");
        const htmlParts: string[] = [];
        let bulletBuffer: string[]  = [];
        let numberedBuffer: string[] = [];
        const self = this;

        const staticKnownHeadings = [
            "summary", "key findings", "key finding", "key insights", "important findings", "highlights",
            "top performers", "watchouts", "recommendation", "recommendations",
            "overall patterns", "top kpi signals", "key risks", "key opportunities", "critical anomalies",
            "brand performance analysis", "highest and lowest analysis", "summary of significant changes", "executive summary", "key observations", "overview"
        ];

        // Headings the user explicitly asked for in their own prompt (see
        // extractHeadingsFromPrompt / buildFinalOutputFormatInstruction) always count
        // as headings too, even if they aren't in the static list above.
        const promptHeadings = (this.lastKnownHeadingsFromPrompt || []).map(function(h) {
            return String(h || "").trim().toLowerCase();
        }).filter(function(h) { return h.length > 0; });

        const knownHeadings = staticKnownHeadings.concat(promptHeadings);

        const flushBullets = function(): void {
            if (bulletBuffer.length > 0) {
                htmlParts.push("<ul class=\"ai-bullet-list\">" + bulletBuffer.join("") + "</ul>");
                bulletBuffer = [];
            }
        };

        const flushNumbered = function(): void {
            if (numberedBuffer.length > 0) {
                htmlParts.push("<ol class=\"ai-numbered-list\">" + numberedBuffer.join("") + "</ol>");
                numberedBuffer = [];
            }
        };

        const stripMarkdownHeading = function(value: string): string {
            return value
                .replace(/^#{1,6}\s+/, "")
                .replace(/^\*\*(.*?)\*\*:?$/, "$1:")
                .replace(/^__(.*?)__:?$/, "$1:")
                .trim();
        };

        const isHeadingLine = function(value: string): boolean {
            const stripped = stripMarkdownHeading(value);
            const noColon  = stripped.replace(/:$/, "").trim().toLowerCase();

            if (knownHeadings.indexOf(noColon) !== -1) { return true; }
            if (/^[A-Za-z][A-Za-z0-9\s\-/&()]{2,}:$/.test(stripped) && stripped.length <= 90) { return true; }
            if (/^\*\*[^*]{2,80}\*\*:?$/.test(value)) { return true; }
            return false;
        };

        for (let i = 0; i < lines.length; i++) {
            const rawLine = lines[i];
            const line    = rawLine.trim();

            if (!line) { flushBullets(); flushNumbered(); continue; }

            if (isHeadingLine(line)) {
                flushBullets();
                flushNumbered();
                const headingText = stripMarkdownHeading(line).replace(/:$/, "");
                const headingHtml = self.applyInlineFormatting(self.escapeHtml(headingText));
                htmlParts.push("<div class=\"ai-section-title\"><strong>" + headingHtml + "</strong></div>");
                continue;
            }

            if (/^[-*-]\s+/.test(line)) {
                flushNumbered();
                const bulletText = self.applyInlineFormatting(self.escapeHtml(line.replace(/^[-*-]\s+/, "")));
                bulletBuffer.push("<li>" + bulletText + "</li>");
                continue;
            }

            if (/^\d+[.)]\s+/.test(line)) {
                flushBullets();
                const numberText = self.applyInlineFormatting(self.escapeHtml(line.replace(/^\d+[.)]\s+/, "")));
                numberedBuffer.push("<li>" + numberText + "</li>");
                continue;
            }

            flushBullets();
            flushNumbered();
            htmlParts.push("<p class=\"ai-paragraph\">" + self.applyInlineFormatting(self.escapeHtml(line)) + "</p>");
        }

        flushBullets();
        flushNumbered();

        return htmlParts.join("");
    }

    // -------------------------------------------------------------------------
    // PPT export - shares the same heading detection as formatInsightTextAsHtml
    // (static list + whatever headings the user's own prompt asked for) so the
    // slide breakdown always matches what's shown on screen.
    // -------------------------------------------------------------------------

    private parseInsightIntoSections(content: string): Array<{ heading: string; bullets: string[] }> {
        const cleaned = String(content || "").trim();
        if (!cleaned) { return []; }

        const normalized = cleaned
            .replace(/\r\n/g, "\n")
            .replace(/\r/g, "\n")
            .replace(/\u2022/g, "-");
        const lines = normalized.split("\n");

        const staticKnownHeadings = [
            "summary", "key findings", "key finding", "key insights", "important findings", "highlights",
            "insights", "top performers", "watchouts", "recommendation", "recommendations",
            "overall patterns", "top kpi signals", "key risks", "key opportunities", "critical anomalies",
            "brand performance analysis", "highest and lowest analysis", "summary of significant changes",
            "executive summary", "key observations", "overview"
        ];
        const promptHeadings = (this.lastKnownHeadingsFromPrompt || [])
            .map(function(h) { return String(h || "").trim().toLowerCase(); })
            .filter(function(h) { return h.length > 0; });
        const knownHeadings = staticKnownHeadings.concat(promptHeadings)
            .filter(function(value, index, array) { return array.indexOf(value) === index; })
            .sort(function(a, b) { return b.length - a.length; });

        const stripMarkdownHeading = function(value: string): string {
            return value
                .replace(/^#{1,6}\s+/, "")
                .replace(/^\*\*(.*?)\*\*:?$/, "$1:")
                .replace(/^__(.*?)__:?$/, "$1:")
                .trim();
        };

        const normaliseHeading = function(value: string): string {
            const clean = stripMarkdownHeading(value).replace(/:$/, "").trim();
            return clean.toLowerCase() === "insights" ? "Summary" : clean;
        };

        const exactHeading = function(value: string): string {
            const stripped = stripMarkdownHeading(value);
            const noColon = stripped.replace(/:$/, "").trim().toLowerCase();
            if (knownHeadings.indexOf(noColon) !== -1) { return normaliseHeading(stripped); }
            if (/^[A-Za-z][A-Za-z0-9\s\-/&()]{2,}:$/.test(stripped) && stripped.length <= 90) {
                return normaliseHeading(stripped);
            }
            if (/^\*\*[^*]{2,80}\*\*:?$/.test(value)) { return normaliseHeading(value); }
            return "";
        };

        const headingWithContent = function(value: string): { heading: string; remainder: string } | null {
            const stripped = stripMarkdownHeading(value);
            const lower = stripped.toLowerCase();
            for (let i = 0; i < knownHeadings.length; i++) {
                const candidate = knownHeadings[i];
                const prefix = candidate + ":";
                if (lower.indexOf(prefix) === 0) {
                    return {
                        heading: candidate === "insights"
                            ? "Summary"
                            : candidate.replace(/\b\w/g, function(letter) { return letter.toUpperCase(); }),
                        remainder: stripped.slice(prefix.length).trim()
                    };
                }
            }
            return null;
        };

        const sections: Array<{ heading: string; bullets: string[] }> = [];
        let current: { heading: string; bullets: string[] } | null = null;
        let previousWasBlank = true;
        let previousWasHeading = false;

        const ensureSection = function(): { heading: string; bullets: string[] } {
            if (!current) {
                current = { heading: "Summary", bullets: [] };
                sections.push(current);
            }
            return current;
        };

        const pushBullet = function(value: string, appendToPrevious: boolean): void {
            const bullet = String(value || "")
                .replace(/^[-*---]\s+/, "")
                .replace(/^\d+[.)]\s+/, "")
                .trim();
            if (!bullet) { return; }

            const section = ensureSection();
            if (appendToPrevious && section.bullets.length > 0) {
                const lastIndex = section.bullets.length - 1;
                section.bullets[lastIndex] = (section.bullets[lastIndex] + " " + bullet)
                    .replace(/\s+/g, " ")
                    .trim();
            } else {
                section.bullets.push(bullet);
            }
        };

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) {
                previousWasBlank = true;
                previousWasHeading = false;
                continue;
            }

            const exact = exactHeading(line);
            if (exact) {
                current = { heading: exact, bullets: [] };
                sections.push(current);
                previousWasBlank = false;
                previousWasHeading = true;
                continue;
            }

            const prefixed = headingWithContent(line);
            if (prefixed) {
                current = { heading: prefixed.heading, bullets: [] };
                sections.push(current);
                if (prefixed.remainder) { pushBullet(prefixed.remainder, false); }
                previousWasBlank = false;
                previousWasHeading = true;
                continue;
            }

            const isExplicitBullet = /^[-*---]\s+/.test(line) || /^\d+[.)]\s+/.test(line);
            const appendToPrevious = !isExplicitBullet && !previousWasBlank && !previousWasHeading;
            pushBullet(line, appendToPrevious);

            previousWasBlank = false;
            previousWasHeading = false;
        }

        return sections.filter(function(section) {
            return section.bullets.some(function(bullet) {
                return String(bullet || "").trim().length > 0;
            });
        });
    }

    private buildPptSubtitle(): string {
        if (this.highlightedCategoryValuesFromPrompt && this.highlightedCategoryValuesFromPrompt.length > 0) {
            return this.takeUnique(this.highlightedCategoryValuesFromPrompt, 6).join("  -  ");
        }
        return new Date().toLocaleString();
    }

    private showPptStatus(
        message: string,
        isError: boolean = false,
        autoHideMs: number = 0
    ): void {
        if (this.pptStatusTimeoutId !== undefined) {
            window.clearTimeout(this.pptStatusTimeoutId);
            this.pptStatusTimeoutId = undefined;
        }

        let status = this.target.querySelector(".ppt-download-status") as HTMLElement | null;
        if (!status) {
            status = document.createElement("div");
            status.className = "ppt-download-status";
            status.setAttribute("role", isError ? "alert" : "status");
            status.setAttribute("aria-live", "polite");
            const panel = this.target.querySelector(".ai-panel-shell");
            if (panel) { panel.appendChild(status); }
        }

        status.classList.toggle("is-error", isError);
        status.setAttribute("role", isError ? "alert" : "status");
        status.replaceChildren();

        const icon = document.createElement("span");
        icon.className = "ppt-status-icon";
        icon.setAttribute("aria-hidden", "true");
        icon.textContent = isError ? "!" : "?";

        const text = document.createElement("span");
        text.className = "ppt-status-text";
        text.textContent = String(message || "").trim();

        status.appendChild(icon);
        status.appendChild(text);

        if (autoHideMs > 0) {
            this.pptStatusTimeoutId = window.setTimeout(() => {
                const currentStatus = this.target.querySelector(".ppt-download-status") as HTMLElement | null;
                if (currentStatus) { currentStatus.remove(); }
                this.pptStatusTimeoutId = undefined;
            }, autoHideMs);
        }
    }

    private showPptBrowserFallbackStatus(fileName: string, blob: Blob): void {
        if (this.pptStatusTimeoutId !== undefined) {
            window.clearTimeout(this.pptStatusTimeoutId);
            this.pptStatusTimeoutId = undefined;
        }

        let status = this.target.querySelector(".ppt-download-status") as HTMLElement | null;
        if (!status) {
            status = document.createElement("div");
            status.className = "ppt-download-status";
            status.setAttribute("role", "status");
            status.setAttribute("aria-live", "polite");
            const panel = this.target.querySelector(".ai-panel-shell");
            if (panel) { panel.appendChild(status); }
        }

        status.classList.remove("is-error");
        status.replaceChildren();

        const icon = document.createElement("span");
        icon.className = "ppt-status-icon";
        icon.setAttribute("aria-hidden", "true");
        icon.textContent = "?";

        const text = document.createElement("span");
        text.className = "ppt-status-text";
        text.appendChild(document.createTextNode("PowerPoint is ready. If it did not open automatically,"));

        const objectUrl = URL.createObjectURL(blob);
        const saveLink = document.createElement("a");
        saveLink.className = "ppt-save-link";
        saveLink.href = objectUrl;
        saveLink.download = fileName;
        saveLink.target = "_self";
        saveLink.rel = "noopener";
        saveLink.textContent = "Save PPT";
        saveLink.setAttribute("aria-label", "Save " + fileName);
        text.appendChild(saveLink);

        status.appendChild(icon);
        status.appendChild(text);

        let cleanedUp = false;
        const cleanup = (): void => {
            if (cleanedUp) { return; }
            cleanedUp = true;
            URL.revokeObjectURL(objectUrl);
            const currentStatus = this.target.querySelector(".ppt-download-status") as HTMLElement | null;
            if (currentStatus === status) { currentStatus.remove(); }
            this.pptStatusTimeoutId = undefined;
        };

        saveLink.addEventListener("click", function() {
            window.setTimeout(cleanup, 1200);
        }, { once: true });

        this.pptStatusTimeoutId = window.setTimeout(cleanup, 20000);
    }

    private getChartSeriesFromDataView(): Array<{ title: string; labels: string[]; values: number[] }> {
        const result: Array<{ title: string; labels: string[]; values: number[] }> = [];
        const categorical = this.dataView && this.dataView.categorical;
        if (!categorical || !categorical.values || categorical.values.length === 0) { return result; }

        const categories = categorical.categories || [];
        const valueColumns: powerbi.DataViewValueColumn[] = Array.from(categorical.values) as powerbi.DataViewValueColumn[];
        const split = this.splitCategoriesByRoles(categories);
        const analyticalCategories = this.getAnalyticalCategoryColumns(categories);

        let metricCategory: powerbi.DataViewCategoryColumn | null = split.metricCategories.length > 0
            ? split.metricCategories[0]
            : null;

        if (!metricCategory) {
            for (let i = 0; i < categories.length; i++) {
                const name = this.getColumnDisplayName(categories[i].source, "").toLowerCase();
                if (/\b(metric|kpi|funnel stage|journey stage)\b/.test(name)) {
                    metricCategory = categories[i];
                    break;
                }
            }
        }

        const categoryInfo = analyticalCategories.map((category, index) => {
            const name = this.getColumnDisplayName(category.source, "Dimension_" + (index + 1));
            const values = this.getUniqueNonBlankValues(category.values, 1000);
            const nameToken = this.normalisePromptFieldToken(name);
            
            const highlighted = this.highlightedCategoryFieldsFromPrompt.some((field) => {
                return this.normalisePromptFieldToken(field) === nameToken;
            });
            let priority = values.length > 1 ? 0 : 20;
            if (highlighted) { priority -= 10; }
            if (this.isDemographicDimensionName(name)) { priority -= 5; }
            if (nameToken.indexOf("brand") !== -1) { priority += 1; }
            if (nameToken.indexOf("market") !== -1 || nameToken.indexOf("country") !== -1) { priority += 2; }
            if (nameToken.indexOf("category") !== -1) { priority += 3; }
            return { category: category, name: name, values: values, priority: priority, order: index };
        }).sort(function(left, right) {
            return left.priority - right.priority || left.order - right.order;
        });

        interface ChartBucket {
            sum: number;
            count: number;
            label: string;
        }

        const makePoint = (rowIndex: number): { key: string; label: string } => {
            const keyParts: string[] = [];
            const labelParts: string[] = [];

            for (let i = 0; i < analyticalCategories.length; i++) {
                const category = analyticalCategories[i];
                const name = this.getColumnDisplayName(category.source, "Dimension_" + (i + 1));
                const values = category.values || [];
                const value = String(values[rowIndex] == null ? "" : values[rowIndex]).trim();
                keyParts.push(name + "=" + value);
            }

            for (let i = 0; i < categoryInfo.length && labelParts.length < 6; i++) {
                const item = categoryInfo[i];
                const values = item.category.values || [];
                const value = String(values[rowIndex] == null ? "" : values[rowIndex]).trim();
                if (!value) { continue; }
                const showFieldName = item.values.length > 1 || this.isDemographicDimensionName(item.name);
                const part = showFieldName ? item.name + ": " + value : value;
                if (labelParts.indexOf(part) === -1) { labelParts.push(part); }
            }

            const fullLabel = labelParts.length > 0 ? labelParts.join(" | ") : "Row " + (rowIndex + 1);
            return {
                key: keyParts.join("||") || ("row_" + rowIndex),
                label: fullLabel.length > 96 ? fullLabel.substring(0, 93) + "..." : fullLabel
            };
        };

        const toPercent = function(value: number): number {
            const percent = Math.abs(value) <= 1 ? value * 100 : value;
            return Number(percent.toFixed(1));
        };

        const addChart = (title: string, grouped: { [key: string]: ChartBucket }): void => {
            if (result.length >= 8) { return; }
            const points = Object.keys(grouped)
                .map(function(key) {
                    const bucket = grouped[key];
                    return {
                        label: bucket.label,
                        value: bucket.sum / Math.max(1, bucket.count)
                    };
                })
                .sort(function(left, right) { return right.value - left.value; })
                .slice(0, 12);

            if (points.length === 0) { return; }
            result.push({
                title: title,
                labels: points.map(function(point) { return point.label; }),
                values: points.map(function(point) { return toPercent(point.value); })
            });
        };

        for (let valueIndex = 0; valueIndex < valueColumns.length && result.length < 8; valueIndex++) {
            const valueColumn = valueColumns[valueIndex];
            if (!valueColumn) { continue; }
            const measureName = this.getColumnDisplayName(valueColumn.source, "Value");
            const valueArray = valueColumn.values || [];
            const rowCount = valueArray.length;

            if (metricCategory) {
                const metricGroups: { [metric: string]: { [key: string]: ChartBucket } } = {};
                const metricOrder: string[] = [];

                for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    const numericValue = Number(valueArray[rowIndex]);
                    if (!isFinite(numericValue)) { continue; }
                    const metricValues = metricCategory.values || [];
                    const metric = String(metricValues[rowIndex] == null ? "Metric" : metricValues[rowIndex]).trim() || "Metric";
                    const point = makePoint(rowIndex);

                    if (!metricGroups[metric]) {
                        metricGroups[metric] = {};
                        metricOrder.push(metric);
                    }
                    if (!metricGroups[metric][point.key]) {
                        metricGroups[metric][point.key] = { sum: 0, count: 0, label: point.label };
                    }
                    metricGroups[metric][point.key].sum += numericValue;
                    metricGroups[metric][point.key].count += 1;
                }

                for (let metricIndex = 0; metricIndex < metricOrder.length && result.length < 8; metricIndex++) {
                    const metric = metricOrder[metricIndex];
                    const title = /^(value|score|percentage|percent|measure)$/i.test(measureName)
                        ? metric
                        : metric + " - " + measureName;
                    addChart(title, metricGroups[metric]);
                }
            } else {
                const grouped: { [key: string]: ChartBucket } = {};
                for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    const numericValue = Number(valueArray[rowIndex]);
                    if (!isFinite(numericValue)) { continue; }
                    const point = makePoint(rowIndex);
                    if (!grouped[point.key]) {
                        grouped[point.key] = { sum: 0, count: 0, label: point.label };
                    }
                    grouped[point.key].sum += numericValue;
                    grouped[point.key].count += 1;
                }
                addChart(measureName, grouped);
            }
        }

        return result;
    }

    private async createPptBase64(pres: any): Promise<string> {
        try {
            const value = await pres.write({ outputType: "base64" });
            return String(value || "");
        } catch (firstError) {
            const value = await pres.write("base64");
            return String(value || "");
        }
    }

    private base64ToPptBlob(base64Content: string): Blob {
        const raw = String(base64Content || "").trim();
        if (!raw) { throw new Error("The generated PowerPoint content is empty."); }

        const commaIndex = raw.indexOf(",");
        const payload = commaIndex >= 0 ? raw.substring(commaIndex + 1) : raw;
        const binary = window.atob(payload.replace(/\s/g, ""));
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
        }

        return new Blob([bytes], {
            type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        });
    }

    private triggerBrowserBlobDownload(blob: Blob, fileName: string): void {
        const ownerDocument = this.target.ownerDocument || document;
        const body = ownerDocument.body;
        if (!body) { throw new Error("The browser download surface is unavailable."); }

        const objectUrl = URL.createObjectURL(blob);
        const anchor = ownerDocument.createElement("a");
        anchor.href = objectUrl;
        anchor.download = fileName;
        anchor.target = "_self";
        anchor.rel = "noopener";
        anchor.style.display = "none";
        body.appendChild(anchor);

        try {
            anchor.click();
        } finally {
            window.setTimeout(function() {
                if (anchor.parentNode) { anchor.parentNode.removeChild(anchor); }
                URL.revokeObjectURL(objectUrl);
            }, 1500);
        }
    }

    private async uploadPptBlobToAzureFunction(
        blob: Blob,
        fileName: string,
        functionUrl: string
    ): Promise<PptServerUploadResponse> {
        const endpoint = String(functionUrl || "").trim();
        if (!endpoint) {
            throw new Error("PPT export Function URL is not configured.");
        }

        const response = await fetch(endpoint, {
            method: "POST",
            headers: {
                "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "X-File-Name": encodeURIComponent(fileName)
            },
            body: blob
        });

        const responseText = await response.text();
        let payload: any = {};
        if (responseText) {
            try {
                payload = JSON.parse(responseText);
            } catch (error) {
                payload = { detail: responseText };
            }
        }

        if (!response.ok) {
            const detail = String(payload.error || payload.detail || response.statusText || "Unknown server error");
            throw new Error("PPT export service returned " + response.status + ": " + detail);
        }

        const downloadUrl = String(payload.downloadUrl || "").trim();
        if (!/^https:\/\//i.test(downloadUrl)) {
            throw new Error("PPT export service did not return a valid HTTPS download URL.");
        }

        return {
            fileName: String(payload.fileName || fileName),
            downloadUrl: downloadUrl,
            expiresAt: payload.expiresAt ? String(payload.expiresAt) : undefined,
            sizeBytes: typeof payload.sizeBytes === "number" ? payload.sizeBytes : undefined
        };
    }

    private async downloadPptContent(pres: any, fileName: string): Promise<PptDownloadResult> {
        // Generate once and pass the binary presentation to Power BI's privileged
        // download service first. In Power BI Desktop, PptxGenJS writeFile() can
        // resolve even when the sandbox suppresses the actual browser download.
        const base64 = await this.createPptBase64(pres);
        if (!base64) { throw new Error("The generated PowerPoint content is empty."); }

        const pptBlob = this.base64ToPptBlob(base64);
        const functionUrl = String((config as any).PPT_EXPORT_FUNCTION_URL || "").trim();
        let serviceError = "";

        // Reliable Power BI Desktop path: upload the generated PPTX to an Azure
        // Function, receive a short-lived HTTPS Blob SAS URL, and ask the Power BI
        // host to open that URL in the user's default browser. The blob is uploaded
        // with Content-Disposition: attachment, so the browser downloads the file.
        if (functionUrl) {
            try {
                const serverResult = await this.uploadPptBlobToAzureFunction(
                    pptBlob,
                    fileName,
                    functionUrl
                );
                this.host.launchUrl(serverResult.downloadUrl);
                return {
                    fileName: serverResult.fileName,
                    confirmed: true,
                    method: "server",
                    downloadUrl: serverResult.downloadUrl
                };
            } catch (error) {
                serviceError = error instanceof Error ? error.message : String(error);
                console.warn("[PptExport] Azure Function download bridge failed; trying Power BI fallbacks.", error);
            }
        }
        if (this.downloadService) {
            try {
                let isAllowed = true;
                if (typeof this.downloadService.exportStatus === "function") {
                    const status = await this.downloadService.exportStatus();
                    isAllowed = status === powerbi.PrivilegeStatus.Allowed;
                    if (!isAllowed) {
                        serviceError = "Power BI download permission is not allowed (status: " + String(status) + ").";
                    }
                }

                if (isAllowed && typeof this.downloadService.exportVisualsContentExtended === "function") {
                    const result = await this.downloadService.exportVisualsContentExtended(
                        base64,
                        fileName,
                        "base64",
                        "AI insights PowerPoint presentation"
                    );
                    if (result && result.downloadCompleted === true) {
                        return {
                            fileName: String(result.fileName || result.filename || fileName),
                            confirmed: true,
                            method: "powerbi"
                        };
                    }
                    serviceError = "Power BI did not confirm that the download completed.";
                } else if (isAllowed && typeof this.downloadService.exportVisualsContent === "function") {
                    const completed = await this.downloadService.exportVisualsContent(
                        base64,
                        fileName,
                        "base64",
                        "AI insights PowerPoint presentation"
                    );
                    if (completed === true) {
                        return { fileName: fileName, confirmed: true, method: "powerbi" };
                    }
                    serviceError = "Power BI did not confirm that the download completed.";
                } else if (isAllowed && !serviceError) {
                    serviceError = "The Power BI download service is unavailable.";
                }
            } catch (error) {
                serviceError = error instanceof Error ? error.message : String(error);
                console.warn("[PptExport] Power BI download service failed; trying browser fallback.", error);
            }
        } else {
            serviceError = "The Power BI download service is unavailable.";
        }

        // The async build may have lost browser user activation. Return the real
        // PPTX Blob and let a second direct click on Save PPT perform the save.
        try {
            return {
                fileName: fileName,
                confirmed: false,
                method: "browser",
                fallbackBlob: pptBlob
            };
        } catch (browserError) {
            const browserMessage = browserError instanceof Error ? browserError.message : String(browserError);
            throw new Error(
                [serviceError, browserMessage]
                    .filter(function(value) { return Boolean(value); })
                    .join(" ") || "The PowerPoint could not be prepared for saving."
            );
        }
    }

    private clearPendingPptDownload(): void {
        if (this.pendingPptObjectUrl) {
            try { URL.revokeObjectURL(this.pendingPptObjectUrl); } catch (error) { /* no-op */ }
        }
        this.pendingPptBlob = null;
        this.pendingPptFileName = "";
        this.pendingPptObjectUrl = "";
        this.pendingExportKind = "ppt";
    }

    private preparePendingPptDownload(
        blob: Blob,
        fileName: string,
        kind: "ppt" | "pdf" = "ppt"
    ): void {
        this.clearPendingPptDownload();
        this.pendingPptBlob = blob;
        this.pendingPptFileName = fileName;
        this.pendingExportKind = kind;
        this.pendingPptObjectUrl = URL.createObjectURL(blob);
    }

    private async savePendingPptFromUserGesture(): Promise<void> {
        const blob = this.pendingPptBlob;
        const fileName = this.pendingPptFileName;
        const isPdf = this.pendingExportKind === "pdf";
        const fileLabel = isPdf ? "PDF" : "PowerPoint";
        const saveLabel = isPdf ? "Save PDF" : "Save PPT";
        const mimeType = isPdf
            ? "application/pdf"
            : "application/vnd.openxmlformats-officedocument.presentationml.presentation";
        const extensions = isPdf ? [".pdf"] : [".pptx"];

        if (!blob || !fileName) {
            this.showPptStatus(
                "The prepared file is no longer available. Click Download PPT again.",
                true,
                5500
            );
            this.setDownloadPptButtonState(false);
            return;
        }

        const picker = (window as any).showSaveFilePicker;
        if (typeof picker === "function") {
            try {
                const handle = await picker({
                    suggestedName: fileName,
                    types: [{
                        description: fileLabel + " file",
                        accept: { [mimeType]: extensions }
                    }]
                });
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                this.clearPendingPptDownload();
                this.setDownloadPptButtonState(false);
                this.showPptStatus(fileLabel + " saved successfully: " + fileName + ".", false, 5500);
                return;
            } catch (error) {
                const name = String((error as any) && (error as any).name || "");
                if (name === "AbortError") {
                    this.showPptStatus("Save was cancelled. Click " + saveLabel + " to try again.", false, 5000);
                    this.setDownloadPptButtonState(false);
                    return;
                }
                console.info("[FileExport] Save File Picker unavailable; using direct download fallback.", error);
            }
        }

        const navigatorAny = navigator as any;
        if (typeof navigatorAny.msSaveOrOpenBlob === "function") {
            navigatorAny.msSaveOrOpenBlob(blob, fileName);
            this.clearPendingPptDownload();
            this.setDownloadPptButtonState(false);
            this.showPptStatus(fileLabel + " save dialog opened: " + fileName + ".", false, 5500);
            return;
        }

        try {
            if (!this.pendingPptObjectUrl) {
                this.pendingPptObjectUrl = URL.createObjectURL(blob);
            }
            const ownerDocument = this.target.ownerDocument || document;
            const link = ownerDocument.createElement("a");
            link.href = this.pendingPptObjectUrl;
            link.download = fileName;
            link.target = "_self";
            link.rel = "noopener";
            link.style.display = "none";
            ownerDocument.body.appendChild(link);
            link.click();
            ownerDocument.body.removeChild(link);

            this.showPptStatus(
                "Save requested for " + fileName + ". Check Downloads; click " + saveLabel + " again if Power BI blocked it.",
                false,
                7000
            );
            this.setDownloadPptButtonState(false);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.showPptStatus("Save failed: " + message, true, 8000);
            this.setDownloadPptButtonState(false);
        }
    }

    private setDownloadPptButtonState(isBusy: boolean): void {
        const button = this.target.querySelector(".button_DownloadPpt") as HTMLButtonElement | null;
        if (!button) { return; }
        button.disabled = isBusy;
        button.setAttribute("aria-busy", isBusy ? "true" : "false");
        this.renderDownloadButtonContent(button, isBusy);
    }

    // Builds and downloads a .pptx entirely in the browser, styled to echo the
    // existing PPTX template used by ppt_generator_script_v1-3.py (Arial typeface,
    // bold ~28-30pt titles, ~16-18pt body bullets, 13.333in x 7.5in widescreen,
    // blue/orange/green accent palette) without needing that script's server-side
    // pipeline (win32com, trigger folders, email delivery), none of which is
    // available inside a Power BI custom visual's sandboxed browser context.
    private async blobToBase64(blob: Blob): Promise<string> {
        return new Promise<string>(function(resolve, reject) {
            const reader = new FileReader();
            reader.onload = function(): void {
                const result = String(reader.result || "");
                const commaIndex = result.indexOf(",");
                resolve(commaIndex >= 0 ? result.substring(commaIndex + 1) : result);
            };
            reader.onerror = function(): void {
                reject(reader.error || new Error("Unable to read generated file content."));
            };
            reader.readAsDataURL(blob);
        });
    }

    private buildInsightsPdfBlob(
        responseText: string,
        chartSeries: Array<{ title: string; labels: string[]; values: number[] }>
    ): Blob {
        const pdf: any = new jsPDF({
            orientation: "landscape",
            unit: "pt",
            format: "a4",
            compress: true
        });
        const accent = [215, 25, 32];
        const dark = [34, 34, 34];
        const grey = [89, 89, 89];
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 42;

        const addTopRule = (): void => {
            pdf.setFillColor(accent[0], accent[1], accent[2]);
            pdf.rect(0, 0, pageWidth, 14, "F");
        };
        const addPageTitle = (title: string): void => {
            addTopRule();
            pdf.setTextColor(dark[0], dark[1], dark[2]);
            pdf.setFont("helvetica", "bold");
            pdf.setFontSize(22);
            pdf.text(title, margin, 50);
            pdf.setDrawColor(accent[0], accent[1], accent[2]);
            pdf.setLineWidth(2);
            pdf.line(margin, 62, pageWidth - margin, 62);
        };

        addTopRule();
        pdf.setTextColor(dark[0], dark[1], dark[2]);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(30);
        pdf.text("AI Powered Summary", pageWidth / 2, pageHeight / 2 - 18, { align: "center" });
        pdf.setTextColor(grey[0], grey[1], grey[2]);
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(15);
        pdf.text(this.buildPptSubtitle(), pageWidth / 2, pageHeight / 2 + 18, { align: "center" });

        for (let chartIndex = 0; chartIndex < chartSeries.length; chartIndex++) {
            const series = chartSeries[chartIndex];
            pdf.addPage("a4", "landscape");
            addPageTitle(series.title);

            const labels = series.labels.slice(0, 12);
            const values = series.values.slice(0, 12);
            const chartTop = 88;
            const chartBottom = pageHeight - 38;
            const labelWidth = 245;
            const valueAreaLeft = margin + labelWidth;
            const valueAreaWidth = pageWidth - valueAreaLeft - margin - 48;
            const rowHeight = Math.max(28, Math.min(40, (chartBottom - chartTop) / Math.max(1, labels.length)));

            pdf.setFont("helvetica", "normal");
            pdf.setFontSize(9.5);
            for (let rowIndex = 0; rowIndex < labels.length; rowIndex++) {
                const y = chartTop + rowIndex * rowHeight;
                const value = Math.max(0, Math.min(100, Number(values[rowIndex]) || 0));
                const wrappedLabel = pdf.splitTextToSize(labels[rowIndex], labelWidth - 14);
                pdf.setTextColor(dark[0], dark[1], dark[2]);
                pdf.text(wrappedLabel.slice(0, 2), margin, y + 13);

                pdf.setFillColor(235, 235, 235);
                pdf.roundedRect(valueAreaLeft, y + 2, valueAreaWidth, 17, 4, 4, "F");
                pdf.setFillColor(accent[0], accent[1], accent[2]);
                pdf.roundedRect(valueAreaLeft, y + 2, valueAreaWidth * value / 100, 17, 4, 4, "F");
                pdf.setFont("helvetica", "bold");
                pdf.text(value.toFixed(1) + "%", valueAreaLeft + valueAreaWidth + 8, y + 14);
                pdf.setFont("helvetica", "normal");
            }
        }

        const sections = this.parseInsightIntoSections(responseText);
        const effectiveSections = sections.length > 0
            ? sections
            : [{ heading: "Summary", bullets: [responseText] }];

        for (let sectionIndex = 0; sectionIndex < effectiveSections.length; sectionIndex++) {
            const section = effectiveSections[sectionIndex];
            pdf.addPage("a4", "landscape");
            addPageTitle(section.heading || "Summary");
            let y = 88;
            pdf.setFontSize(12);

            const bullets = section.bullets.length > 0 ? section.bullets : ["No insights available."];
            for (let bulletIndex = 0; bulletIndex < bullets.length; bulletIndex++) {
                const bulletText = String(bullets[bulletIndex] || "").trim();
                const wrapped = pdf.splitTextToSize(bulletText, pageWidth - margin * 2 - 28);
                const requiredHeight = Math.max(24, wrapped.length * 16 + 10);

                if (y + requiredHeight > pageHeight - 36) {
                    pdf.addPage("a4", "landscape");
                    addPageTitle((section.heading || "Summary") + " (continued)");
                    y = 88;
                }

                pdf.setFillColor(accent[0], accent[1], accent[2]);
                pdf.circle(margin + 4, y - 4, 3, "F");
                pdf.setTextColor(dark[0], dark[1], dark[2]);
                pdf.setFont("helvetica", "normal");
                pdf.text(wrapped, margin + 18, y);
                y += requiredHeight;
            }
        }

        return pdf.output("blob") as Blob;
    }

    private async downloadPdfContent(
        responseText: string,
        chartSeries: Array<{ title: string; labels: string[]; values: number[] }>,
        fileName: string
    ): Promise<PdfDownloadResult> {
        const pdfBlob = this.buildInsightsPdfBlob(responseText, chartSeries);
        if (!pdfBlob || pdfBlob.size <= 0) {
            throw new Error("The PDF fallback could not be generated.");
        }

        if (this.downloadService) {
            try {
                let allowed = true;
                if (typeof this.downloadService.exportStatus === "function") {
                    const status = await this.downloadService.exportStatus();
                    allowed = status === powerbi.PrivilegeStatus.Allowed;
                }

                if (allowed) {
                    const base64 = await this.blobToBase64(pdfBlob);
                    if (typeof this.downloadService.exportVisualsContentExtended === "function") {
                        const result = await this.downloadService.exportVisualsContentExtended(
                            base64,
                            fileName,
                            "base64",
                            "AI insights PDF fallback"
                        );
                        if (result && result.downloadCompleted === true) {
                            return {
                                fileName: String(result.fileName || result.filename || fileName),
                                confirmed: true,
                                method: "powerbi"
                            };
                        }
                    } else if (typeof this.downloadService.exportVisualsContent === "function") {
                        const completed = await this.downloadService.exportVisualsContent(
                            base64,
                            fileName,
                            "base64",
                            "AI insights PDF fallback"
                        );
                        if (completed === true) {
                            return { fileName: fileName, confirmed: true, method: "powerbi" };
                        }
                    }
                }
            } catch (error) {
                console.warn("[PdfFallback] Power BI PDF download failed; using browser fallback.", error);
            }
        }

        try {
            this.triggerBrowserBlobDownload(pdfBlob, fileName);
            return {
                fileName: fileName,
                confirmed: false,
                method: "browser",
                fallbackBlob: pdfBlob
            };
        } catch (error) {
            return {
                fileName: fileName,
                confirmed: false,
                method: "browser",
                fallbackBlob: pdfBlob
            };
        }
    }

    private async downloadInsightsAsPpt(): Promise<void> {
        if (this.isDownloadingPpt) { return; }

        const responseText = String(this.gptResponse || "").trim();
        if (!responseText || !this.isValidAiResponse(responseText)) {
            this.showPptStatus("Generate the AI insights before downloading.", true, 4500);
            return;
        }

        this.clearPendingPptDownload();
        this.isDownloadingPpt = true;
        this.setDownloadPptButtonState(true);
        this.showPptStatus("Preparing PowerPoint with charts...", false, 0);

        try {
            const sections = this.parseInsightIntoSections(responseText);
            const chartSeries = this.getChartSeriesFromDataView();

            const ACCENT = "D71920";
            const DARK = "222222";
            const GREY = "595959";

            const pres: any = new (pptxgen as any)();
            pres.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
            pres.layout = "WIDE";
            pres.author = "AI Powered Summary";
            pres.subject = "Power BI AI insights and data charts";
            pres.title = "AI Powered Summary";
            pres.company = "Kantar";
            pres.lang = "en-US";

            const cover = pres.addSlide();
            cover.background = { color: "FFFFFF" };
            cover.addShape(pres.ShapeType ? pres.ShapeType.rect : "rect", {
                x: 0, y: 0, w: 13.333, h: 0.35,
                line: { color: ACCENT, transparency: 100 }, fill: { color: ACCENT }
            });
            cover.addText("AI Powered Summary", {
                x: 0.7, y: 2.25, w: 11.9, h: 0.8,
                fontFace: "Arial", fontSize: 30, bold: true, color: DARK, align: "center"
            });
            cover.addText(this.buildPptSubtitle(), {
                x: 0.8, y: 3.15, w: 11.7, h: 0.5,
                fontFace: "Arial", fontSize: 16, color: GREY, align: "center"
            });

            // Native, editable PowerPoint charts generated from the current Power BI dataView.
            for (let chartIndex = 0; chartIndex < chartSeries.length; chartIndex++) {
                const series = chartSeries[chartIndex];
                const slide = pres.addSlide();
                slide.background = { color: "FFFFFF" };
                slide.addText(series.title, {
                    x: 0.55, y: 0.3, w: 12.2, h: 0.55,
                    fontFace: "Arial", fontSize: 23, bold: true, color: DARK
                });
                slide.addShape(pres.ShapeType ? pres.ShapeType.rect : "rect", {
                    x: 0.55, y: 0.95, w: 12.2, h: 0.04,
                    line: { color: ACCENT, transparency: 100 }, fill: { color: ACCENT }
                });

                const chartType = pres.ChartType && pres.ChartType.bar ? pres.ChartType.bar : "bar";
                slide.addChart(chartType, [{
                    name: series.title,
                    labels: series.labels,
                    values: series.values
                }], {
                    x: 0.7, y: 1.25, w: 11.9, h: 5.75,
                    catAxisLabelFontFace: "Arial",
                    catAxisLabelFontSize: 11,
                    valAxisLabelFontFace: "Arial",
                    valAxisLabelFontSize: 11,
                    showLegend: false,
                    showTitle: false,
                    showValue: true,
                    showCatName: false,
                    showSerName: false,
                    dataLabelPosition: "outEnd",
                    dataLabelFontSize: 10,
                    dataLabelFormatCode: '0.0"%"',
                    chartColors: [ACCENT],
                    showValueAsPercentage: false,
                    valAxisLabelFormatCode: '0.0"%"',
                    valAxisMinVal: 0,
                    valAxisMaxVal: 100,
                    valGridLine: { color: "D9D9D9", size: 1 },
                    showBorder: false
                });
            }

            const slideSections = sections.length > 0
                ? sections
                : [{ heading: "Summary", bullets: [responseText] }];

            for (let sectionIndex = 0; sectionIndex < slideSections.length; sectionIndex++) {
                const section = slideSections[sectionIndex];
                const bullets = section.bullets.length > 0 ? section.bullets : ["No insights available."];
                const chunks: string[][] = [];
                for (let i = 0; i < bullets.length; i += 8) { chunks.push(bullets.slice(i, i + 8)); }

                for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
                    const slide = pres.addSlide();
                    slide.background = { color: "FFFFFF" };
                    const titleSuffix = chunks.length > 1 ? " (" + (chunkIndex + 1) + "/" + chunks.length + ")" : "";
                    const pptHeading = String(section.heading || "Summary").trim().toLowerCase() === "insights"
                        ? "Summary"
                        : String(section.heading || "Summary").trim();
                    slide.addText(pptHeading + titleSuffix, {
                        x: 0.55, y: 0.3, w: 12.2, h: 0.6,
                        fontFace: "Arial", fontSize: 23, bold: true, color: DARK
                    });
                    slide.addShape(pres.ShapeType ? pres.ShapeType.rect : "rect", {
                        x: 0.55, y: 0.95, w: 12.2, h: 0.04,
                        line: { color: ACCENT, transparency: 100 }, fill: { color: ACCENT }
                    });

                    const runs = chunks[chunkIndex].map(function(text: string) {
                        return {
                            text: text,
                            options: {
                                bullet: { indent: 18 },
                                hanging: 4,
                                breakLine: true,
                                fontFace: "Arial",
                                fontSize: 15,
                                color: DARK,
                                paraSpaceAfterPt: 10
                            }
                        };
                    });

                    slide.addText(runs, {
                        x: 0.75, y: 1.25, w: 11.8, h: 5.85,
                        valign: "top",
                        margin: 0.08,
                        breakLine: false,
                        fit: "shrink"
                    });
                }
            }

            const timestamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 19);
            const fileName = "AI_Insights_" + timestamp + ".pptx";
            const downloadResult = await this.downloadPptContent(pres, fileName);
            if (downloadResult.confirmed) {
                this.clearPendingPptDownload();
                const successMessage = downloadResult.method === "server"
                    ? "PowerPoint opened in your default browser for download: " + downloadResult.fileName + "."
                    : "Power BI confirmed the download: " + downloadResult.fileName + ". Check your Downloads folder.";
                this.showPptStatus(
                    successMessage,
                    false,
                    6500
                );
            } else if (downloadResult.fallbackBlob) {
                this.showPptStatus("PowerPoint download was blocked. Preparing PDF fallback...", false, 0);
                const pdfFileName = downloadResult.fileName.replace(/\.pptx$/i, ".pdf");
                const pdfResult = await this.downloadPdfContent(responseText, chartSeries, pdfFileName);

                if (pdfResult.confirmed) {
                    this.clearPendingPptDownload();
                    this.showPptStatus(
                        "PowerPoint was blocked, so Power BI downloaded a PDF instead: " + pdfResult.fileName + ".",
                        false,
                        7000
                    );
                } else if (pdfResult.fallbackBlob) {
                    // PDF is an officially supported Power BI custom-visual export
                    // type. If the host still blocks the automatic save, retain the
                    // generated PDF for a second direct user click.
                    this.preparePendingPptDownload(
                        pdfResult.fallbackBlob,
                        pdfResult.fileName,
                        "pdf"
                    );
                    this.showPptStatus(
                        "PowerPoint was blocked. The PDF is ready; click Save PDF below if it did not appear automatically.",
                        false,
                        0
                    );
                } else {
                    this.preparePendingPptDownload(
                        downloadResult.fallbackBlob,
                        downloadResult.fileName,
                        "ppt"
                    );
                    this.showPptStatus(
                        "PowerPoint and PDF downloads were blocked. Click Save PPT below to try the prepared PowerPoint directly.",
                        false,
                        0
                    );
                }
            } else {
                throw new Error("The PowerPoint was generated, but no downloadable file content was returned.");
            }
        } catch (error) {
            console.error("[PptExport] Failed to generate/download PPTX:", error);
            const message = error instanceof Error ? error.message : String(error);
            this.showPptStatus("Download failed: " + message, true, 8000);
        } finally {
            this.isDownloadingPpt = false;
            this.setDownloadPptButtonState(false);
        }
    }

    // private applyInlineFormatting(escapedText: string): string {
    //     let text = escapedText;
    //     text = text.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    //     text = text.replace(/([+-]?\d+(?:\.\d+)?%)/g, "<span class=\"ai-percent\">$1</span>");
    //     text = text.replace(/\b(Brand\s+\d+)\b/g, "<strong class=\"ai-brand\">$1</strong>");
    //     return text;
    // }

    // private applyInlineFormatting(escapedText: string): string {
    // let text = escapedText;
 
    // text = text.replace(/([+-]?\d+(?:\.\d+)?%)/g, "<span class=\"ai-percent\">$1</span>");
 
    // const values = (this.highlightedCategoryValuesFromPrompt || [])
    //     .filter(function(value) {
    //         return value && String(value).trim().length > 0;
    //     })
    //     .sort(function(a, b) {
    //         return String(b).length - String(a).length;
    //     });
 
    // for (let i = 0; i < values.length; i++) {
    //     const rawValue = String(values[i]).trim();
    //     if (!rawValue) { continue; }
 
    //     const escapedValue = this.escapeHtml(rawValue);
    //     const pattern = new RegExp(this.escapeRegExp(escapedValue), "g");
 
    //     text = text.replace(pattern, "<strong class=\"ai-dynamic-highlight\">" + escapedValue + "</strong>");
    // }
 
    // return text;
    // }

    private applyInlineFormatting(escapedText: string): string {
    let text = escapedText;
 
    // const values = (this.highlightedCategoryValuesFromPrompt || [])
    const values = (((this.highlightedCategoryFieldsFromPrompt || []) as string[])

        .concat(this.highlightedCategoryValuesFromPrompt || []))
        .filter(function(value) {
            return value && String(value).trim().length > 0;
        })
        .sort(function(a, b) {
            return String(b).length - String(a).length;
        });
 
    for (let i = 0; i < values.length; i++) {
        const rawValue = String(values[i]).trim();
 
        if (!rawValue) {
            continue;
        }
 
        const escapedValue = this.escapeHtml(rawValue);
        const pattern = new RegExp(this.escapeRegExp(escapedValue), "g");

        text = text.replace(
            pattern,
            function(m: string) { return "<strong class=\"ai-dynamic-highlight\">" + m + "</strong>"; }
        );
    }
 
    text = text.replace(
        /([+-]?\d+(?:\.\d+)?%)/g,
        "<span class=\"ai-percent\">$1</span>"
    );
 
    return text;
}
 

    private escapeRegExp(value: string): string {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
 
 

    private escapeHtml(value: string): string {
        return String(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    private normaliseCacheValue(value: any): string {
        if (value === null || value === undefined) { return ""; }
        if (typeof value === "number") {
            if (!isFinite(value)) { return ""; }
            return String(Number(value.toFixed(10)));
        }
        return String(value).replace(/\s+/g, " ").trim();
    }

    private hashCacheText(value: string): string {
        const input = String(value || "");
        let hashOne = 2166136261;
        let hashTwo = 2166136261;

        for (let i = 0; i < input.length; i++) {
            hashOne ^= input.charCodeAt(i);
            hashOne = Math.imul(hashOne, 16777619);
        }
        for (let i = input.length - 1; i >= 0; i--) {
            hashTwo ^= input.charCodeAt(i);
            hashTwo = Math.imul(hashTwo, 16777619);
        }

        const first = (hashOne >>> 0).toString(16).padStart(8, "0");
        const second = (hashTwo >>> 0).toString(16).padStart(8, "0");
        return first + second;
    }

    private buildCurrentDataFingerprint(): string {
        const categorical = this.dataView && this.dataView.categorical;
        if (!categorical) { return "NO_DATA"; }

        const rawCategories = categorical.categories || [];
        const rawValues: powerbi.DataViewValueColumn[] = categorical.values
            ? Array.from(categorical.values) as powerbi.DataViewValueColumn[]
            : [];

        // Fingerprint only analytical dimensions and the row-level Metric field.
        // Highlight Entity is presentation-only, so adding or removing a highlight
        // assignment must not invalidate the analytical cache or call Azure again.
        // When the same source field is present in Category Data / Metric and also in
        // Highlight Entity, identity-based de-duplication keeps one analytical copy.
        const split = this.splitCategoriesByRoles(rawCategories);
        const analyticalCategories = this.getAnalyticalCategoryColumns(rawCategories);
        const analyticalMetricCategories = split.metricCategories.filter((category) => {
            const source: any = category && category.source ? category.source : {};
            const roles: any = source.roles || {};
            const roleNames = Object.keys(roles).filter(function(roleName) {
                return roles[roleName] === true;
            });
            const isHighlightOnly = roleNames.length > 0 &&
                roleNames.indexOf("highlightEntity") !== -1 &&
                roleNames.every(function(roleName) {
                    return roleName === "highlightEntity";
                });
            return !isHighlightOnly;
        });
        const fingerprintCategoryCandidates = analyticalCategories.concat(analyticalMetricCategories);

        const categoryMap: { [key: string]: powerbi.DataViewCategoryColumn } = {};
        for (let i = 0; i < fingerprintCategoryCandidates.length; i++) {
            const category = fingerprintCategoryCandidates[i];
            const key = this.getCategoryIdentityKey(category, i);
            const existing = categoryMap[key];
            const existingLength = existing && existing.values ? existing.values.length : -1;
            const candidateLength = category.values ? category.values.length : 0;
            if (!existing || candidateLength > existingLength) { categoryMap[key] = category; }
        }

        const valueMap: { [key: string]: powerbi.DataViewValueColumn } = {};
        for (let i = 0; i < rawValues.length; i++) {
            const valueColumn = rawValues[i];
            const source: any = valueColumn && valueColumn.source ? valueColumn.source : {};
            const key = this.normalisePromptFieldToken(
                String(source.queryName || source.displayName || ("measure_" + i))
            ) || ("measure_" + i);
            const existing = valueMap[key];
            const existingLength = existing && existing.values ? existing.values.length : -1;
            const candidateLength = valueColumn.values ? valueColumn.values.length : 0;
            if (!existing || candidateLength > existingLength) { valueMap[key] = valueColumn; }
        }

        const categoryKeys = Object.keys(categoryMap).sort();
        const valueKeys = Object.keys(valueMap).sort();
        let rowCount = 0;

        for (let i = 0; i < categoryKeys.length; i++) {
            const column = categoryMap[categoryKeys[i]];
            rowCount = Math.max(rowCount, column.values ? column.values.length : 0);
        }
        for (let i = 0; i < valueKeys.length; i++) {
            const column = valueMap[valueKeys[i]];
            rowCount = Math.max(rowCount, column.values ? column.values.length : 0);
        }

        const headers = categoryKeys.map(function(key) { return "C:" + key; })
            .concat(valueKeys.map(function(key) { return "M:" + key; }));
        const rows: string[] = [];

        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            const rowParts: string[] = [];
            for (let c = 0; c < categoryKeys.length; c++) {
                const column = categoryMap[categoryKeys[c]];
                const value = column.values && column.values.length > rowIndex ? column.values[rowIndex] : null;
                rowParts.push(this.normaliseCacheValue(value));
            }
            for (let v = 0; v < valueKeys.length; v++) {
                const column = valueMap[valueKeys[v]];
                const value = column.values && column.values.length > rowIndex ? column.values[rowIndex] : null;
                rowParts.push(this.normaliseCacheValue(value));
            }
            rows.push(rowParts.join("\u001f"));
        }

        rows.sort();
        return this.hashCacheText(headers.join("\u001f") + "\u001e" + rows.join("\u001e"));
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // private buildPromptFromData(): { basePrompt: string; fullPrompt: string; selectionKey: string; } {
    //     const defaultPrompt =
    //         "You are a helpful market research analyst and data analyst. Summarise the filtered data and give important highlights.";

    //     if (!this.dataView || !this.dataView.categorical || !this.dataView.categorical.categories) {
    //         return { basePrompt: defaultPrompt, fullPrompt: defaultPrompt, selectionKey: "NO_DATA" };
    //     }

    //     const categories = this.dataView.categorical.categories || [];
    //     if (categories.length === 0) {
    //         return { basePrompt: defaultPrompt, fullPrompt: defaultPrompt, selectionKey: "NO_CATEGORIES" };
    //     }

    //     const split = this.splitCategoriesByRoles(categories);
    //     const promptInstruction = this.extractPromptInstructionFromPromptCategories(split.promptCategories);
    //     const basePrompt = promptInstruction || defaultPrompt;
    //     this.preferredEntityFieldFromPrompt = this.inferPreferredEntityFieldFromPrompt(basePrompt, split.dataCategories);
    //     const selectionCategories = split.dataCategories.concat(split.metricCategories);

    //     const selectionSummary: string[] = [];
    //     const selectionKeyObject: {
    //         prompt: string;
    //         selections: Array<{ column: string; values: string[] }>;
    //         measures: string[];
    //     } = { prompt: basePrompt, selections: [], measures: [] };

    //     for (let i = 0; i < selectionCategories.length; i++) {
    //         const category = selectionCategories[i];
    //         const columnName = this.getColumnDisplayName(category.source, "Unknown");
    //         const uniqueValuesForSummary = this.getUniqueNonBlankValues(category.values, this.MAX_CONTEXT_VALUES_PER_FIELD);
    //         const uniqueValuesForKey = this.getUniqueNonBlankValues(category.values);

    //         if (uniqueValuesForSummary.length > 0) {
    //             selectionSummary.push(columnName + ": " + uniqueValuesForSummary.join(", "));
    //         }

    //         selectionKeyObject.selections.push({ column: columnName, values: uniqueValuesForKey });
    //     }

    //     const categoricalValues = this.dataView.categorical.values;
    //     const measureNames: string[] = categoricalValues
    //         ? Array.from(categoricalValues).map((v) => this.getColumnDisplayName(v.source, "Measure"))
    //         : [];
    //     selectionKeyObject.measures = measureNames;

    //     const fullPrompt =
    //         "USER PROVIDED PAGE / BUSINESS INSTRUCTION:\n" +
    //         basePrompt +
    //         "\n\nIMPORTANT:\n" +
    //         "- The instruction above is the highest-priority business instruction.\n" +
    //         "- Use it to decide the correct analysis focus/entity level.\n" +
    //         "- Do not assume the analysis must be brand-level unless the instruction says so.\n" +
    //         "- If the instruction says multiple countries/markets, focus on country/market-level insights.\n" +
    //         "- If the instruction says multiple brands, focus on brand-level insights.\n" +
    //         "- If the instruction says demographics, focus on demographic-level insights.\n" +
    //         "- Scores must come only from Measure Data numeric fields.\n" +
    //         "- Metrics must come only from the Metric / KPI Breakout field when that field is provided.\n" +
    //         "- Category Data fields are context/breakdowns only, even when their values are numeric.\n" +
    //         "\n\nCurrent selection context:\n" +
    //         (selectionSummary.length > 0 ? selectionSummary.join("\n") : "No explicit selections detected.") +
    //         "\n\nOutput formatting rules:\n" +
    //         "- Use short section headings.\n" +
    //         "- Use bullet points only.\n" +
    //         "- Put each main entity insight on a separate bullet.\n" +
    //         "- Avoid long paragraphs.\n" +
    //         "- Do not return markdown tables.\n" +
    //         "- Use this structure exactly:\n" +
    //         "  Summary:\n" +
    //         "  - one-line overall summary\n" +
    //         "  Key findings:\n" +
    //         "  - Entity X: finding with value\n" +
    //         "  - Entity Y: finding with value\n" +
    //         "  Top performers:\n" +
    //         "  - Entity X: reason\n" +
    //         "  Watchouts:\n" +
    //         "  - Entity Y: reason\n" +
    //         "  Recommendation:\n" +
    //         "  - action-oriented recommendation\n" +
    //         "\nAnalysis rules:\n" +
    //         "- Analyse only the currently filtered data.\n" +
    //         "- Compare the relevant main entities clearly based on the Prompt as text and current data fields.\n" +
    //         "- Mention important highs, lows, drops, strengths, weaknesses, significance, and unusual movements.\n" +
    //         "- If the dataset is previewed or truncated, mention it briefly.";

    //     return {
    //         basePrompt: basePrompt,
    //         fullPrompt: fullPrompt,
    //         selectionKey: JSON.stringify(selectionKeyObject)
    //     };
    // }

    // Default prompt used whenever no custom instruction has been entered in the
    // "Prompt" field well. This intentionally mirrors the market-research analyst use
    // case: it states the analyst persona, the data-only / no-recommendations rules,
    // and - critically - an explicit "Output format:" block with real heading labels
    // and bullet placeholders. extractHeadingsFromPrompt() reads that block directly,
    // so "Summary" and "Key findings" are recognised as headings the same way any
    // custom heading names typed into the Prompt field would be.
    private readonly DEFAULT_ANALYST_PROMPT: string =
        "You are a market research analyst with strong expertise in customer journey and subgroup analysis. Accuracy is very important.\n\n" +
        "The dataset contains one or more Category Data dimensions (for example Market, Category, Brand, Gender, Age, Income, Education, or Time Period), a Metric field, and one or more numeric Value measures.\n\n" +
        "Task: generate a data-only summary of the currently filtered dataset.\n\n" +
        "Important instructions:\n" +
        "1. Do not give recommendations, actions, next steps, or strategic suggestions.\n" +
        "2. Use exact visible values instead of generic phrases when specific values are available.\n" +
        "3. Preserve every visible Category Data dimension. Never average observations that differ on Gender, Age, Income, Education, Market, Category, Brand, Time Period, or any other Category Data field.\n" +
        "4. Treat rows as duplicates only when ALL visible Category Data values and the Metric are identical; average only those exact duplicates.\n" +
        "5. If a Category Data field has more than one visible value, compare those values separately where relevant. Report Male and Female separately rather than averaging them.\n" +
        "6. All percentage scores must be percentages rounded to one decimal place.\n" +
        "7. Use only the data provided. Do not assume missing dimensions, entities, periods, or trends.\n" +
        "8. If only one time period is available, describe differences, gaps, highs, lows, and relative performance rather than trends over time.\n" +
        "9. Every high, low, or comparison must state the exact Metric, Score, and every relevant visible dimension value.\n\n" +
        "Output format:\n\n" +
        "Summary\n" +
        "- [Bullet 1]\n" +
        "- [Bullet 2]\n" +
        "- [Bullet 3]\n\n" +
        "Key findings\n" +
        "- [Bullet 1]\n" +
        "- [Bullet 2]\n" +
        "- [Bullet 3]\n\n" +
        "Style requirements:\n" +
        "- Use clear, concise, professional language.\n" +
        "- Use bullet points only, each starting with '- '.\n" +
        "- Do not include tables or methodology notes.";

    private buildPromptFromData(): { basePrompt: string; fullPrompt: string; selectionKey: string; basePromptIsDefault: boolean; } {
        const defaultPrompt = this.DEFAULT_ANALYST_PROMPT;

        if (!this.dataView || !this.dataView.categorical || !this.dataView.categorical.categories) {
            return { basePrompt: defaultPrompt, fullPrompt: defaultPrompt, selectionKey: "NO_DATA", basePromptIsDefault: true };
        }

        const categories = this.dataView.categorical.categories || [];
        if (categories.length === 0) {
            return { basePrompt: defaultPrompt, fullPrompt: defaultPrompt, selectionKey: "NO_CATEGORIES", basePromptIsDefault: true };
        }

        const split = this.splitCategoriesByRoles(categories);
        const analyticalCategories = this.getAnalyticalCategoryColumns(categories);
        const descriptors = this.getAnalyticalDimensionDescriptors(categories);
        const promptInstruction = this.extractPromptInstructionFromPromptCategories(split.promptCategories);
        const basePrompt = promptInstruction || defaultPrompt;
        const basePromptIsDefault = !promptInstruction;

        this.syncPromptOutputFormatting(
            basePrompt
        );

        this.preferredEntityFieldFromPrompt = this.inferPreferredEntityFieldFromPrompt(
            basePrompt,
            analyticalCategories
        );

        const highlightValues: string[] = [];
        const highlightFields: string[] = [];
        const seenHighlightValue: { [key: string]: boolean } = {};
        const seenHighlightField: { [key: string]: boolean } = {};

        for (let i = 0; i < split.highlightCategories.length; i++) {
            const category = split.highlightCategories[i];
            const fieldName = this.getColumnDisplayName(category.source, "Highlight Entity");
            const fieldKey = this.normalisePromptFieldToken(fieldName);
            if (fieldName && !seenHighlightField[fieldKey]) {
                seenHighlightField[fieldKey] = true;
                highlightFields.push(fieldName);
            }
            const values = this.getUniqueNonBlankValues(
                category.values,
                this.MAX_CONTEXT_VALUES_PER_FIELD
            );
            for (let valueIndex = 0; valueIndex < values.length; valueIndex++) {
                const value = String(values[valueIndex] || "").trim();
                if (!value || value === "... more selected") { continue; }
                const valueKey = value.toLowerCase();
                if (!seenHighlightValue[valueKey]) {
                    seenHighlightValue[valueKey] = true;
                    highlightValues.push(value);
                }
            }
        }

        if (split.highlightCategories.length === 0) {
            for (let i = 0; i < analyticalCategories.length; i++) {
                const category = analyticalCategories[i];
                const fieldName = this.getColumnDisplayName(category.source, "");
                const token = this.normalisePromptFieldToken(fieldName);
                if (token.indexOf("brand") === -1 && token.indexOf("market") === -1 && token.indexOf("country") === -1) {
                    continue;
                }
                if (!seenHighlightField[token]) {
                    seenHighlightField[token] = true;
                    highlightFields.push(fieldName);
                }
                const values = this.getUniqueNonBlankValues(
                    category.values,
                    this.MAX_CONTEXT_VALUES_PER_FIELD
                );
                for (let valueIndex = 0; valueIndex < values.length; valueIndex++) {
                    const value = String(values[valueIndex] || "").trim();
                    if (!value || value === "... more selected") { continue; }
                    const valueKey = value.toLowerCase();
                    if (!seenHighlightValue[valueKey]) {
                        seenHighlightValue[valueKey] = true;
                        highlightValues.push(value);
                    }
                }
                break;
            }
        }

        this.highlightedCategoryFieldsFromPrompt = highlightFields;
        this.highlightedCategoryValuesFromPrompt = highlightValues;

        const categorySummary: string[] = [];
        const metricSummary: string[] = [];
        const selectionKeyObject: {
            prompt: string;
            selections: Array<{ column: string; values: string[]; role: string }>;
            measures: string[];
        } = {
            prompt: basePrompt,
            selections: [],
            measures: []
        };

        for (let i = 0; i < analyticalCategories.length; i++) {
            const category = analyticalCategories[i];
            const columnName = this.getColumnDisplayName(category.source, "Dimension_" + (i + 1));
            const valuesForSummary = this.getUniqueNonBlankValues(
                category.values,
                this.MAX_CONTEXT_VALUES_PER_FIELD
            );
            const valuesForKey = this.getUniqueNonBlankValues(category.values);
            if (valuesForSummary.length > 0) {
                categorySummary.push(columnName + ": " + valuesForSummary.join(", "));
            }
            selectionKeyObject.selections.push({
                column: columnName,
                values: valuesForKey,
                role: valuesForKey.length > 1 ? "Comparison Category Data" : "Context Category Data"
            });
        }

        for (let i = 0; i < split.metricCategories.length; i++) {
            const category = split.metricCategories[i];
            const columnName = this.getColumnDisplayName(category.source, "Metric / KPI");
            const valuesForSummary = this.getUniqueNonBlankValues(
                category.values,
                this.MAX_CONTEXT_VALUES_PER_FIELD
            );
            const valuesForKey = this.getUniqueNonBlankValues(category.values);
            if (valuesForSummary.length > 0) {
                metricSummary.push(columnName + ": " + valuesForSummary.join(", "));
            }
            selectionKeyObject.selections.push({
                column: columnName,
                values: valuesForKey,
                role: "Metric / KPI Breakout"
            });
        }

        const categoricalValues = this.dataView.categorical.values;
        const measureNames: string[] = categoricalValues
            ? Array.from(categoricalValues).map((valueColumn: powerbi.DataViewValueColumn) => {
                return this.getColumnDisplayName(valueColumn.source, "Measure");
            })
            : [];
        selectionKeyObject.measures = measureNames;

        const fullPrompt =
            "USER PROVIDED PAGE / BUSINESS INSTRUCTION:\n" +
            basePrompt +
            "\n\nSYSTEM BEHAVIOR RULES:\n" +
            "- The user instruction is the highest-priority business instruction.\n" +
            "- Analyse the complete currently filtered dataset.\n" +
            "- Highlight Entity controls only red/bold emphasis and never analysis, grouping, averaging, comparisons, or cache.\n" +
            "- A field in both Category Data and Highlight Entity remains a full analytical dimension.\n" +
            "- Keep differently named metrics separate.\n" +
            "- Never output a missing/unavailable placeholder when an explicit field value exists.\n" +
            this.buildAnalyticalGrainInstruction(descriptors, basePrompt) +
            "\nCurrent Category Data values:\n" +
            (categorySummary.length > 0
                ? categorySummary.map(function(item) { return "- " + item; }).join("\n")
                : "- No Category Data values detected.") +
            "\n\nCurrent Metric / KPI values:\n" +
            (metricSummary.length > 0
                ? metricSummary.map(function(item) { return "- " + item; }).join("\n")
                : "- No explicit Metric / KPI values detected.") +
            "\n\nIf the user requests a specific output structure, follow it exactly.";

        const canonicalSelections = selectionKeyObject.selections
            .map(function(selection) {
                return {
                    column: String(selection.column || "").trim(),
                    values: selection.values.slice().sort(function(left, right) {
                        return left.localeCompare(right, undefined, { sensitivity: "base" });
                    }),
                    role: selection.role
                };
            })
            .sort(function(left, right) {
                return (left.role + "::" + left.column).localeCompare(
                    right.role + "::" + right.column,
                    undefined,
                    { sensitivity: "base" }
                );
            });

        const canonicalKeyPayload = JSON.stringify({
            version: 17,
            promptHash: this.hashCacheText(basePrompt.replace(/\s+/g, " ").trim()),
            selections: canonicalSelections,
            measures: measureNames.slice().sort(function(left, right) {
                return left.localeCompare(right, undefined, { sensitivity: "base" });
            }),
            dataFingerprint: this.buildCurrentDataFingerprint()
        });

        return {
            basePrompt: basePrompt,
            fullPrompt: fullPrompt,
            selectionKey: "v17_" + this.hashCacheText(canonicalKeyPayload),
            basePromptIsDefault: basePromptIsDefault
        };
    }

    private getUniqueNonBlankValues(vals: powerbi.PrimitiveValue[] | undefined, limit?: number): string[] {
        if (!vals) { return []; }

        const seen: { [key: string]: boolean } = {};
        const uniqueValues: string[] = [];

        for (let i = 0; i < vals.length; i++) {
            const str = String(vals[i] == null ? "" : vals[i]).trim();
            if (str !== "" && !seen[str]) {
                seen[str] = true;
                uniqueValues.push(str);
            }
        }

        if (typeof limit === "number" && uniqueValues.length > limit) {
            const displayed = uniqueValues.slice(0, limit);
            displayed.push("... more selected");
            return displayed;
        }

        return uniqueValues;
    }

    private extractTextFromAiResponse(data: any): string {
        try {
            if (data && data.choices && data.choices.length > 0) {
                const choice = data.choices[0];

                if (choice.message && typeof choice.message.content === "string") {
                    return choice.message.content.trim();
                }

                if (choice.message && Array.isArray(choice.message.content)) {
                    const parts = choice.message.content
                        .map(function(part: any) {
                            if (typeof part === "string")              { return part; }
                            if (part && typeof part.text === "string") { return part.text; }
                            if (part && typeof part.content === "string") { return part.content; }
                            return "";
                        })
                        .join("\n")
                        .trim();
                    if (parts) { return parts; }
                }

                if (typeof choice.text === "string") { return choice.text.trim(); }
            }

            if (data && typeof data.output_text === "string") { return data.output_text.trim(); }

            if (data && Array.isArray(data.output)) {
                const outputText = data.output
                    .map(function(item: any) {
                        if (typeof item === "string") { return item; }
                        if (item && typeof item.content === "string") { return item.content; }
                        if (item && Array.isArray(item.content)) {
                            return item.content
                                .map(function(part: any) {
                                    if (typeof part === "string")              { return part; }
                                    if (part && typeof part.text === "string") { return part.text; }
                                    return "";
                                })
                                .join("\n");
                        }
                        return "";
                    })
                    .join("\n")
                    .trim();
                if (outputText) { return outputText; }
            }
        } catch (error) {
            console.error("Error extracting AI text:", error);
        }

        return "";
    }

    private getCleanErrorMessage(error: any): string {
        console.error("Detailed Error:", error);
        const message = error && error.message ? String(error.message) : "";

        if (message.indexOf("401") !== -1 || message.indexOf("403") !== -1) {
            return "Authentication error. Please check API configuration.";
        }
        if (message.indexOf("429") !== -1) {
            return "Too many requests. Please try again later.";
        }
        if (message.indexOf("500") !== -1) {
            return "Server error. Please try again.";
        }
        if (message.indexOf("AI service returned an empty message") !== -1) {
            return (
                "Summary:\n" +
                "- The selected data was processed, but the AI service returned an empty message. This result was not cached.\n\n" +
                "Recommendation:\n" +
                "- Click Show again to retry. If it repeats, check the browser console for the raw Azure response."
            );
        }
        if (message.indexOf("parse AI response") !== -1) {
            return "The AI service returned an unreadable response. Please try again.";
        }
        if (message.indexOf("Unexpected response format") !== -1) {
            return "Received an unexpected response from the AI service.";
        }

        return "Unable to generate insights. Please try again.";
    }

    private isStructuralMessage(message: string): boolean {
        return (
            message === "No data available"             ||
            message === "Not enough category columns"   ||
            message === "Insufficient category structure" ||
            message === "No metric values available"    ||
            message === "No valid rows after filtering"
        );
    }

    private intelligentlyReduceDataset(
        totalRows: number,
        dimensionCategories: powerbi.DataViewCategoryColumn[],
        valueColumns: powerbi.DataViewValueColumns,
        maxRows: number
    ): number[] {
        if (totalRows <= maxRows) {
            const result: number[] = [];
            for (let i = 0; i < totalRows; i++) { result.push(i); }
            return result;
        }

        const selected = new Set<number>();

        const addIndex = function(index: number): void {
            if (index >= 0 && index < totalRows && selected.size < maxRows) {
                selected.add(index);
            }
        };

        const recentRowsToKeep = Math.min(40, Math.floor(maxRows * 0.2));
        for (let i = Math.max(0, totalRows - recentRowsToKeep); i < totalRows; i++) { addIndex(i); }

        const diversityTarget    = Math.min(80, Math.floor(maxRows * 0.4));
        const seenDimensionKeys  = new Set<string>();
        for (let i = 0; i < totalRows && selected.size < diversityTarget; i++) {
            const key = dimensionCategories
                .map(function(category) {
                    const vals = category.values || [];
                    return String(vals[i] == null ? "" : vals[i]);
                })
                .join("||");
            if (!seenDimensionKeys.has(key)) { seenDimensionKeys.add(key); addIndex(i); }
        }

        const scoredRows: Array<{ index: number; score: number }> = [];
        for (let i = 0; i < totalRows; i++) {
            let score = 0;
            Array.from(valueColumns).forEach((valueColumn: powerbi.DataViewValueColumn) => {
                const rawValue = valueColumn.values && valueColumn.values.length > i ? valueColumn.values[i] : null;
                const numericVal = this.toNumber(rawValue);
                if (numericVal !== null) { score += Math.abs(numericVal); }
            });
            scoredRows.push({ index: i, score: score });
        }

        const sortedHigh       = scoredRows.slice().sort(function(a, b) { return b.score - a.score; });
        const sortedLow        = scoredRows.slice().sort(function(a, b) { return a.score - b.score; });
        const numericRowsToKeep = Math.min(80, Math.floor(maxRows * 0.4));

        for (let i = 0; i < sortedHigh.length && selected.size < maxRows && i < numericRowsToKeep / 2; i++) {
            addIndex(sortedHigh[i].index);
        }
        for (let i = 0; i < sortedLow.length && selected.size < maxRows && i < numericRowsToKeep / 2; i++) {
            addIndex(sortedLow[i].index);
        }

        if (selected.size < maxRows) {
            const remainingSlots = maxRows - selected.size;
            const step           = Math.max(1, Math.floor(totalRows / remainingSlots));
            for (let i = 0; i < totalRows && selected.size < maxRows; i += step) { addIndex(i); }
        }

        return Array.from(selected).sort(function(a, b) { return a - b; });
    }

    private toNumber(value: powerbi.PrimitiveValue | null | undefined): number | null {
        if (value === null || value === undefined) { return null; }
        if (typeof value === "number" && isFinite(value)) { return value; }
        const cleaned = String(value).replace(/,/g, "").replace(/%/g, "").trim();
        if (!cleaned) { return null; }
        const parsed = Number(cleaned);
        return isFinite(parsed) ? parsed : null;
    }

    private formatPivotedDataDynamic(): string {
        const full = this.formatFullPivotedData();
        const lines = String(full || "").split("\n").filter(function(line) { return line.trim().length > 0; });
        if (lines.length <= this.MAX_DATA_ROWS + 1) { return full; }
        return lines.slice(0, this.MAX_DATA_ROWS + 1).join("\n") + "\nNOTE: Previewed " + this.MAX_DATA_ROWS + " rows out of " + (lines.length - 1) + " formatted rows.";
    }



}
