/*
 *  Power BI Visual — Atlyn Calendar Slicer
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import VisualObjectInstancesToPersist = powerbi.VisualObjectInstancesToPersist;

import { VisualFormattingSettingsModel } from "./settings";
import {
    buildDayFilter,
    buildMultiDayFilter,
    buildRangeFilter,
    FilterTarget,
    targetFromQueryName
} from "./dateFilter";
import { PRESETS, PresetContext, presetByKey } from "./presets";
import {
    addDays,
    addMonths,
    buildMonthGrid,
    CalendarDay,
    getISOWeek,
    isSameDay,
    makeDate,
    parseDateWithoutTimezone,
    startOfDay,
    startOfWeek,
    WeekStart
} from "./utils/dateMath";

type Selection =
    | { type: "none" }
    | { type: "range"; start: Date; end: Date }
    | { type: "days"; days: Date[] };

// FilterAction is an ambient const enum in the API typings (merge = 0,
// remove = 1). esbuild (and other bundlers that do not inline ambient .d.ts
// const enums) would otherwise leave a runtime reference to an undefined
// object, so the values are pinned here as typed literals. Do NOT "simplify"
// this back into `import { FilterAction }` — that breaks at test/runtime. The
// literal-member types (`.merge`/`.remove`) fail compilation if the values
// ever drift from the API. Matches the fleet convention (radar/gantt).
const MERGE_FILTER_ACTION: powerbi.FilterAction.merge = 0;
const REMOVE_FILTER_ACTION: powerbi.FilterAction.remove = 1;

// Must match the categorical `top` count in capabilities.json. When the
// received category reaches this length the table may have been truncated, so
// data-completeness features (e.g. "grey days without data") are suppressed to
// avoid mislabelling real days as empty.
const DATA_REDUCTION_COUNT = 30000;

interface VisibleMonth {
    year: number;
    month: number;
}

export class Visual implements IVisual {
    private readonly target: HTMLElement;
    private readonly root: HTMLElement;
    private readonly host: IVisualHost;
    private readonly selectionManager: ISelectionManager;
    private readonly localizationManager: ILocalizationManager;
    private readonly formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;

    /**
     * Theme-derived state. Resolved live on every {@link update} (never cached
     * across updates) so the visual reacts to report-theme swaps and Windows
     * high-contrast toggles without waiting to be recreated.
     */
    private isHighContrast = false;
    private hcForeground = "#000000";
    private hcBackground = "#ffffff";
    /**
     * False in non-interactive host contexts (PowerPoint export, email
     * subscriptions). When false the slicer renders read-only: no click/drag/
     * keyboard filtering, no focusable cells, no hover affordances.
     */
    private interactive = true;
    private readonly locale: string;

    private filterTarget: FilterTarget | null = null;
    private dataMin: Date | null = null;
    private dataMax: Date | null = null;
    /** Per-day aggregated measure values, keyed by day key. */
    private dataValues = new Map<string, number>();
    private valueMin = 0;
    private valueMax = 0;
    private hasValues = false;
    /**
     * True when the inbound category hit the capabilities data-reduction cap,
     * so the received dates may be an incomplete subset of the real table.
     */
    private dataTruncated = false;

    private selection: Selection = { type: "none" };
    private activePreset: string | null = null;
    private visible: VisibleMonth | null = null;
    private focusedDate: Date | null = null;

    /** Drag state for mouse range selection. */
    private dragAnchor: Date | null = null;
    private isDragging = false;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.locale = this.host.locale || "en-US";
        this.selectionManager = this.host.createSelectionManager();
        this.localizationManager = this.host.createLocalizationManager();
        this.formattingSettingsService =
            new FormattingSettingsService(this.localizationManager);
        this.formattingSettings = new VisualFormattingSettingsModel();

        this.root = document.createElement("div");
        this.root.className = "atlynCalendarSlicer";
        this.target.appendChild(this.root);

        // End any drag even if the pointer is released outside a day cell.
        this.root.addEventListener("pointerup", () => this.endDrag());
        this.root.addEventListener("pointerleave", () => this.endDrag());
    }

    /**
     * Resolve live theme colours for this update only. Reading these on every
     * update (rather than caching them in the constructor) is what lets the
     * visual respond to report-theme changes and high-contrast toggles while
     * the report stays open.
     */
    private resolveThemeColors(): void {
        this.isHighContrast = this.host.colorPalette.isHighContrast === true;
        const palette = this.host.colorPalette as unknown as {
            foreground?: { value?: string };
            background?: { value?: string };
        };
        this.hcForeground = palette.foreground?.value || "#000000";
        this.hcBackground = palette.background?.value || "#ffffff";

        this.root.classList.toggle("high-contrast", this.isHighContrast);
        if (this.isHighContrast) {
            this.root.style.color = this.hcForeground;
            this.root.style.background = this.hcBackground;
        } else {
            this.root.style.removeProperty("color");
            this.root.style.removeProperty("background");
        }
    }

    public update(options: VisualUpdateOptions): void {
        this.host.eventService?.renderingStarted(options);
        try {
            this.resolveThemeColors();
            this.interactive = this.host.hostCapabilities?.allowInteractions !== false;
            this.root.classList.toggle("read-only", !this.interactive);
            const dataView: DataView | undefined = options.dataViews?.[0];
            this.formattingSettings =
                this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    dataView
                );

            const category = dataView?.categorical?.categories?.[0];
            if (!category) {
                this.filterTarget = null;
                this.renderLanding(
                    this.localize("Landing_AddField", "Add a date field to the Date bucket")
                );
                this.host.eventService?.renderingFinished(options);
                return;
            }

            if (!this.isDateColumn(category)) {
                this.filterTarget = null;
                this.renderLanding(this.localize(
                    "Landing_BadType",
                    "The Date field must be a date column or a date hierarchy level"
                ));
                this.host.eventService?.renderingFinished(options);
                return;
            }

            this.filterTarget = targetFromQueryName(category.source.queryName || "");
            this.parseData(dataView, category);
            this.restoreVisibleMonth(dataView);
            this.restoreActivePreset(dataView);
            this.restoreSelectionFromFilters(options.jsonFilters);

            if (!this.focusedDate) {
                this.focusedDate = this.defaultFocusDate();
            }

            this.renderCalendar();
            this.host.eventService?.renderingFinished(options);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.host.eventService?.renderingFailed(options, message);
            throw error;
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ---- data parsing ---------------------------------------------------

    private isDateColumn(category: DataViewCategoryColumn): boolean {
        const type = category.source.type as
            | { dateTime?: boolean; numeric?: boolean; integer?: boolean }
            | undefined;
        if (type?.dateTime) {
            return true;
        }
        if (type?.numeric || type?.integer) {
            return true;
        }
        // Fall back to inspecting the first non-null value.
        const first = category.values?.find((v) => v !== null && v !== undefined);
        return first instanceof Date;
    }

    private parseData(dataView: DataView | undefined, category: DataViewCategoryColumn): void {
        let min: Date | null = null;
        let max: Date | null = null;
        this.dataValues = new Map<string, number>();
        this.hasValues = false;

        const raws = category.values || [];
        this.dataTruncated = raws.length >= DATA_REDUCTION_COUNT;

        const values = dataView?.categorical?.values?.[0];
        const measures = values?.values;
        this.hasValues = Array.isArray(measures) && measures.length > 0;

        for (let i = 0; i < raws.length; i++) {
            const date = this.coerceDate(raws[i]);
            if (!date) {
                continue;
            }
            if (!min || date < min) {
                min = date;
            }
            if (!max || date > max) {
                max = date;
            }
            if (this.hasValues && measures) {
                const raw = measures[i];
                const num = typeof raw === "number" ? raw : Number(raw);
                if (!isNaN(num)) {
                    const key = this.dayKey(date);
                    this.dataValues.set(key, (this.dataValues.get(key) || 0) + num);
                }
            }
        }
        this.dataMin = min;
        this.dataMax = max;

        this.valueMin = 0;
        this.valueMax = 0;
        let first = true;
        for (const v of this.dataValues.values()) {
            if (first) {
                this.valueMin = v;
                this.valueMax = v;
                first = false;
            } else {
                if (v < this.valueMin) {
                    this.valueMin = v;
                }
                if (v > this.valueMax) {
                    this.valueMax = v;
                }
            }
        }
    }

    private coerceDate(raw: powerbi.PrimitiveValue): Date | null {
        if (raw instanceof Date) {
            return isNaN(raw.getTime()) ? null : startOfDay(raw);
        }
        if (typeof raw === "number") {
            // A bare year from a date hierarchy's Year level.
            if (raw >= 1000 && raw <= 9999) {
                return makeDate(raw, 0, 1);
            }
            const fromEpoch = new Date(raw);
            return isNaN(fromEpoch.getTime()) ? null : startOfDay(fromEpoch);
        }
        if (typeof raw === "string") {
            const parsed = parseDateWithoutTimezone(raw);
            return isNaN(parsed.getTime()) ? null : startOfDay(parsed);
        }
        return null;
    }

    // ---- persisted view state + bookmark restore ------------------------

    private restoreVisibleMonth(dataView: DataView | undefined): void {
        if (this.visible) {
            return; // keep the month the user navigated to within this instance
        }
        const general = dataView?.metadata?.objects?.general as
            | { visibleYear?: number; visibleMonth?: number }
            | undefined;
        if (
            general &&
            typeof general.visibleYear === "number" &&
            typeof general.visibleMonth === "number"
        ) {
            this.visible = { year: general.visibleYear, month: general.visibleMonth };
            return;
        }
        const anchor = this.dataMax || new Date();
        this.visible = { year: anchor.getFullYear(), month: anchor.getMonth() };
    }

    private restoreActivePreset(dataView: DataView | undefined): void {
        if (this.activePreset !== null) {
            return;
        }
        const general = dataView?.metadata?.objects?.general as
            | { activePreset?: string }
            | undefined;
        if (general && typeof general.activePreset === "string" && general.activePreset) {
            this.activePreset = general.activePreset;
        }
    }

    /**
     * Restore selection from inbound bookmark/report state.
     *
     * NOTE: `pbiviz package` emits a "Bookmarks" warning for this visual. It is
     * a KNOWN FALSE POSITIVE: the packager's heuristic only detects the
     * SelectionManager bookmark path (`applySelectionFromFilter` /
     * `registerOnSelectCallback`). We are a FILTER visual, and Microsoft
     * documents a second valid path — filter visuals restore from
     * `options.jsonFilters` (here), combined with the `general.filter` object
     * and `filterState: true` on visibleYear/visibleMonth/activePreset. Do NOT
     * add `registerOnSelectCallback` just to silence the warning; that would
     * wire up an unused selection path we don't use.
     */
    private restoreSelectionFromFilters(jsonFilters: powerbi.IFilter[] | undefined): void {
        if (!jsonFilters || jsonFilters.length === 0) {
            return;
        }
        for (const filter of jsonFilters) {
            const restored = this.selectionFromFilter(filter);
            if (restored) {
                this.selection = restored;
                return;
            }
        }
    }

    private selectionFromFilter(filter: powerbi.IFilter): Selection | null {
        const advanced = filter as unknown as {
            conditions?: Array<{ operator?: string; value?: string | number }>;
        };
        if (advanced.conditions && advanced.conditions.length >= 2) {
            const start = this.filterValueToDate(advanced.conditions[0].value);
            const endExclusive = this.filterValueToDate(advanced.conditions[1].value);
            if (start && endExclusive) {
                return { type: "range", start, end: addDays(endExclusive, -1) };
            }
        }
        const basic = filter as unknown as { values?: Array<string | number> };
        if (basic.values && basic.values.length > 0) {
            const days = basic.values
                .map((v) => this.filterValueToDate(v))
                .filter((d): d is Date => d !== null);
            if (days.length > 0) {
                return { type: "days", days };
            }
        }
        return null;
    }

    private filterValueToDate(value: string | number | undefined): Date | null {
        if (value === undefined || value === null) {
            return null;
        }
        const parsed = typeof value === "string"
            ? parseDateWithoutTimezone(value)
            : new Date(value);
        return isNaN(parsed.getTime()) ? null : startOfDay(parsed);
    }

    private persistVisibleMonth(): void {
        if (!this.visible) {
            return;
        }
        const instances: VisualObjectInstancesToPersist = {
            merge: [{
                objectName: "general",
                selector: null,
                properties: {
                    visibleYear: this.visible.year,
                    visibleMonth: this.visible.month,
                    activePreset: this.activePreset || ""
                }
            }]
        };
        this.host.persistProperties?.(instances);
    }

    private defaultFocusDate(): Date {
        if (this.selection.type === "range") {
            return this.selection.start;
        }
        if (this.selection.type === "days") {
            return this.selection.days[0];
        }
        const today = startOfDay(new Date());
        if (this.visible &&
            today.getFullYear() === this.visible.year &&
            today.getMonth() === this.visible.month) {
            return today;
        }
        return this.visible ? makeDate(this.visible.year, this.visible.month, 1) : today;
    }

    // ---- filter application ---------------------------------------------

    private applySelection(): void {
        if (!this.filterTarget) {
            return;
        }
        if (this.selection.type === "none") {
            this.host.applyJsonFilter(null, "general", "filter", REMOVE_FILTER_ACTION);
            return;
        }
        if (this.selection.type === "range") {
            const filter = buildRangeFilter(
                this.selection.start,
                addDays(this.selection.end, 1),
                this.filterTarget
            );
            this.host.applyJsonFilter(filter, "general", "filter", MERGE_FILTER_ACTION);
            return;
        }
        // days
        if (this.selection.days.length === 1) {
            const filter = buildDayFilter(this.selection.days[0], this.filterTarget);
            this.host.applyJsonFilter(filter, "general", "filter", MERGE_FILTER_ACTION);
            return;
        }
        const multi = buildMultiDayFilter(this.selection.days, this.filterTarget);
        this.host.applyJsonFilter(multi, "general", "filter", MERGE_FILTER_ACTION);
    }

    private clearSelection(): void {
        this.selection = { type: "none" };
        this.activePreset = null;
        this.dragAnchor = null;
        this.applySelection();
        this.persistVisibleMonth();
        this.renderCalendar();
    }

    private applyPreset(key: string): void {
        if (!this.filterTarget) {
            return;
        }
        const preset = presetByKey(key);
        if (!preset) {
            return;
        }
        const ctx: PresetContext = {
            now: startOfDay(new Date()),
            fiscalStartMonth: this.fiscalStartMonth(),
            weekStart: this.weekStart(),
            target: this.filterTarget
        };
        const result = preset.compute(ctx);
        this.activePreset = key;
        this.selection = {
            type: "range",
            start: result.start,
            end: addDays(result.endExclusive, -1)
        };
        this.dragAnchor = result.start;
        this.focusedDate = result.start;
        this.moveVisibleTo(result.start);
        this.host.applyJsonFilter(result.filter, "general", "filter", MERGE_FILTER_ACTION);
        this.persistVisibleMonth();
        this.renderCalendar();
    }

    // ---- interaction ----------------------------------------------------

    private onDayPointerDown(date: Date, event: PointerEvent): void {
        if (!this.interactive) {
            return;
        }
        event.preventDefault();
        this.activePreset = null;
        this.focusedDate = date;
        this.ensureVisible(date);

        const multiSelectEnabled = this.formattingSettings.interactionCard.multiSelect.value;

        if ((event.ctrlKey || event.metaKey) && multiSelectEnabled) {
            this.toggleDay(date);
            this.dragAnchor = null;
            this.applySelection();
            this.renderCalendar();
            return;
        }

        if (event.shiftKey && this.dragAnchor) {
            this.selection = this.rangeBetween(this.dragAnchor, date);
            this.applySelection();
            this.renderCalendar();
            return;
        }

        // Begin a drag range anchored on this day (a plain click is a zero-length drag).
        this.dragAnchor = date;
        this.isDragging = true;
        this.selection = { type: "range", start: date, end: date };
        this.renderCalendar();
    }

    private onDayPointerEnter(date: Date): void {
        if (!this.isDragging || !this.dragAnchor) {
            return;
        }
        this.selection = this.rangeBetween(this.dragAnchor, date);
        this.focusedDate = date;
        this.renderCalendar();
    }

    private endDrag(): void {
        if (!this.isDragging) {
            return;
        }
        this.isDragging = false;
        this.applySelection();
    }

    private toggleDay(date: Date): void {
        const days: Date[] = this.selection.type === "days"
            ? [...this.selection.days]
            : this.selection.type === "range"
                ? this.rangeDays(this.selection.start, this.selection.end)
                : [];
        const idx = days.findIndex((d) => isSameDay(d, date));
        if (idx >= 0) {
            days.splice(idx, 1);
        } else {
            days.push(date);
        }
        this.selection = days.length === 0 ? { type: "none" } : { type: "days", days };
    }

    private rangeBetween(a: Date, b: Date): Selection {
        return a <= b
            ? { type: "range", start: a, end: b }
            : { type: "range", start: b, end: a };
    }

    private rangeDays(start: Date, end: Date): Date[] {
        const days: Date[] = [];
        for (let d = start; d <= end; d = addDays(d, 1)) {
            days.push(d);
        }
        return days;
    }

    private moveVisibleTo(date: Date): void {
        this.visible = { year: date.getFullYear(), month: date.getMonth() };
    }

    /**
     * Re-anchor the visible range only when `date` falls outside the currently
     * shown months, so keyboard/mouse focus can move within a multi-month view
     * without scrolling the whole range on every step.
     */
    private ensureVisible(date: Date): void {
        if (!this.visible) {
            return;
        }
        const start = makeDate(this.visible.year, this.visible.month, 1);
        const end = addMonths(start, this.monthsToShow());
        const d = startOfDay(date);
        if (d >= start && d < end) {
            return;
        }
        this.moveVisibleTo(date);
    }

    private navigateMonths(delta: number): void {
        if (!this.visible) {
            return;
        }
        const anchor = addMonths(makeDate(this.visible.year, this.visible.month, 1), delta);
        this.visible = { year: anchor.getFullYear(), month: anchor.getMonth() };
        if (this.focusedDate) {
            const day = Math.min(this.focusedDate.getDate(), 28);
            this.focusedDate = makeDate(this.visible.year, this.visible.month, day);
        }
        this.persistVisibleMonth();
        this.renderCalendar();
    }

    private goToToday(): void {
        const today = startOfDay(new Date());
        this.moveVisibleTo(today);
        this.focusedDate = today;
        this.persistVisibleMonth();
        this.renderCalendar();
    }

    private selectFocused(event: KeyboardEvent): void {
        if (!this.focusedDate) {
            return;
        }
        this.activePreset = null;
        const multiSelectEnabled = this.formattingSettings.interactionCard.multiSelect.value;
        if ((event.ctrlKey || event.metaKey) && multiSelectEnabled) {
            this.toggleDay(this.focusedDate);
        } else if (event.shiftKey && this.dragAnchor) {
            this.selection = this.rangeBetween(this.dragAnchor, this.focusedDate);
        } else {
            this.dragAnchor = this.focusedDate;
            this.selection = { type: "range", start: this.focusedDate, end: this.focusedDate };
        }
        this.applySelection();
        this.renderCalendar();
    }

    private onGridKeyDown(event: KeyboardEvent): void {
        if (!this.interactive || !this.focusedDate) {
            return;
        }
        const weekStart = this.weekStart();
        let next: Date | null = null;
        switch (event.key) {
            case "ArrowLeft": next = addDays(this.focusedDate, -1); break;
            case "ArrowRight": next = addDays(this.focusedDate, 1); break;
            case "ArrowUp": next = addDays(this.focusedDate, -7); break;
            case "ArrowDown": next = addDays(this.focusedDate, 7); break;
            case "Home": next = startOfWeek(this.focusedDate, weekStart); break;
            case "End": next = addDays(startOfWeek(this.focusedDate, weekStart), 6); break;
            case "PageUp": this.navigateMonths(-1); event.preventDefault(); return;
            case "PageDown": this.navigateMonths(1); event.preventDefault(); return;
            case "Enter":
            case " ":
            case "Spacebar":
                event.preventDefault();
                this.selectFocused(event);
                return;
            case "Escape":
            case "Delete":
            case "Backspace":
                event.preventDefault();
                this.clearSelection();
                return;
            default:
                return;
        }
        if (next) {
            event.preventDefault();
            this.focusedDate = next;
            this.ensureVisible(next);
            this.renderCalendar();
            this.focusActiveCell();
        }
    }

    // ---- rendering ------------------------------------------------------

    private clear(): void {
        while (this.root.firstChild) {
            this.root.removeChild(this.root.firstChild);
        }
    }

    private renderLanding(message: string): void {
        this.clear();
        const landing = document.createElement("div");
        landing.className = "cs-landing";
        const title = document.createElement("div");
        title.className = "cs-landing-title";
        title.textContent = this.localize("Visual_Title", "Atlyn Calendar Slicer");
        landing.appendChild(title);
        const body = document.createElement("div");
        body.textContent = message;
        landing.appendChild(body);
        this.root.appendChild(landing);
    }

    private renderCalendar(): void {
        this.clear();
        if (!this.visible) {
            return;
        }
        this.root.appendChild(this.buildToolbar());
        if (this.formattingSettings.presetsCard.show.value) {
            this.root.appendChild(this.buildPresets());
        }

        const months = this.monthsToShow();
        if (months <= 1) {
            this.root.appendChild(this.buildGrid(this.visible.year, this.visible.month, false));
            return;
        }

        const container = document.createElement("div");
        container.className = "cs-months";
        const first = makeDate(this.visible.year, this.visible.month, 1);
        for (let i = 0; i < months; i++) {
            const m = addMonths(first, i);
            container.appendChild(this.buildGrid(m.getFullYear(), m.getMonth(), true));
        }
        this.root.appendChild(container);
    }

    private buildPresets(): HTMLElement {
        const bar = document.createElement("div");
        bar.className = "cs-presets";
        bar.setAttribute("role", "group");
        bar.setAttribute("aria-label", this.localize("Aria_Presets", "Relative date presets"));
        for (const preset of PRESETS) {
            const label = this.localize(preset.labelKey, preset.label);
            const btn = this.button(label, label, () => this.applyPreset(preset.key));
            if (this.activePreset === preset.key) {
                btn.classList.add("active");
                btn.setAttribute("aria-pressed", "true");
            } else {
                btn.setAttribute("aria-pressed", "false");
            }
            bar.appendChild(btn);
        }
        return bar;
    }

    private buildToolbar(): HTMLElement {
        const bar = document.createElement("div");
        bar.className = "cs-toolbar";

        const nav = document.createElement("div");
        nav.className = "cs-nav";
        nav.appendChild(this.button(
            "\u2039",
            this.localize("Nav_PrevMonth", "Previous month"),
            () => this.navigateMonths(-1)
        ));
        nav.appendChild(this.button(
            "\u203A",
            this.localize("Nav_NextMonth", "Next month"),
            () => this.navigateMonths(1)
        ));
        bar.appendChild(nav);

        const title = document.createElement("div");
        title.className = "cs-title";
        title.textContent = this.rangeTitle();
        bar.appendChild(title);

        const actions = document.createElement("div");
        actions.className = "cs-nav";
        actions.appendChild(this.button(
            this.localize("Nav_Today", "Today"),
            this.localize("Nav_Today", "Today"),
            () => this.goToToday()
        ));
        actions.appendChild(this.button(
            this.localize("Nav_Clear", "Clear"),
            this.localize("Nav_Clear", "Clear"),
            () => this.clearSelection()
        ));
        bar.appendChild(actions);
        return bar;
    }

    private button(label: string, ariaLabel: string, onClick: () => void): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "cs-btn";
        btn.textContent = label;
        btn.setAttribute("aria-label", ariaLabel);
        if (this.interactive) {
            btn.addEventListener("click", onClick);
        } else {
            btn.disabled = true;
        }
        return btn;
    }

    private buildGrid(year: number, month: number, withCaption: boolean): HTMLElement {
        const weekStart = this.weekStart();
        const cells = this.formattingSettings.cellsCard;
        const showWeekNumbers = this.formattingSettings.calendarCard.showWeekNumbers.value;

        const table = document.createElement("table");
        table.className = "cs-grid";
        table.setAttribute("role", "grid");
        table.setAttribute("aria-label", this.monthTitle(year, month));
        if (this.formattingSettings.interactionCard.multiSelect.value) {
            table.setAttribute("aria-multiselectable", "true");
        }

        if (withCaption) {
            const caption = document.createElement("caption");
            caption.textContent = this.monthTitle(year, month);
            table.appendChild(caption);
        }

        const thead = document.createElement("thead");
        const headRow = document.createElement("tr");
        headRow.setAttribute("role", "row");
        if (showWeekNumbers) {
            const wk = document.createElement("th");
            wk.setAttribute("role", "columnheader");
            wk.setAttribute("scope", "col");
            wk.className = "cs-week-number";
            wk.textContent = this.localize("Aria_WeekNumber", "Wk");
            headRow.appendChild(wk);
        }
        for (const label of this.weekdayLabels(weekStart)) {
            const th = document.createElement("th");
            th.setAttribute("role", "columnheader");
            th.setAttribute("scope", "col");
            th.textContent = label;
            if (!this.isHighContrast) {
                th.style.color = cells.headerColor.value.value;
            }
            headRow.appendChild(th);
        }
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        const grid = buildMonthGrid(year, month, weekStart);
        for (const week of grid) {
            const tr = document.createElement("tr");
            tr.setAttribute("role", "row");
            if (showWeekNumbers) {
                const wk = document.createElement("td");
                wk.className = "cs-week-number";
                wk.setAttribute("role", "rowheader");
                // Use a mid-week day so the ISO week is unambiguous regardless
                // of the configured display week start.
                wk.textContent = String(getISOWeek(week[3].date));
                tr.appendChild(wk);
            }
            for (const cell of week) {
                tr.appendChild(this.buildDayCell(cell, month));
            }
            tbody.appendChild(tr);
        }
        table.appendChild(tbody);
        table.addEventListener("keydown", (e) => this.onGridKeyDown(e));
        table.addEventListener("contextmenu", (e) => this.onContextMenu(e));
        return table;
    }

    private buildDayCell(cell: CalendarDay, displayMonth: number): HTMLElement {
        const cells = this.formattingSettings.cellsCard;
        const calendar = this.formattingSettings.calendarCard;
        const td = document.createElement("td");

        const day = document.createElement("div");
        day.className = "cs-day";
        day.textContent = String(cell.date.getDate());
        day.setAttribute("role", "gridcell");
        day.setAttribute("aria-label", this.dayLabel(cell.date));
        day.dataset.key = this.dayKey(cell.date);

        const inMonth = cell.date.getMonth() === displayMonth;
        if (!inMonth) {
            day.classList.add("other-month");
        }
        if (!this.isHighContrast) {
            day.style.color = cells.textColor.value.value;
        }

        const isWeekend = cell.date.getDay() === 0 || cell.date.getDay() === 6;
        if (cells.weekendShading.value && isWeekend) {
            day.classList.add("weekend");
            if (!this.isHighContrast) {
                day.style.background = cells.weekendColor.value.value;
            }
        }

        const heatmap = this.formattingSettings.heatmapCard;
        const dayKey = this.dayKey(cell.date);
        const hasData = this.dataValues.has(dayKey);

        if (heatmap.show.value && this.hasValues && hasData && !this.isHighContrast) {
            const value = this.dataValues.get(dayKey) as number;
            day.style.background = this.heatColor(
                value,
                heatmap.minColor.value.value,
                heatmap.maxColor.value.value
            );
        }

        const noData = heatmap.datesWithDataOnly.value && this.hasValues && !hasData && !this.dataTruncated;
        if (noData) {
            day.classList.add("no-data");
            day.setAttribute("aria-disabled", "true");
        }

        if (calendar.showTodayMarker.value && isSameDay(cell.date, startOfDay(new Date()))) {
            day.classList.add("today");
            if (!this.isHighContrast) {
                day.style.boxShadow = `inset 0 0 0 1px ${cells.todayColor.value.value}`;
            }
        }

        const selected = this.isSelected(cell.date);
        day.setAttribute("aria-selected", selected ? "true" : "false");
        if (selected) {
            day.classList.add("selected");
            if (!this.isHighContrast) {
                day.style.background = cells.selectedColor.value.value;
                day.style.color = "#ffffff";
            }
        }

        const focused = this.focusedDate !== null && isSameDay(cell.date, this.focusedDate);
        day.tabIndex = this.interactive && focused && inMonth ? 0 : -1;

        if (this.interactive && !noData) {
            day.addEventListener("pointerdown", (e) => this.onDayPointerDown(cell.date, e));
            day.addEventListener("pointerenter", () => this.onDayPointerEnter(cell.date));
        }

        td.appendChild(day);
        return td;
    }

    private focusActiveCell(): void {
        if (!this.focusedDate) {
            return;
        }
        const key = this.dayKey(this.focusedDate);
        const el = this.root.querySelector<HTMLElement>(`.cs-day[data-key="${key}"]:not(.other-month)`)
            ?? this.root.querySelector<HTMLElement>(`.cs-day[data-key="${key}"]`);
        el?.focus();
    }

    private onContextMenu(event: MouseEvent): void {
        event.preventDefault();
        this.selectionManager.showContextMenu(
            {},
            { x: event.clientX, y: event.clientY }
        );
    }

    // ---- selection helpers ----------------------------------------------

    private isSelected(date: Date): boolean {
        if (this.selection.type === "range") {
            return date >= this.selection.start && date <= this.selection.end;
        }
        if (this.selection.type === "days") {
            return this.selection.days.some((d) => isSameDay(d, date));
        }
        return false;
    }

    // ---- formatting-derived values --------------------------------------

    private weekStart(): WeekStart {
        const value = Number(this.formattingSettings.calendarCard.weekStartDay.value.value);
        return (value === 1 || value === 6 ? value : 0) as WeekStart;
    }

    private fiscalStartMonth(): number {
        const value = Number(this.formattingSettings.calendarCard.fiscalYearStartMonth.value.value);
        return value >= 1 && value <= 12 ? value : 1;
    }

    private monthsToShow(): number {
        const value = Number(this.formattingSettings.calendarCard.monthsToShow.value);
        if (isNaN(value)) {
            return 1;
        }
        return Math.min(4, Math.max(1, Math.round(value)));
    }

    private rangeTitle(): string {
        if (!this.visible) {
            return "";
        }
        const months = this.monthsToShow();
        const first = this.monthTitle(this.visible.year, this.visible.month);
        if (months <= 1) {
            return first;
        }
        const lastDate = addMonths(makeDate(this.visible.year, this.visible.month, 1), months - 1);
        const last = this.monthTitle(lastDate.getFullYear(), lastDate.getMonth());
        return `${first} \u2013 ${last}`;
    }

    /**
     * Linear interpolation between two hex colours by the value's position in
     * the observed [min, max] range. Returns an "rgb(r, g, b)" string.
     */
    private heatColor(value: number, minHex: string, maxHex: string): string {
        const span = this.valueMax - this.valueMin;
        const t = span > 0 ? (value - this.valueMin) / span : 1;
        const from = this.hexToRgb(minHex);
        const to = this.hexToRgb(maxHex);
        const r = Math.round(from.r + (to.r - from.r) * t);
        const g = Math.round(from.g + (to.g - from.g) * t);
        const b = Math.round(from.b + (to.b - from.b) * t);
        return `rgb(${r}, ${g}, ${b})`;
    }

    private hexToRgb(hex: string): { r: number; g: number; b: number } {
        let value = hex.replace("#", "");
        if (value.length === 3) {
            value = value.split("").map((c) => c + c).join("");
        }
        const int = parseInt(value, 16);
        if (isNaN(int)) {
            return { r: 0, g: 0, b: 0 };
        }
        return {
            r: (int >> 16) & 0xff,
            g: (int >> 8) & 0xff,
            b: int & 0xff
        };
    }

    private weekdayLabels(weekStart: WeekStart): string[] {
        const formatter = new Intl.DateTimeFormat(this.locale, { weekday: "short" });
        const labels: string[] = [];
        // 2024-01-07 is a Sunday; rotate from the configured week start.
        const base = makeDate(2024, 0, 7);
        for (let i = 0; i < 7; i++) {
            labels.push(formatter.format(addDays(base, weekStart + i)));
        }
        return labels;
    }

    private monthTitle(year: number, month: number): string {
        return new Intl.DateTimeFormat(this.locale, { month: "long", year: "numeric" })
            .format(makeDate(year, month, 1));
    }

    private dayLabel(date: Date): string {
        return new Intl.DateTimeFormat(this.locale, {
            weekday: "long",
            year: "numeric",
            month: "long",
            day: "numeric"
        }).format(date);
    }

    private dayKey(date: Date): string {
        return `${date.getFullYear()}-${date.getMonth()}-${date.getDate()}`;
    }

    private localize(key: string, fallback: string): string {
        const resolved = this.localizationManager?.getDisplayName(key);
        return resolved && resolved !== key ? resolved : fallback;
    }
}
