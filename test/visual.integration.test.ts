/**
 * Integration tests for the Atlyn Calendar Slicer visual.
 *
 * Drives the visual through a mock IVisualHost in a happy-dom document, covering
 * the landing page, grid rendering, single-day selection -> AdvancedFilter,
 * clearing, and bookmark restore from an inbound jsonFilter.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import powerbi from "powerbi-visuals-api";
import { Visual } from "../src/visual";
import { buildEmptyDataView, buildMockDataView } from "./helpers/mockDataView";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

interface AppliedFilter {
    filter: unknown;
    objectName: string;
    propertyName: string;
    action: number;
}

function createHost() {
    const applied: AppliedFilter[] = [];
    const persisted: unknown[] = [];
    const host = {
        locale: "en-US",
        colorPalette: {
            isHighContrast: false,
            foreground: { value: "#ffffff" },
            background: { value: "#000000" }
        },
        eventService: {
            renderingStarted: vi.fn(),
            renderingFinished: vi.fn(),
            renderingFailed: vi.fn()
        },
        createSelectionManager: () => ({
            showContextMenu: vi.fn(),
            registerOnSelectCallback: vi.fn()
        }),
        createLocalizationManager: () => ({
            getDisplayName: (key: string) => key
        }),
        applyJsonFilter: (
            filter: unknown,
            objectName: string,
            propertyName: string,
            action: number
        ) => {
            applied.push({ filter, objectName, propertyName, action });
        },
        persistProperties: (instances: unknown) => {
            persisted.push(instances);
        }
    };
    return { host, applied, persisted };
}

function createVisual() {
    const element = document.createElement("div");
    document.body.appendChild(element);
    const { host, applied, persisted } = createHost();
    const visual = new Visual({ element, host } as unknown as VisualConstructorOptions);
    return { visual, element, applied, persisted, host };
}

function updateOptions(dataView: unknown, jsonFilters: unknown[] = []) {
    return {
        dataViews: dataView ? [dataView] : [],
        jsonFilters,
        type: 2,
        viewport: { width: 400, height: 400 }
    } as unknown as VisualUpdateOptions;
}

describe("Atlyn Calendar Slicer visual", () => {
    beforeEach(() => {
        document.body.replaceChildren();
    });

    it("shows a landing page when no date field is bound", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildEmptyDataView()));
        expect(element.querySelector(".cs-landing")).not.toBeNull();
        expect(element.querySelector(".cs-grid")).toBeNull();
    });

    it("renders a 7-column month grid for a bound date column", () => {
        const { visual, element } = createVisual();
        const dates = [new Date(2024, 2, 1), new Date(2024, 2, 15), new Date(2024, 2, 31)];
        visual.update(updateOptions(buildMockDataView({ dates })));

        const grid = element.querySelector(".cs-grid");
        expect(grid).not.toBeNull();
        const headers = grid!.querySelectorAll("th[role='columnheader']");
        expect(headers.length).toBe(7);
        const cells = grid!.querySelectorAll(".cs-day[role='gridcell']");
        // 5 or 6 weeks * 7 days.
        expect(cells.length % 7).toBe(0);
        expect(cells.length).toBeGreaterThanOrEqual(28);
    });

    it("applies a half-open AdvancedFilter on single-day click", () => {
        const { visual, element, applied } = createVisual();
        const dates = [new Date(2024, 2, 1), new Date(2024, 2, 31)];
        visual.update(updateOptions(buildMockDataView({ dates })));

        const target = element.querySelector<HTMLElement>(
            ".cs-day[data-key='2024-2-15']"
        );
        expect(target).not.toBeNull();
        target!.dispatchEvent(new Event("pointerdown", { bubbles: true }));
        // Release the drag to commit the filter.
        element.querySelector(".atlynCalendarSlicer");
        const root = element.firstElementChild as HTMLElement;
        root.dispatchEvent(new Event("pointerup", { bubbles: true }));

        const merges = applied.filter((a) => a.action === 0 /* merge */ && a.filter);
        expect(merges.length).toBeGreaterThanOrEqual(1);
        const filter = merges[merges.length - 1].filter as {
            conditions: Array<{ operator: string; value: string }>;
        };
        expect(filter.conditions[0].operator).toBe("GreaterThanOrEqual");
        expect(filter.conditions[1].operator).toBe("LessThan");
    });

    it("removes the filter when Clear is pressed", () => {
        const { visual, element, applied } = createVisual();
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 1)] })));

        const buttons = Array.from(element.querySelectorAll<HTMLButtonElement>(".cs-btn"));
        const clearBtn = buttons.find((b) => b.getAttribute("aria-label") === "Clear");
        expect(clearBtn).toBeDefined();
        clearBtn!.dispatchEvent(new Event("click", { bubbles: true }));

        const removes = applied.filter((a) => a.action === 1 /* remove */);
        expect(removes.length).toBeGreaterThanOrEqual(1);
    });

    it("restores a range selection from an inbound AdvancedFilter (bookmark)", () => {
        const { visual, element } = createVisual();
        const dates = [new Date(2024, 2, 1), new Date(2024, 2, 31)];
        const jsonFilter = {
            target: { table: "Calendar", column: "Date" },
            logicalOperator: "And",
            conditions: [
                { operator: "GreaterThanOrEqual", value: "2024-03-10T00:00:00.000Z" },
                { operator: "LessThan", value: "2024-03-16T00:00:00.000Z" }
            ]
        };
        visual.update(updateOptions(buildMockDataView({ dates }), [jsonFilter]));

        const selected = element.querySelectorAll(".cs-day.selected");
        // 2024-03-10 .. 2024-03-15 inclusive = 6 days.
        expect(selected.length).toBe(6);
    });

    it("navigates to the next month when the next button is clicked", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 15)] })));

        const titleBefore = element.querySelector(".cs-title")?.textContent;
        const nextBtn = Array.from(
            element.querySelectorAll<HTMLButtonElement>(".cs-btn")
        ).find((b) => b.getAttribute("aria-label") === "Next month");
        nextBtn!.dispatchEvent(new Event("click", { bubbles: true }));
        const titleAfter = element.querySelector(".cs-title")?.textContent;
        expect(titleAfter).not.toBe(titleBefore);
    });

    it("renders preset buttons and applies a filter when one is clicked", () => {
        const { visual, element, applied } = createVisual();
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 15)] })));

        const presets = element.querySelector(".cs-presets");
        expect(presets).not.toBeNull();
        const last7 = Array.from(
            presets!.querySelectorAll<HTMLButtonElement>(".cs-btn")
        ).find((b) => b.textContent === "Last 7 Days");
        expect(last7).toBeDefined();

        last7!.dispatchEvent(new Event("click", { bubbles: true }));
        const merges = applied.filter((a) => a.action === 0 && a.filter);
        expect(merges.length).toBeGreaterThanOrEqual(1);
        // Relative filter round-trips as InLast (0) over 7 Days (0).
        const filter = merges[merges.length - 1].filter as {
            operator: number; timeUnitsCount: number;
        };
        expect(filter.operator).toBe(0);
        expect(filter.timeUnitsCount).toBe(7);
        const activeAfter = Array.from(
            element.querySelectorAll<HTMLButtonElement>(".cs-presets .cs-btn")
        ).find((b) => b.textContent === "Last 7 Days");
        expect(activeAfter!.classList.contains("active")).toBe(true);
    });

    it("heat-shades day cells when a measure is bound and the heatmap is enabled", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)],
            values: [10, 100],
            objects: { heatmap: { show: true } }
        })));

        const low = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-1']");
        const high = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-31']");
        expect(low!.style.background).toContain("rgb");
        expect(high!.style.background).toContain("rgb");
        expect(low!.style.background).not.toBe(high!.style.background);
    });

    it("greys days without data when 'dates with data only' is enabled", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 15)],
            values: [10, 20],
            objects: { heatmap: { datesWithDataOnly: true } }
        })));

        const withData = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']");
        const withoutData = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']");
        expect(withData!.classList.contains("no-data")).toBe(false);
        expect(withoutData!.classList.contains("no-data")).toBe(true);
        expect(withoutData!.getAttribute("aria-disabled")).toBe("true");
    });

    it("does not grey empty days when the category hit the data-reduction cap", () => {
        const { visual, element } = createVisual();
        // Exactly the capabilities `top` count => the table may be truncated,
        // so we cannot trust that a missing day is genuinely empty.
        const CAP = 30000;
        const dates: Date[] = new Array(CAP).fill(0).map(() => new Date(2024, 2, 15));
        const values: number[] = new Array(CAP).fill(1);
        visual.update(updateOptions(buildMockDataView({
            dates,
            values,
            objects: { heatmap: { datesWithDataOnly: true } }
        })));

        // March 10 has no data, but because the dataView is truncated it must
        // NOT be mislabelled as empty.
        const day10 = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']");
        expect(day10).not.toBeNull();
        expect(day10!.classList.contains("no-data")).toBe(false);
        expect(day10!.getAttribute("aria-disabled")).toBeNull();
    });

    it("greys empty days when the category is comfortably below the cap", () => {
        const { visual, element } = createVisual();
        // A full month of daily rows is well under the 30000 cap, so the
        // received dates are complete and greying is trustworthy.
        const dates: Date[] = [];
        const values: number[] = [];
        for (let d = 1; d <= 28; d++) {
            if (d === 10) {
                continue;
            }
            dates.push(new Date(2024, 2, d));
            values.push(d);
        }
        visual.update(updateOptions(buildMockDataView({
            dates,
            values,
            objects: { heatmap: { datesWithDataOnly: true } }
        })));

        const day10 = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']");
        expect(day10!.classList.contains("no-data")).toBe(true);
        expect(day10!.getAttribute("aria-disabled")).toBe("true");
    });

    it("reacts to a high-contrast toggle between updates (no stale theme state)", () => {
        const { visual, element, host } = createVisual();
        const dates = [new Date(2024, 2, 15)];

        // First update: normal theme -> a selected cell gets an inline colour.
        visual.update(updateOptions(buildMockDataView({ dates })));
        expect(element.querySelector(".atlynCalendarSlicer")?.classList.contains("high-contrast"))
            .toBe(false);

        // Author flips Windows high contrast while the report is open.
        host.colorPalette.isHighContrast = true;
        visual.update(updateOptions(buildMockDataView({ dates })));

        const root = element.querySelector<HTMLElement>(".atlynCalendarSlicer")!;
        expect(root.classList.contains("high-contrast")).toBe(true);
        // High-contrast suppresses inline header colours (theme-driven instead).
        const header = element.querySelector<HTMLElement>(".cs-grid thead th:not(.cs-week-number)")!;
        expect(header.style.color).toBe("");

        // And back again — the value must not be stuck on.
        host.colorPalette.isHighContrast = false;
        visual.update(updateOptions(buildMockDataView({ dates })));
        expect(element.querySelector(".atlynCalendarSlicer")?.classList.contains("high-contrast"))
            .toBe(false);
    });

    it("renders read-only and applies no filter when host interactions are disabled", () => {
        const { visual, element, host, applied } = createVisual();
        host.hostCapabilities = { allowInteractions: false };
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 15)] })));

        const root = element.querySelector<HTMLElement>(".atlynCalendarSlicer")!;
        expect(root.classList.contains("read-only")).toBe(true);

        // No cell is tabbable in read-only mode.
        expect(element.querySelectorAll(".cs-day[tabindex='0']").length).toBe(0);
        // Toolbar/preset buttons are disabled.
        const buttons = element.querySelectorAll<HTMLButtonElement>("button.cs-btn");
        expect(buttons.length).toBeGreaterThan(0);
        buttons.forEach((b) => expect(b.disabled).toBe(true));

        // A click must not produce a filter.
        const cell = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']")!;
        cell.dispatchEvent(new Event("pointerdown", { bubbles: true }));
        const root2 = element.firstElementChild as HTMLElement;
        root2.dispatchEvent(new Event("pointerup", { bubbles: true }));
        expect(applied.length).toBe(0);
    });

    it("is interactive by default when hostCapabilities is absent", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 15)] })));
        const root = element.querySelector<HTMLElement>(".atlynCalendarSlicer")!;
        expect(root.classList.contains("read-only")).toBe(false);
        expect(element.querySelectorAll(".cs-day[tabindex='0']").length).toBe(1);
    });

    it("renders multiple month grids when monthsToShow > 1", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 15)],
            objects: { calendar: { monthsToShow: 3 } }
        })));

        const grids = element.querySelectorAll(".cs-grid");
        expect(grids.length).toBe(3);
        const container = element.querySelector(".cs-months");
        expect(container).not.toBeNull();
        // Each grid is captioned with its own month.
        const captions = element.querySelectorAll(".cs-grid caption");
        expect(captions.length).toBe(3);
        // Toolbar title spans the visible range.
        const title = element.querySelector(".cs-title");
        expect(title!.textContent).toContain("\u2013");
    });

    it("renders an ISO week-number column when showWeekNumbers is enabled", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 0, 15)],
            objects: { calendar: { showWeekNumbers: true } }
        })));

        const grid = element.querySelector(".cs-grid")!;
        const headerCells = grid.querySelectorAll("thead th");
        // 7 weekday columns + 1 week-number column.
        expect(headerCells.length).toBe(8);
        const weekHeader = grid.querySelector("thead th.cs-week-number");
        expect(weekHeader).not.toBeNull();
        const weekCells = grid.querySelectorAll("tbody td.cs-week-number");
        expect(weekCells.length).toBeGreaterThanOrEqual(5);
        // First tbody week-number is a positive ISO week number.
        const first = Number(weekCells[0].textContent);
        expect(first).toBeGreaterThan(0);
        expect(first).toBeLessThanOrEqual(53);
    });
});
