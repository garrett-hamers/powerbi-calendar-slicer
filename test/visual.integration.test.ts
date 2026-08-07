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
import { addDays, serializeDate, serializeDateNaive } from "../src/utils/dateMath";

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
    const contextMenus: unknown[][] = [];
    const createSelectionIdBuilder = () => {
        let index = -1;
        const builder = {
            withCategory: (_category: unknown, categoryIndex: number) => {
                index = categoryIndex;
                return builder;
            },
            createSelectionId: () => ({ key: `date:${index}` })
        };
        return builder;
    };
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
            showContextMenu: vi.fn((selectionId: unknown, position: unknown) => {
                contextMenus.push([selectionId, position]);
            }),
            registerOnSelectCallback: vi.fn()
        }),
        createSelectionIdBuilder,
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
    return { host, applied, persisted, contextMenus };
}

function createVisual() {
    const element = document.createElement("div");
    document.body.appendChild(element);
    const { host, applied, persisted, contextMenus } = createHost();
    const visual = new Visual({ element, host } as unknown as VisualConstructorOptions);
    return { visual, element, applied, persisted, host, contextMenus };
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

    it("rejects automatic numeric hierarchy levels", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [2024, 2025],
            dateType: "numeric",
            dateQueryName: "Calendar.Date Hierarchy.Year"
        })));
        expect(element.querySelector(".cs-landing")?.textContent).toContain("hierarchies");
        expect(element.querySelector(".cs-grid")).toBeNull();
    });

    it("rejects a DateTime hierarchy path without a literal hierarchy label", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 1)],
            dateQueryName: "Calendar.Fiscal.Date"
        })));
        expect(element.querySelector(".cs-landing")?.textContent).toContain("hierarchies");
        expect(element.querySelector(".cs-grid")).toBeNull();
    });

    it("accepts concrete DateTime identifiers containing hierarchy text", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 1)],
            dateQueryName: "HierarchyCalendar.Date"
        })));
        expect(element.querySelector(".cs-grid")).not.toBeNull();
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

    it("applies a BasicFilter of naive-local values on Ctrl+click multi-select", () => {
        const { visual, element, applied } = createVisual();
        const dates = [new Date(2024, 10, 1), new Date(2024, 10, 30)];
        visual.update(updateOptions(buildMockDataView({ dates })));

        // Ctrl+click days 9, 11, 13, 15 — the non-contiguous case that shipped
        // broken in v1.0.0.0 (BasicFilter values carried a UTC "Z" and matched
        // nothing). PointerEvent isn't constructible in happy-dom, so set the
        // modifier flag on a plain Event.
        for (const day of [9, 11, 13, 15]) {
            const cell = element.querySelector<HTMLElement>(
                `.cs-day[data-key='2024-10-${day}']`
            );
            expect(cell).not.toBeNull();
            const ev = new Event("pointerdown", { bubbles: true });
            (ev as unknown as { ctrlKey: boolean }).ctrlKey = true;
            cell!.dispatchEvent(ev);
        }

        const merges = applied.filter((a) => a.action === 0 /* merge */ && a.filter);
        expect(merges.length).toBeGreaterThanOrEqual(1);
        const filter = merges[merges.length - 1].filter as {
            operator: string;
            values: string[];
        };
        expect(filter.operator).toBe("In");
        expect(filter.values).toEqual([
            "2024-11-09T00:00:00",
            "2024-11-11T00:00:00",
            "2024-11-13T00:00:00",
            "2024-11-15T00:00:00"
        ]);
        // The v1.0.0.0 defect: any value ending in "Z" is silently unmatchable.
        for (const v of filter.values) {
            expect(v.endsWith("Z")).toBe(false);
        }
    });

    it("restores a non-contiguous selection from an inbound BasicFilter (bookmark)", () => {
        const { visual, element } = createVisual();
        const dates = [new Date(2024, 10, 1), new Date(2024, 10, 30)];
        const jsonFilter = {
            target: { table: "Calendar", column: "Date" },
            operator: "In",
            values: ["2024-11-09T00:00:00", "2024-11-13T00:00:00"]
        };
        visual.update(updateOptions(buildMockDataView({ dates }), [jsonFilter]));

        const selected = Array.from(element.querySelectorAll(".cs-day.selected"));
        expect(selected.length).toBe(2);
        const keys = selected
            .map((el) => el.getAttribute("data-key"))
            .sort();
        expect(keys).toEqual(["2024-10-13", "2024-10-9"]);
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

    it("reconciles bookmark A to B to clear in the same visual instance", () => {
        const { visual, element } = createVisual();
        const dataView = buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)]
        });
        const bookmarkA = {
            target: { table: "Calendar", column: "Date" },
            logicalOperator: "And",
            conditions: [
                { operator: "GreaterThanOrEqual", value: "2024-03-10T00:00:00.000Z" },
                { operator: "LessThan", value: "2024-03-16T00:00:00.000Z" }
            ]
        };
        const bookmarkB = {
            target: { table: "Calendar", column: "Date" },
            operator: "In",
            values: ["2024-03-20T00:00:00", "2024-03-22T00:00:00"]
        };

        visual.update(updateOptions(dataView, [bookmarkA]));
        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(6);

        visual.update(updateOptions(dataView, [bookmarkB]));
        expect(Array.from(element.querySelectorAll<HTMLElement>(".cs-day.selected"))
            .map((day) => day.dataset.key)
            .sort()).toEqual(["2024-2-20", "2024-2-22"]);

        visual.update(updateOptions(dataView, []));
        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(0);
    });

    it("preserves selection when an update omits jsonFilters", () => {
        const { visual, element } = createVisual();
        const dataView = buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)]
        });
        const filter = {
            target: { table: "Calendar", column: "Date" },
            operator: "In",
            values: ["2024-03-10T00:00:00"]
        };
        visual.update(updateOptions(dataView, [filter]));
        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(1);

        const withoutFilters = updateOptions(dataView);
        withoutFilters.jsonFilters = undefined;
        visual.update(withoutFilters);

        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(1);
    });

    it("preserves the bound model through resize updates without dataViews", () => {
        const { visual, element } = createVisual();
        const dataView = buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)]
        });
        visual.update(updateOptions(dataView));
        const before = element.querySelector(".cs-grid");
        visual.update({
            type: 4,
            viewport: { width: 640, height: 480 }
        } as unknown as VisualUpdateOptions);
        expect(before).not.toBeNull();
        expect(element.querySelector(".cs-grid")).not.toBeNull();
        expect(element.querySelector(".cs-landing")).toBeNull();
    });

    it("ignores inbound filters for a different target", () => {
        const { visual, element } = createVisual();
        const dataView = buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)]
        });
        const matching = {
            target: { table: "Calendar", column: "Date" },
            operator: "In",
            values: ["2024-03-10T00:00:00"]
        };
        const unrelated = {
            target: { table: "Sales", column: "OrderDate" },
            operator: "In",
            values: ["2024-03-20T00:00:00"]
        };

        visual.update(updateOptions(dataView, [matching]));
        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(1);
        visual.update(updateOptions(dataView, [unrelated]));
        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(0);
    });

    it("restores a RelativeDateFilter using persisted preset state", () => {
        vi.useFakeTimers();
        vi.setSystemTime(new Date(2024, 2, 15, 12));
        try {
            const { visual, element } = createVisual();
            const dataView = buildMockDataView({
                dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)],
                objects: { general: { activePreset: "last7" } }
            });
            const relative = {
                target: { table: "Calendar", column: "Date" },
                operator: 0,
                timeUnitsCount: 7,
                timeUnitType: 0,
                includeToday: true
            };

            visual.update(updateOptions(dataView, [relative]));

            expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(7);
            const active = Array.from(
                element.querySelectorAll<HTMLButtonElement>(".cs-presets .cs-btn")
            ).find((button) => button.textContent === "Last 7 Days");
            expect(active?.getAttribute("aria-pressed")).toBe("true");
        } finally {
            vi.useRealTimers();
        }
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

    it("distinguishes a null measure from a real zero", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 15)],
            values: [null, 0],
            objects: {
                heatmap: { show: true, datesWithDataOnly: true }
            }
        })));

        const blank = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-1']");
        const zero = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']");
        expect(blank?.classList.contains("no-data")).toBe(true);
        expect(zero?.classList.contains("no-data")).toBe(false);
        expect(zero?.style.background).toContain("rgb");
    });

    it("extends a touch-compatible pointer range before applying it", () => {
        const { visual, element, applied } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)]
        })));

        const start = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']")!;
        const down = new Event("pointerdown", { bubbles: true });
        (down as unknown as { pointerType: string }).pointerType = "touch";
        start.dispatchEvent(down);

        const end = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']")!;
        const move = new Event("pointermove", { bubbles: true });
        (move as unknown as { pointerType: string }).pointerType = "touch";
        end.dispatchEvent(move);
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!
            .dispatchEvent(new Event("pointerup", { bubbles: true }));

        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(6);
        const merges = applied.filter((entry) => entry.action === 0 && entry.filter);
        const filter = merges.at(-1)?.filter as {
            conditions: Array<{ value: string }>;
        };
        expect(filter.conditions.map((condition) => condition.value)).toEqual([
            "2024-03-10T00:00:00.000Z",
            "2024-03-16T00:00:00.000Z"
        ]);
    });

    it("does not extend a touch range onto an aria-disabled day", () => {
        const { visual, element, applied } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 10)],
            values: [1],
            objects: { heatmap: { datesWithDataOnly: true } }
        })));

        element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']")!
            .dispatchEvent(new Event("pointerdown", { bubbles: true }));
        element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-12']")!
            .dispatchEvent(new Event("pointermove", { bubbles: true }));
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!
            .dispatchEvent(new Event("pointerup", { bubbles: true }));

        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(1);
        const filter = applied.filter((entry) => entry.action === 0).at(-1)?.filter as {
            conditions: Array<{ value: string }>;
        };
        expect(filter.conditions.map((condition) => condition.value)).toEqual([
            "2024-03-10T00:00:00.000Z",
            "2024-03-11T00:00:00.000Z"
        ]);
    });

    it("keeps vertical touch movement available for scrolling", () => {
        const { visual, element, applied } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 10), new Date(2024, 2, 15)]
        })));

        const cell = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']")!;
        const down = new Event("pointerdown", { bubbles: true });
        Object.assign(down, { pointerType: "touch", clientX: 10, clientY: 10, pointerId: 1 });
        cell.dispatchEvent(down);
        const move = new Event("pointermove", { bubbles: true });
        Object.assign(move, { pointerType: "touch", clientX: 10, clientY: 40, pointerId: 1 });
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!.dispatchEvent(move);
        const up = new Event("pointerup", { bubbles: true });
        Object.assign(up, { pointerType: "touch", pointerId: 1 });
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!.dispatchEvent(up);

        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(0);
    });

    it("owns a gesture by pointerId and ignores competing move/up/cancel events", () => {
        const { visual, element, applied } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 10), new Date(2024, 2, 15)]
        })));

        const first = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']")!;
        const down = new Event("pointerdown", { bubbles: true });
        Object.assign(down, { pointerType: "mouse", pointerId: 1, button: 0 });
        first.dispatchEvent(down);

        const competing = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']")!;
        const secondDown = new Event("pointerdown", { bubbles: true });
        Object.assign(secondDown, { pointerType: "mouse", pointerId: 2, button: 0 });
        competing.dispatchEvent(secondDown);
        const secondEnter = new Event("pointerenter", { bubbles: true });
        Object.assign(secondEnter, { pointerType: "mouse", pointerId: 2 });
        competing.dispatchEvent(secondEnter);
        const secondMove = new Event("pointermove", { bubbles: true });
        Object.assign(secondMove, { pointerType: "mouse", pointerId: 2 });
        competing.dispatchEvent(secondMove);
        const secondCancel = new Event("pointercancel", { bubbles: true });
        Object.assign(secondCancel, { pointerType: "mouse", pointerId: 2 });
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!.dispatchEvent(secondCancel);
        const secondUp = new Event("pointerup", { bubbles: true });
        Object.assign(secondUp, { pointerType: "mouse", pointerId: 2 });
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!.dispatchEvent(secondUp);

        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(0);
        expect(element.querySelectorAll(".cs-day.selected")).toHaveLength(1);

        const ownerUp = new Event("pointerup", { bubbles: true });
        Object.assign(ownerUp, { pointerType: "mouse", pointerId: 1 });
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!.dispatchEvent(ownerUp);
        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(1);
        expect(element.querySelector(".cs-day[data-key='2024-2-15']")!
            .classList.contains("selected")).toBe(false);
    });

    it("announces both endpoints after a range selection", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 10), new Date(2024, 2, 15)]
        })));

        element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-10']")!
            .dispatchEvent(new Event("pointerdown", { bubbles: true }));
        element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']")!
            .dispatchEvent(new Event("pointermove", { bubbles: true }));
        element.querySelector<HTMLElement>(".atlynCalendarSlicer")!
            .dispatchEvent(new Event("pointerup", { bubbles: true }));

        const announcement = element.querySelector("#cs-live-status")?.textContent || "";
        expect(announcement).toContain("March 10, 2024");
        expect(announcement).toContain("March 15, 2024");
    });

    it("keeps a 5,000-day selection synchronized at the limit", () => {
        const { visual, element, applied } = createVisual();
        const start = new Date(2020, 0, 1);
        const values = Array.from({ length: 4999 }, (_, index) =>
            serializeDateNaive(addDays(start, index))
        );
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2100, 0, 1)],
            objects: { general: { visibleYear: 2100, visibleMonth: 0 } }
        }), [{
            target: { table: "Calendar", column: "Date" },
            operator: "In",
            values
        }]));

        const first = element.querySelector<HTMLElement>(".cs-day[data-key='2100-0-1']")!;
        const addAtLimit = new Event("pointerdown", { bubbles: true });
        Object.assign(addAtLimit, { ctrlKey: true, button: 0 });
        first.dispatchEvent(addAtLimit);
        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(1);
        expect(element.querySelector<HTMLElement>(".cs-day[data-key='2100-0-1']")
            ?.classList.contains("selected")).toBe(true);
        expect(element.querySelector("#cs-live-status")?.textContent)
            .toContain("5,000");

        const second = element.querySelector<HTMLElement>(".cs-day[data-key='2100-0-2']")!;
        const overLimit = new Event("pointerdown", { bubbles: true });
        Object.assign(overLimit, { ctrlKey: true, button: 0 });
        second.dispatchEvent(overLimit);
        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(1);
        expect(second.classList.contains("selected")).toBe(false);
        expect(element.querySelector("#cs-live-status")?.textContent)
            .toContain("5,000");
    });

    it("rejects toggling inside an oversized contiguous range", () => {
        const { visual, element, applied } = createVisual();
        const start = new Date(2020, 0, 1);
        visual.update(updateOptions(buildMockDataView({
            dates: [start],
            objects: { general: { visibleYear: 2020, visibleMonth: 0 } }
        }), [{
            target: { table: "Calendar", column: "Date" },
            logicalOperator: "And",
            conditions: [
                { operator: "GreaterThanOrEqual", value: serializeDate(start) },
                { operator: "LessThan", value: serializeDate(addDays(start, 5001)) }
            ]
        }]));

        const inside = element.querySelector<HTMLElement>(".cs-day[data-key='2020-0-2']")!;
        expect(inside.classList.contains("selected")).toBe(true);
        const toggle = new Event("pointerdown", { bubbles: true });
        Object.assign(toggle, { ctrlKey: true, button: 0 });
        inside.dispatchEvent(toggle);

        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(0);
        expect(inside.classList.contains("selected")).toBe(true);
        expect(element.querySelector("#cs-live-status")?.textContent)
            .toContain("5,000");
    });

    it("uses data-point and empty-space SelectionIds without filtering on context menus", () => {
        const { visual, element, applied, contextMenus } = createVisual();
        visual.update(updateOptions(buildMockDataView({
            dates: [new Date(2024, 2, 15)]
        })));

        const cell = element.querySelector<HTMLElement>(".cs-day[data-key='2024-2-15']")!;
        const rightDown = new Event("pointerdown", { bubbles: true });
        Object.assign(rightDown, { button: 2, pointerType: "mouse" });
        cell.dispatchEvent(rightDown);
        expect(applied).toHaveLength(0);

        cell.dispatchEvent(new MouseEvent("contextmenu", {
            bubbles: true,
            clientX: 10,
            clientY: 20
        }));
        expect(contextMenus[0]?.[0]).toEqual({ key: "date:0" });

        const grid = element.querySelector<HTMLElement>(".cs-grid")!;
        grid.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true }));
        expect(contextMenus[1]?.[0]).toEqual({});
        expect(applied).toHaveLength(0);
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
        expect(element.querySelector(".cs-disclosure")?.textContent).toContain("30,000");
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

    it("cleans DOM and pointer listeners on destroy", () => {
        const { visual, element } = createVisual();
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 15)] })));
        visual.destroy();
        expect(element.querySelector(".atlynCalendarSlicer")?.childElementCount).toBe(0);
        visual.update(updateOptions(buildMockDataView({ dates: [new Date(2024, 2, 15)] })));
        expect(element.querySelector(".cs-grid")).toBeNull();
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
