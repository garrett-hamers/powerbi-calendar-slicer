/**
 * Accessibility tests for the Atlyn Calendar Slicer visual: ARIA roles/labels,
 * roving tabindex + keyboard navigation, keyboard selection/clear, and the
 * high-contrast colour path.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import powerbi from "powerbi-visuals-api";
import { Visual } from "../src/visual";
import { buildMockDataView } from "./helpers/mockDataView";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

function createHost(highContrast = false) {
    const applied: Array<{ filter: unknown; action: number }> = [];
    const host = {
        locale: "en-US",
        colorPalette: {
            isHighContrast: highContrast,
            foreground: { value: "#ffffff" },
            background: { value: "#000000" }
        },
        eventService: {
            renderingStarted: vi.fn(),
            renderingFinished: vi.fn(),
            renderingFailed: vi.fn()
        },
        createSelectionManager: () => ({ showContextMenu: vi.fn() }),
        createLocalizationManager: () => ({ getDisplayName: (k: string) => k }),
        applyJsonFilter: (filter: unknown, _o: string, _p: string, action: number) => {
            applied.push({ filter, action });
        },
        persistProperties: vi.fn()
    };
    return { host, applied };
}

function mount(highContrast = false) {
    const element = document.createElement("div");
    document.body.appendChild(element);
    const { host, applied } = createHost(highContrast);
    const visual = new Visual({ element, host } as unknown as VisualConstructorOptions);
    return { visual, element, applied };
}

function update(visual: Visual, dataView: unknown) {
    visual.update({
        dataViews: [dataView],
        jsonFilters: [],
        type: 2,
        viewport: { width: 400, height: 400 }
    } as unknown as VisualUpdateOptions);
}

function focusedKey(element: HTMLElement): string | undefined {
    return element.querySelector<HTMLElement>(".cs-day[tabindex='0']")?.dataset.key;
}

function pressKey(element: HTMLElement, key: string) {
    const focused = element.querySelector<HTMLElement>(".cs-day[tabindex='0']");
    focused?.focus();
    (focused ?? element).dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
}

const marchData = () => buildMockDataView({
    dates: [new Date(2024, 2, 1), new Date(2024, 2, 15), new Date(2024, 2, 31)]
});

describe("Atlyn Calendar Slicer accessibility", () => {
    beforeEach(() => document.body.replaceChildren());

    it("exposes grid/gridcell roles, per-day labels and aria-selected", () => {
        const { visual, element } = mount();
        update(visual, marchData());

        const grid = element.querySelector("[role='grid']");
        expect(grid).not.toBeNull();
        expect(grid!.getAttribute("aria-multiselectable")).toBe("true");

        const gridcells = element.querySelectorAll(".cs-day[role='gridcell']");
        expect(gridcells.length).toBeGreaterThan(0);
        for (const cell of Array.from(gridcells)) {
            expect(cell.getAttribute("aria-label")).toBeTruthy();
            expect(cell.getAttribute("aria-selected")).toBe("false");
        }
    });

    it("keeps exactly one roving tabstop and moves focus with arrow keys", () => {
        const { visual, element } = mount();
        update(visual, marchData());

        const tabbable = element.querySelectorAll(".cs-day[tabindex='0']");
        expect(tabbable.length).toBe(1);

        const before = focusedKey(element);
        pressKey(element, "ArrowRight");
        const after = focusedKey(element);
        expect(after).not.toBe(before);
        expect(element.querySelectorAll(".cs-day[tabindex='0']").length).toBe(1);
    });

    it("selects the focused day with Enter and clears with Escape", () => {
        const { visual, element, applied } = mount();
        update(visual, marchData());

        pressKey(element, "Enter");
        const merges = applied.filter((a) => a.action === 0 && a.filter);
        expect(merges.length).toBeGreaterThanOrEqual(1);
        expect(element.querySelector(".cs-day.selected")).not.toBeNull();
        expect(document.activeElement).toBe(
            element.querySelector(".cs-day[tabindex='0']")
        );

        pressKey(element, "Escape");
        const removes = applied.filter((a) => a.action === 1);
        expect(removes.length).toBeGreaterThanOrEqual(1);
        expect(document.activeElement).toBe(
            element.querySelector(".cs-day[tabindex='0']")
        );
    });

    it("preserves focus when paging to another month", () => {
        const { visual, element } = mount();
        update(visual, marchData());

        pressKey(element, "PageDown");

        expect(document.activeElement).toBe(
            element.querySelector(".cs-day[tabindex='0']")
        );
    });

    it("reanchors the grid tab stop when a bookmark changes the visible month", () => {
        const { visual, element } = mount();
        update(visual, marchData());
        pressKey(element, "ArrowRight");

        update(visual, buildMockDataView({
            dates: [new Date(2024, 2, 1), new Date(2024, 2, 31)],
            objects: { general: { visibleYear: 2024, visibleMonth: 3 } }
        }));

        const focused = element.querySelector<HTMLElement>(".cs-day[tabindex='0']");
        expect(focused).not.toBeNull();
        expect(focused?.dataset.key?.startsWith("2024-3-")).toBe(true);
    });

    it("restores focus to the preset Today button rather than the toolbar Today button", () => {
        const { visual, element } = mount();
        update(visual, marchData());
        const presetToday = element.querySelector<HTMLButtonElement>(
            "button[data-focus-id='preset:today']"
        )!;
        presetToday.focus();
        presetToday.click();

        expect(document.activeElement).toBe(
            element.querySelector("button[data-focus-id='preset:today']")
        );
    });

    it("does not keyboard-select a date marked aria-disabled", () => {
        const { visual, element, applied } = mount();
        update(visual, buildMockDataView({
            dates: [new Date(2024, 2, 15)],
            values: [1],
            objects: { heatmap: { datesWithDataOnly: true } }
        }));

        const focused = element.querySelector<HTMLElement>(".cs-day[tabindex='0']");
        expect(focused?.getAttribute("aria-disabled")).toBe("true");
        pressKey(element, "Enter");

        expect(applied.filter((entry) => entry.action === 0)).toHaveLength(0);
        expect(document.activeElement).toBe(focused);
    });

    it("does not paint inline cell colours in high-contrast mode", () => {
        const { visual, element } = mount(true);
        update(visual, marchData());

        const root = element.querySelector(".atlynCalendarSlicer") as HTMLElement;
        expect(root.classList.contains("high-contrast")).toBe(true);

        const day = element.querySelector<HTMLElement>(".cs-day");
        expect(day).not.toBeNull();
        // Colours must come from the system palette / stylesheet, not inline styles.
        expect(day!.style.color).toBe("");
    });
});
