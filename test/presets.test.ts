/**
 * Unit tests for the relative-date presets: window boundaries, half-open range
 * filters, live RelativeDateFilter mapping, and fiscal-year awareness.
 */
import { describe, expect, it } from "vitest";
import { RelativeDateFilterTimeUnit, RelativeDateOperators } from "powerbi-models";
import { PRESETS, presetByKey, PresetContext } from "../src/presets";
import { makeDate, serializeDate } from "../src/utils/dateMath";

const target = { table: "Sales", column: "OrderDate" };

function ctx(now: Date, fiscalStartMonth = 1): PresetContext {
    return { now, fiscalStartMonth, weekStart: 0, target };
}

function compute(key: string, context: PresetContext) {
    const preset = presetByKey(key);
    expect(preset, `preset ${key} exists`).toBeDefined();
    return preset!.compute(context);
}

describe("relative-date presets", () => {
    it("defines the full documented preset set", () => {
        const keys = PRESETS.map((p) => p.key);
        expect(keys).toEqual([
            "today", "yesterday", "thisWeek",
            "last7", "last14", "last30",
            "mtd", "qtd", "ytd",
            "lastMonth", "lastQuarter", "lastYear"
        ]);
    });

    it("Today uses a live relative day filter", () => {
        const now = makeDate(2024, 4, 15);
        const r = compute("today", ctx(now));
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2024, 4, 15)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 4, 16)));
        const filter = r.filter as unknown as {
            operator: number;
            timeUnitsCount: number;
            timeUnitType: number;
            includeToday: boolean;
        };
        expect(filter.operator).toBe(RelativeDateOperators.InThis);
        expect(filter.timeUnitsCount).toBe(1);
        expect(filter.timeUnitType).toBe(RelativeDateFilterTimeUnit.Days);
        expect(filter.includeToday).toBe(true);
    });

    it("Yesterday uses a live relative day filter excluding today", () => {
        const now = makeDate(2024, 4, 15);
        const r = compute("yesterday", ctx(now));
        const filter = r.filter as unknown as {
            operator: number;
            timeUnitsCount: number;
            includeToday: boolean;
        };
        expect(filter.operator).toBe(RelativeDateOperators.InLast);
        expect(filter.timeUnitsCount).toBe(1);
        expect(filter.includeToday).toBe(false);
    });

    it("This Week is relative when the host week starts Sunday", () => {
        const r = compute("thisWeek", ctx(makeDate(2024, 4, 15)));
        const filter = r.filter as unknown as {
            operator: number;
            timeUnitType: number;
        };
        expect(filter.operator).toBe(RelativeDateOperators.InThis);
        expect(filter.timeUnitType).toBe(RelativeDateFilterTimeUnit.CalendarWeeks);
    });

    it("Last 7 Days is a live InLast/Days relative filter with a 7-day window", () => {
        const now = makeDate(2024, 4, 15);
        const r = compute("last7", ctx(now));
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2024, 4, 9)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 4, 16)));
        const filter = r.filter as unknown as {
            operator: number;
            timeUnitsCount: number;
            timeUnitType: number;
            includeToday: boolean;
        };
        expect(filter.operator).toBe(RelativeDateOperators.InLast);
        expect(filter.timeUnitsCount).toBe(7);
        expect(filter.timeUnitType).toBe(RelativeDateFilterTimeUnit.Days);
        expect(filter.includeToday).toBe(true);
    });

    it("Last Month maps to InLast 1 CalendarMonths and spans the prior month", () => {
        const now = makeDate(2024, 4, 15);
        const r = compute("lastMonth", ctx(now));
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2024, 3, 1)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 4, 1)));
        const filter = r.filter as unknown as {
            operator: number; timeUnitsCount: number; timeUnitType: number;
        };
        expect(filter.operator).toBe(RelativeDateOperators.InLast);
        expect(filter.timeUnitsCount).toBe(1);
        expect(filter.timeUnitType).toBe(RelativeDateFilterTimeUnit.CalendarMonths);
    });

    it("Last Year maps to InLast 1 CalendarYears and spans the prior calendar year", () => {
        const now = makeDate(2024, 4, 15);
        const r = compute("lastYear", ctx(now));
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2023, 0, 1)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 0, 1)));
        const filter = r.filter as unknown as {
            operator: number; timeUnitType: number;
        };
        expect(filter.operator).toBe(RelativeDateOperators.InLast);
        expect(filter.timeUnitType).toBe(RelativeDateFilterTimeUnit.CalendarYears);
    });

    it("YTD honours a July fiscal-year start", () => {
        const now = makeDate(2024, 4, 15); // 15 May 2024
        const r = compute("ytd", ctx(now, 7));
        // Fiscal year began 1 July 2023.
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2023, 6, 1)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 4, 16)));
    });

    it("QTD honours a July fiscal-year start", () => {
        const now = makeDate(2024, 4, 15); // 15 May 2024
        const r = compute("qtd", ctx(now, 7));
        // Fiscal quarter began 1 April 2024.
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2024, 3, 1)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 4, 16)));
    });

    it("Last Quarter spans the prior fiscal quarter", () => {
        const now = makeDate(2024, 4, 15);
        const r = compute("lastQuarter", ctx(now, 7));
        expect(serializeDate(r.start)).toBe(serializeDate(makeDate(2024, 0, 1)));
        expect(serializeDate(r.endExclusive)).toBe(serializeDate(makeDate(2024, 3, 1)));
    });
});
