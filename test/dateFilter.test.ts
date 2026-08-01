import { describe, expect, it } from "vitest";
import { RelativeDateFilterTimeUnit, RelativeDateOperators } from "powerbi-models";
import {
    buildDayFilter,
    buildMultiDayFilter,
    buildRangeFilter,
    buildRelativeFilter,
    MAX_DISCRETE_DAYS,
    targetFromQueryName,
    FilterTarget
} from "../src/dateFilter";
import { makeDate } from "../src/utils/dateMath";

const target: FilterTarget = { table: "Calendar", column: "Date" };

describe("targetFromQueryName", () => {
    it("splits a concrete table + column path", () => {
        expect(targetFromQueryName("Calendar.Date")).toEqual({
            table: "Calendar",
            column: "Date"
        });
    });

    it("rejects hierarchy query names instead of deriving targets from captions", () => {
        expect(targetFromQueryName("Calendar.Date Hierarchy.Date", "Renamed Date")).toBeNull();
    });

    it("rejects extra semantic path separators without relying on a hierarchy label", () => {
        expect(targetFromQueryName("Calendar.Fiscal.Date")).toBeNull();
    });

    it("rejects case-insensitive hierarchy variants", () => {
        expect(targetFromQueryName("Calendar.Date hIeRaRcHy.Month")).toBeNull();
    });

    it("ignores a localized or renamed display caption", () => {
        expect(targetFromQueryName("Calendar.Date", "Datum")).toEqual({
            table: "Calendar",
            column: "Date"
        });
    });

    it("accepts legitimate identifiers containing the word hierarchy", () => {
        expect(targetFromQueryName("HierarchyCalendar.Date")).toEqual({
            table: "HierarchyCalendar",
            column: "Date"
        });
        expect(targetFromQueryName("Calendar.DateHierarchy")).toEqual({
            table: "Calendar",
            column: "DateHierarchy"
        });
    });

    it("rejects malformed query names", () => {
        expect(targetFromQueryName("")).toBeNull();
        expect(targetFromQueryName("NoDot")).toBeNull();
        expect(targetFromQueryName(".Leading")).toBeNull();
        expect(targetFromQueryName("Trailing.")).toBeNull();
    });
});

describe("range filter", () => {
    // Mandatory certification lock: ranges must be half-open.
    it("uses a half-open interval for range filters", () => {
        const f = buildRangeFilter(makeDate(2024, 2, 1), makeDate(2024, 3, 1), target);
        expect(f.conditions[0].operator).toBe("GreaterThanOrEqual");
        expect(f.conditions[1].operator).toBe("LessThan"); // never LessThanOrEqual
    });

    it("serialises both bounds as local-midnight UTC and And-combines them", () => {
        const f = buildRangeFilter(makeDate(2024, 2, 1), makeDate(2024, 3, 1), target);
        expect(f.logicalOperator).toBe("And");
        expect(f.conditions[0].value).toBe("2024-03-01T00:00:00.000Z");
        expect(f.conditions[1].value).toBe("2024-04-01T00:00:00.000Z");
        expect(f.target).toEqual(target);
    });

    it("builds a one-day half-open range for a single day", () => {
        const f = buildDayFilter(makeDate(2024, 2, 15), target);
        expect(f.conditions[0].value).toBe("2024-03-15T00:00:00.000Z");
        expect(f.conditions[1].value).toBe("2024-03-16T00:00:00.000Z");
        expect(f.conditions[1].operator).toBe("LessThan");
    });
});

describe("multi-day basic filter", () => {
    it("rejects an oversized discrete payload", () => {
        const days = Array.from({ length: MAX_DISCRETE_DAYS + 1 }, (_, index) =>
            makeDate(2024, 0, index + 1)
        );
        expect(() => buildMultiDayFilter(days, target)).toThrow(RangeError);
    });

    it("uses In with de-duplicated, sorted naive-local values", () => {
        const f = buildMultiDayFilter(
            [makeDate(2024, 2, 15), makeDate(2024, 2, 1), makeDate(2024, 2, 15)],
            target
        );
        expect(f.operator).toBe("In");
        expect(f.values).toEqual([
            "2024-03-01T00:00:00",
            "2024-03-15T00:00:00"
        ]);
        expect(f.target).toEqual(target);
    });

    // Regression lock for the v1.0.0.1 empty-result bug: BasicFilter ("In")
    // exact-matches a datetime column, so a trailing "Z" is read as UTC,
    // converted into the model timezone, and never equals the local-midnight
    // value stored in the column. The values MUST be naive local wall clock.
    // This assertion fails against v1.0.0.0, which emitted "...T00:00:00.000Z".
    it("emits naive local values with no UTC 'Z' designator", () => {
        const f = buildMultiDayFilter([makeDate(2024, 2, 9), makeDate(2024, 2, 11)], target);
        for (const v of f.values as string[]) {
            expect(v.endsWith("Z")).toBe(false);
            expect(v).toMatch(/^\d{4}-\d{2}-\d{2}T00:00:00$/);
        }
    });

    // Each "In" value must denote the SAME calendar day that a single-day
    // filter (the proven AdvancedFilter range path) lower-bounds, so a one-day
    // multi-select and a single click filter to the identical day.
    it("denotes the same local day as the single-day range lower bound", () => {
        const day = makeDate(2024, 2, 15);
        const multi = buildMultiDayFilter([day], target);
        const single = buildDayFilter(day, target);
        const multiDay = (multi.values[0] as string).slice(0, 10);
        const singleLower = (single.conditions[0].value as string).slice(0, 10);
        expect(multiDay).toBe("2024-03-15");
        expect(multiDay).toBe(singleLower);
    });
});

describe("relative filter", () => {
    it("builds an InLast N-days relative filter", () => {
        const f = buildRelativeFilter(
            {
                operator: RelativeDateOperators.InLast,
                count: 7,
                unit: RelativeDateFilterTimeUnit.Days,
                includeToday: true
            },
            target
        );
        expect(f.operator).toBe(RelativeDateOperators.InLast);
        expect(f.timeUnitsCount).toBe(7);
        expect(f.timeUnitType).toBe(RelativeDateFilterTimeUnit.Days);
        expect(f.includeToday).toBe(true);
        expect(f.target).toEqual(target);
    });

    it("builds an InThis calendar-month relative filter", () => {
        const f = buildRelativeFilter(
            {
                operator: RelativeDateOperators.InThis,
                count: 1,
                unit: RelativeDateFilterTimeUnit.CalendarMonths,
                includeToday: true
            },
            target
        );
        expect(f.operator).toBe(RelativeDateOperators.InThis);
        expect(f.timeUnitType).toBe(RelativeDateFilterTimeUnit.CalendarMonths);
    });
});
