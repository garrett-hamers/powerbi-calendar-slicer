import { describe, expect, it } from "vitest";
import { RelativeDateFilterTimeUnit, RelativeDateOperators } from "powerbi-models";
import {
    buildDayFilter,
    buildMultiDayFilter,
    buildRangeFilter,
    buildRelativeFilter,
    targetFromQueryName,
    FilterTarget
} from "../src/dateFilter";
import { makeDate } from "../src/utils/dateMath";

const target: FilterTarget = { table: "Calendar", column: "Date" };

describe("targetFromQueryName", () => {
    it("splits on the first dot into table + column", () => {
        expect(targetFromQueryName("Calendar.Date")).toEqual({
            table: "Calendar",
            column: "Date"
        });
    });

    it("keeps hierarchy levels in the column remainder", () => {
        expect(targetFromQueryName("Calendar.Date Hierarchy.Date")).toEqual({
            table: "Calendar",
            column: "Date Hierarchy.Date"
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
    it("uses In with de-duplicated, sorted ISO values", () => {
        const f = buildMultiDayFilter(
            [makeDate(2024, 2, 15), makeDate(2024, 2, 1), makeDate(2024, 2, 15)],
            target
        );
        expect(f.operator).toBe("In");
        expect(f.values).toEqual([
            "2024-03-01T00:00:00.000Z",
            "2024-03-15T00:00:00.000Z"
        ]);
        expect(f.target).toEqual(target);
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
