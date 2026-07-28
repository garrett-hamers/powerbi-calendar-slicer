/**
 * Filter builders for the Atlyn Calendar Slicer.
 *
 * Selection semantics map to the Power BI filter APIs as follows:
 *   - contiguous range      -> AdvancedFilter (GreaterThanOrEqual + LessThan)
 *   - non-contiguous days    -> BasicFilter ("In", array of ISO dates)
 *   - relative presets       -> RelativeDateFilter (InLast / InThis)
 *
 * Ranges are ALWAYS half-open: [start, endExclusive) where endExclusive is the
 * start of the next period. Using LessThanOrEqual would silently drop fact rows
 * whose date carries a time component (e.g. 2024-03-31T14:30:00). Verified
 * against powerbi-visuals-timeline (src/timeLine.ts).
 */
import {
    AdvancedFilter,
    BasicFilter,
    RelativeDateFilter,
    IFilterColumnTarget,
    RelativeDateFilterTimeUnit,
    RelativeDateOperators
} from "powerbi-models";
import { addDays, serializeDate, startOfDay } from "./utils/dateMath";

export type FilterTarget = IFilterColumnTarget;

/**
 * Derive a column filter target from a categorical column `queryName`. Power BI
 * query names are `Table.Column` (or `Table.Hierarchy.Level`); the table is the
 * part before the first dot, the column is the remainder.
 */
export function targetFromQueryName(queryName: string): FilterTarget | null {
    if (!queryName) {
        return null;
    }
    const dot = queryName.indexOf(".");
    if (dot <= 0 || dot === queryName.length - 1) {
        return null;
    }
    return {
        table: queryName.slice(0, dot),
        column: queryName.slice(dot + 1)
    };
}

/**
 * Contiguous range filter over the half-open interval [start, endExclusive).
 * Both bounds are normalised to local midnight before serialisation.
 */
export function buildRangeFilter(
    start: Date,
    endExclusive: Date,
    target: FilterTarget
): AdvancedFilter {
    return new AdvancedFilter(
        target,
        "And",
        { operator: "GreaterThanOrEqual", value: serializeDate(startOfDay(start)) },
        { operator: "LessThan", value: serializeDate(startOfDay(endExclusive)) }
    );
}

/** Single-day filter: the half-open interval [day, day + 1). */
export function buildDayFilter(day: Date, target: FilterTarget): AdvancedFilter {
    const start = startOfDay(day);
    return buildRangeFilter(start, addDays(start, 1), target);
}

/**
 * Non-contiguous multi-day filter. Days are de-duplicated, sorted, and
 * serialised through the single TZ-safe path.
 */
export function buildMultiDayFilter(days: Date[], target: FilterTarget): BasicFilter {
    const seen = new Set<string>();
    const values: string[] = [];
    for (const day of days) {
        const iso = serializeDate(startOfDay(day));
        if (!seen.has(iso)) {
            seen.add(iso);
            values.push(iso);
        }
    }
    values.sort();
    return new BasicFilter(target, "In", values);
}

export interface RelativeSpec {
    operator: RelativeDateOperators;
    count: number;
    unit: RelativeDateFilterTimeUnit;
    includeToday: boolean;
}

/**
 * Relative-date filter (e.g. "in the last 7 days", "in this month"). Preferred
 * over a fixed range for presets so MTD/YTD stay correct as time passes rather
 * than freezing to the date a bookmark was saved.
 */
export function buildRelativeFilter(
    spec: RelativeSpec,
    target: FilterTarget
): RelativeDateFilter {
    return new RelativeDateFilter(
        target,
        spec.operator,
        spec.count,
        spec.unit,
        spec.includeToday
    );
}
