/**
 * Filter builders for the Atlyn Calendar Slicer.
 *
 * Selection semantics map to the Power BI filter APIs as follows:
 *   - contiguous range      -> AdvancedFilter (GreaterThanOrEqual + LessThan)
 *   - non-contiguous days    -> BasicFilter ("In", array of naive-local dates)
 *   - relative presets       -> RelativeDateFilter (InLast / InThis)
 *
 * Ranges are ALWAYS half-open: [start, endExclusive) where endExclusive is the
 * start of the next period. Using LessThanOrEqual would silently drop fact rows
 * whose date carries a time component (e.g. 2024-03-31T14:30:00). Verified
 * against powerbi-visuals-timeline (src/timeLine.ts).
 *
 * NOTE on the two serialisation forms: range boundaries use the UTC-relabelled
 * `serializeDate` form ("...T00:00:00.000Z"), which Power BI compares tolerantly.
 * The discrete `BasicFilter ("In")` path instead uses the NAIVE local form
 * `serializeDateNaive` ("...T00:00:00", no Z). An `In` filter matches by EXACT
 * equality; a trailing Z is interpreted as UTC and converted into the model
 * timezone before comparison, so a UTC-relabelled local midnight never equals
 * the local-midnight DateTime stored in the column and the filter matches
 * nothing. This asymmetry is deliberate and is locked by unit tests.
 *
 * A single AdvancedFilter cannot express a non-contiguous selection: powerbi-
 * models caps AdvancedFilter at two conditions (models.js:437-438,
 * "AdvancedFilters may not have more than two conditions"), and an array of
 * filters passed to applyJsonFilter is AND-combined, not OR-combined. So
 * OR-of-half-open-day-ranges is not expressible, and BasicFilter remains the
 * correct primitive for discrete multi-day selection.
 */
import {
    AdvancedFilter,
    BasicFilter,
    RelativeDateFilter,
    IFilterColumnTarget,
    RelativeDateFilterTimeUnit,
    RelativeDateOperators
} from "powerbi-models";
import { addDays, serializeDate, serializeDateNaive, startOfDay } from "./utils/dateMath";

export type FilterTarget = IFilterColumnTarget;

/**
 * Maximum number of discrete values accepted by the BasicFilter path. A
 * contiguous selection remains an AdvancedFilter regardless of its size; only
 * Ctrl/⌘-click conversion to a discrete list is bounded.
 */
export const MAX_DISCRETE_DAYS = 5000;

/**
 * Derive a concrete column filter target from the semantic query name emitted
 * by Power BI. Display captions are intentionally ignored: captions can be
 * localized or renamed and are not valid filter targets.
 */
export function targetFromQueryName(
    queryName: string,
    _columnDisplayName?: string
): FilterTarget | null {
    if (!queryName) {
        return null;
    }
    const dot = queryName.indexOf(".");
    if (dot <= 0 || dot === queryName.length - 1 ||
        /hierarchy/i.test(queryName.slice(dot + 1))) {
        return null;
    }
    const column = queryName.slice(dot + 1);
    if (!column) {
        return null;
    }
    return {
        table: queryName.slice(0, dot),
        column
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
 * serialised through the NAIVE local wall-clock path (no `Z`) so Power BI's
 * exact-match `In` semantics compare against the model's local-midnight
 * DateTime values. Using the UTC-relabelled range form here produces a filter
 * that matches nothing — see the module header and serializeDateNaive.
 */
export function buildMultiDayFilter(days: Date[], target: FilterTarget): BasicFilter {
    if (days.length > MAX_DISCRETE_DAYS) {
        throw new RangeError(
            `A discrete date selection cannot contain more than ${MAX_DISCRETE_DAYS} days`
        );
    }
    const seen = new Set<string>();
    const values: string[] = [];
    for (const day of days) {
        const iso = serializeDateNaive(startOfDay(day));
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
