/**
 * Relative-date presets for the Atlyn Calendar Slicer.
 *
 * Each preset produces both a concrete window (used to highlight the grid) and
 * the filter to apply. Host-relative presets use a RelativeDateFilter so they
 * stay correct over time — including inside a saved bookmark — instead of
 * freezing to the date the preset was clicked. To-date, custom-week, and fiscal
 * presets have no faithful relative-unit representation, so they resolve to a
 * half-open range computed at apply time.
 */
import { Filter, RelativeDateFilterTimeUnit, RelativeDateOperators } from "powerbi-models";
import { buildRangeFilter, buildRelativeFilter, FilterTarget } from "./dateFilter";
import {
    addDays,
    addMonths,
    fiscalQuarterStart,
    fiscalYearStart,
    makeDate,
    startOfDay,
    startOfMonth,
    startOfWeek,
    WeekStart
} from "./utils/dateMath";

export interface PresetContext {
    /** Local-midnight "today". */
    now: Date;
    /** Fiscal-year start month, 1 (January) .. 12 (December). */
    fiscalStartMonth: number;
    weekStart: WeekStart;
    target: FilterTarget;
}

export interface PresetResult {
    /** Filter to apply through host.applyJsonFilter. */
    filter: Filter;
    /** Inclusive window start (for grid highlight). */
    start: Date;
    /** Exclusive window end (for grid highlight). */
    endExclusive: Date;
}

export interface PresetDef {
    key: string;
    labelKey: string;
    label: string;
    compute(ctx: PresetContext): PresetResult;
}

function range(start: Date, endExclusive: Date, target: FilterTarget): PresetResult {
    return { filter: buildRangeFilter(start, endExclusive, target), start, endExclusive };
}

function rollingDays(count: number, ctx: PresetContext): PresetResult {
    const endExclusive = addDays(ctx.now, 1);
    const start = addDays(ctx.now, -(count - 1));
    const filter = buildRelativeFilter(
        {
            operator: RelativeDateOperators.InLast,
            count,
            unit: RelativeDateFilterTimeUnit.Days,
            includeToday: true
        },
        ctx.target
    );
    return { filter, start, endExclusive };
}

function relativeDays(
    operator: RelativeDateOperators,
    count: number,
    includeToday: boolean,
    start: Date,
    endExclusive: Date,
    ctx: PresetContext
): PresetResult {
    return {
        filter: buildRelativeFilter(
            {
                operator,
                count,
                unit: RelativeDateFilterTimeUnit.Days,
                includeToday
            },
            ctx.target
        ),
        start,
        endExclusive
    };
}

export const PRESETS: PresetDef[] = [
    {
        key: "today",
        labelKey: "Preset_Today",
        label: "Today",
        compute: (ctx) =>
            relativeDays(
                RelativeDateOperators.InThis,
                1,
                true,
                ctx.now,
                addDays(ctx.now, 1),
                ctx
            )
    },
    {
        key: "yesterday",
        labelKey: "Preset_Yesterday",
        label: "Yesterday",
        compute: (ctx) =>
            relativeDays(
                RelativeDateOperators.InLast,
                1,
                false,
                addDays(ctx.now, -1),
                ctx.now,
                ctx
            )
    },
    {
        key: "thisWeek",
        labelKey: "Preset_ThisWeek",
        label: "This Week",
        compute: (ctx) => {
            const start = startOfWeek(ctx.now, ctx.weekStart);
            if (ctx.weekStart === 0) {
                return {
                    filter: buildRelativeFilter(
                        {
                            operator: RelativeDateOperators.InThis,
                            count: 1,
                            unit: RelativeDateFilterTimeUnit.CalendarWeeks,
                            includeToday: true
                        },
                        ctx.target
                    ),
                    start,
                    endExclusive: addDays(start, 7)
                };
            }
            return range(start, addDays(start, 7), ctx.target);
        }
    },
    {
        key: "last7",
        labelKey: "Preset_Last7",
        label: "Last 7 Days",
        compute: (ctx) => rollingDays(7, ctx)
    },
    {
        key: "last14",
        labelKey: "Preset_Last14",
        label: "Last 14 Days",
        compute: (ctx) => rollingDays(14, ctx)
    },
    {
        key: "last30",
        labelKey: "Preset_Last30",
        label: "Last 30 Days",
        compute: (ctx) => rollingDays(30, ctx)
    },
    {
        key: "mtd",
        labelKey: "Preset_MTD",
        label: "Month to Date",
        compute: (ctx) => range(startOfMonth(ctx.now), addDays(ctx.now, 1), ctx.target)
    },
    {
        key: "qtd",
        labelKey: "Preset_QTD",
        label: "Quarter to Date",
        compute: (ctx) =>
            range(fiscalQuarterStart(ctx.now, ctx.fiscalStartMonth), addDays(ctx.now, 1), ctx.target)
    },
    {
        key: "ytd",
        labelKey: "Preset_YTD",
        label: "Year to Date",
        compute: (ctx) =>
            range(fiscalYearStart(ctx.now, ctx.fiscalStartMonth), addDays(ctx.now, 1), ctx.target)
    },
    {
        key: "lastMonth",
        labelKey: "Preset_LastMonth",
        label: "Last Month",
        compute: (ctx) => {
            const thisMonth = startOfMonth(ctx.now);
            const start = startOfMonth(addMonths(ctx.now, -1));
            const filter = buildRelativeFilter(
                {
                    operator: RelativeDateOperators.InLast,
                    count: 1,
                    unit: RelativeDateFilterTimeUnit.CalendarMonths,
                    includeToday: false
                },
                ctx.target
            );
            return { filter, start, endExclusive: thisMonth };
        }
    },
    {
        key: "lastQuarter",
        labelKey: "Preset_LastQuarter",
        label: "Last Quarter",
        compute: (ctx) => {
            const thisQuarter = fiscalQuarterStart(ctx.now, ctx.fiscalStartMonth);
            const prevQuarter = fiscalQuarterStart(addDays(thisQuarter, -1), ctx.fiscalStartMonth);
            return range(prevQuarter, thisQuarter, ctx.target);
        }
    },
    {
        key: "lastYear",
        labelKey: "Preset_LastYear",
        label: "Last Year",
        compute: (ctx) => {
            const start = makeDate(ctx.now.getFullYear() - 1, 0, 1);
            const endExclusive = makeDate(ctx.now.getFullYear(), 0, 1);
            const filter = buildRelativeFilter(
                {
                    operator: RelativeDateOperators.InLast,
                    count: 1,
                    unit: RelativeDateFilterTimeUnit.CalendarYears,
                    includeToday: false
                },
                ctx.target
            );
            return { filter, start, endExclusive };
        }
    }
];

export function presetByKey(key: string): PresetDef | undefined {
    return PRESETS.find((p) => p.key === key);
}

/** Local-midnight "today", exposed for a single deterministic source of now. */
export function today(): Date {
    return startOfDay(new Date());
}
