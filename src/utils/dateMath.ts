/**
 * Timezone- and DST-safe date math for the Atlyn Calendar Slicer.
 *
 * The #1 source of off-by-one bugs in date visuals is that `Date.toJSON()` is
 * always UTC: in UTC-7, `new Date(2024, 2, 15)` serialises as
 * `2024-03-15T07:00:00.000Z`. Power BI date columns are timezone-naive (local
 * wall-clock), so every date sent to a filter must be serialised as the local
 * wall-clock instant relabelled as UTC. All serialisation is routed through the
 * single `serializeDate` helper; all parsing through `parseDateWithoutTimezone`.
 *
 * Helpers ported from Microsoft's powerbi-visuals-timeline.
 */

export const MS_PER_MINUTE = 60 * 1000;
export const MS_PER_DAY = 24 * 60 * MS_PER_MINUTE;

/** Weekday index of the first day of the week: 0=Sun, 1=Mon, 6=Sat. */
export type WeekStart = 0 | 1 | 6;

export interface CalendarDay {
    /** Local-midnight Date for this cell. */
    date: Date;
    /** Whether the day belongs to the month being displayed. */
    inMonth: boolean;
}

/**
 * Epoch milliseconds shifted so that a UTC formatter prints the LOCAL wall
 * clock. Ported from Timeline's GET_MILLISECONDS_WITHOUT_TIMEZONE.
 */
export function getMillisecondsWithoutTimezone(date: Date): number {
    return date.getTime() - date.getTimezoneOffset() * MS_PER_MINUTE;
}

/**
 * Inverse of {@link getMillisecondsWithoutTimezone}: parse an ISO string that
 * encodes a wall-clock instant back into a local Date. Ported from Timeline's
 * PARSE_DATE_WITHOUT_TIMEZONE.
 */
export function parseDateWithoutTimezone(iso: string): Date {
    const parsed = new Date(iso);
    if (isNaN(parsed.getTime())) {
        return parsed;
    }
    return new Date(parsed.getTime() + parsed.getTimezoneOffset() * MS_PER_MINUTE);
}

/**
 * DST gap between two dates, in milliseconds. Ported from Timeline's
 * GET_DAYLIGHT_SAVING_TIME_OFF.
 */
export function getDaylightSavingTimeOff(from: Date, to: Date): number {
    return (to.getTimezoneOffset() - from.getTimezoneOffset()) * MS_PER_MINUTE;
}

/**
 * The single serialisation path for filter boundaries. Produces an ISO-8601
 * string whose UTC fields equal the LOCAL wall clock of `date`, e.g. local
 * midnight of 2024-03-15 in any timezone serialises to
 * "2024-03-15T00:00:00.000Z".
 */
export function serializeDate(date: Date): string {
    return new Date(getMillisecondsWithoutTimezone(date)).toISOString();
}

/**
 * Serialise a date as a NAIVE local wall-clock ISO string with NO timezone
 * designator, e.g. local midnight of 2024-03-15 -> "2024-03-15T00:00:00".
 *
 * Used ONLY for the discrete-value `BasicFilter ("In")` path, never for range
 * boundaries. The distinction is load-bearing:
 *
 *   - A range boundary (`GreaterThanOrEqual` / `LessThan`) is compared
 *     TOLERANTLY, so the UTC-relabelled `serializeDate` form
 *     ("2024-03-15T00:00:00.000Z") brackets local-stored values correctly.
 *   - An `In` filter requires EXACT equality against the model's stored value.
 *     A trailing `Z` makes Power BI treat the value as UTC and convert it into
 *     the model timezone before comparing, so the UTC-relabelled local midnight
 *     never equals the local-midnight DateTime in the column and the filter
 *     matches nothing (empty result). Emitting the bare wall clock keeps the
 *     value in model-local time so exact equality holds.
 *
 * Because the fields are read off the LOCAL Date, the output is timezone
 * invariant: local midnight of 2024-03-15 serialises to "2024-03-15T00:00:00"
 * in every host timezone.
 */
export function serializeDateNaive(date: Date): string {
    const pad = (n: number): string => String(n).padStart(2, "0");
    return (
        `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
        `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
    );
}

/** Build a local-midnight Date. */
export function makeDate(year: number, month: number, day: number): Date {
    return new Date(year, month, day);
}

/** Local midnight of the given date. */
export function startOfDay(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Add days using the local-date constructor so the result stays at local
 * midnight even across a DST transition (never raw millisecond arithmetic).
 */
export function addDays(date: Date, days: number): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

/** Add months, clamped to the last valid day of the target month. */
export function addMonths(date: Date, months: number): Date {
    const year = date.getFullYear();
    const month = date.getMonth() + months;
    const day = Math.min(date.getDate(), daysInMonth(year + Math.floor(month / 12), ((month % 12) + 12) % 12));
    return new Date(year, month, day);
}

/** First day of the month containing `date`, at local midnight. */
export function startOfMonth(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), 1);
}

/** First day of the month AFTER the one containing `date` (half-open end). */
export function startOfNextMonth(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth() + 1, 1);
}

/** Number of days in the given month (month is 0-based). */
export function daysInMonth(year: number, month: number): number {
    return new Date(year, month + 1, 0).getDate();
}

/** Whole calendar days between two dates (DST-neutralised). */
export function diffInDays(from: Date, to: Date): number {
    return Math.round(
        (getMillisecondsWithoutTimezone(startOfDay(to)) -
            getMillisecondsWithoutTimezone(startOfDay(from))) / MS_PER_DAY
    );
}

export function isSameDay(a: Date, b: Date): boolean {
    return a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
}

/**
 * Start of the week containing `date`, given the configured week-start weekday.
 */
export function startOfWeek(date: Date, weekStart: WeekStart): Date {
    const offset = (date.getDay() - weekStart + 7) % 7;
    return addDays(startOfDay(date), -offset);
}

/**
 * ISO-8601 week number (weeks start Monday; week 1 contains the first Thursday
 * of the year). DST-neutralised so the day count is always an exact multiple.
 */
export function getISOWeek(date: Date): number {
    const target = startOfDay(date);
    const dayNr = (target.getDay() + 6) % 7; // Mon=0 ... Sun=6
    const thursday = addDays(target, 3 - dayNr);
    const firstJan = new Date(thursday.getFullYear(), 0, 1);
    const days = Math.round(
        (getMillisecondsWithoutTimezone(thursday) -
            getMillisecondsWithoutTimezone(firstJan)) / MS_PER_DAY
    );
    return Math.floor(days / 7) + 1;
}

/** The ISO week-numbering year (may differ from the calendar year near Jan 1). */
export function getISOWeekYear(date: Date): number {
    const target = startOfDay(date);
    const dayNr = (target.getDay() + 6) % 7;
    return addDays(target, 3 - dayNr).getFullYear();
}

/**
 * Build a month grid: whole weeks (each 7 days) covering `month`, padded with
 * leading/trailing days from adjacent months. Returns 4–6 week rows as needed.
 */
export function buildMonthGrid(
    year: number,
    month: number,
    weekStart: WeekStart
): CalendarDay[][] {
    const first = new Date(year, month, 1);
    const gridStart = startOfWeek(first, weekStart);
    const total = daysInMonth(year, month);
    const leading = diffInDays(gridStart, first);
    const weeks = Math.ceil((leading + total) / 7);

    const rows: CalendarDay[][] = [];
    let cursor = gridStart;
    for (let w = 0; w < weeks; w++) {
        const row: CalendarDay[] = [];
        for (let d = 0; d < 7; d++) {
            row.push({ date: cursor, inMonth: cursor.getMonth() === month });
            cursor = addDays(cursor, 1);
        }
        rows.push(row);
    }
    return rows;
}

/**
 * First month of the fiscal quarter containing `date`, given a fiscal-year
 * start month (1=January ... 12=December). Returns a local-midnight Date on the
 * first of that month.
 */
export function fiscalQuarterStart(date: Date, fiscalStartMonth: number): Date {
    const fyStart = fiscalStartMonth - 1; // 0-based
    const monthsSinceFyStart = ((date.getMonth() - fyStart) + 12) % 12;
    const quarterIndex = Math.floor(monthsSinceFyStart / 3);
    const startMonthOffset = quarterIndex * 3;
    const monthsBack = monthsSinceFyStart - startMonthOffset;
    return new Date(date.getFullYear(), date.getMonth() - monthsBack, 1);
}

/**
 * First day of the fiscal year containing `date`, given a fiscal-year start
 * month (1=January ... 12=December).
 */
export function fiscalYearStart(date: Date, fiscalStartMonth: number): Date {
    const fyStart = fiscalStartMonth - 1;
    const monthsSinceFyStart = ((date.getMonth() - fyStart) + 12) % 12;
    return new Date(date.getFullYear(), date.getMonth() - monthsSinceFyStart, 1);
}
