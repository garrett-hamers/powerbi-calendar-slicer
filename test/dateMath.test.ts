import { execFileSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import {
    addDays,
    addMonths,
    buildMonthGrid,
    daysInMonth,
    diffInDays,
    fiscalQuarterStart,
    fiscalYearStart,
    getISOWeek,
    getISOWeekYear,
    makeDate,
    parseDateWithoutTimezone,
    serializeDate,
    startOfMonth,
    startOfNextMonth,
    startOfWeek,
    WeekStart
} from "../src/utils/dateMath";

describe("serialisation (host timezone)", () => {
    it("serialises local midnight as the wall-clock instant relabelled UTC", () => {
        // Holds in ANY host timezone: the offset is neutralised before toISOString.
        expect(serializeDate(makeDate(2024, 2, 15))).toBe("2024-03-15T00:00:00.000Z");
        expect(serializeDate(makeDate(2024, 0, 1))).toBe("2024-01-01T00:00:00.000Z");
        expect(serializeDate(makeDate(2023, 11, 31))).toBe("2023-12-31T00:00:00.000Z");
    });

    it("round-trips through parseDateWithoutTimezone back to local midnight", () => {
        for (const d of [makeDate(2024, 2, 15), makeDate(2024, 6, 1), makeDate(2020, 1, 29)]) {
            const restored = parseDateWithoutTimezone(serializeDate(d));
            expect(restored.getFullYear()).toBe(d.getFullYear());
            expect(restored.getMonth()).toBe(d.getMonth());
            expect(restored.getDate()).toBe(d.getDate());
            expect(restored.getHours()).toBe(0);
        }
    });

    it("returns an invalid date for unparseable input rather than throwing", () => {
        expect(isNaN(parseDateWithoutTimezone("not-a-date").getTime())).toBe(true);
    });
});

/**
 * Real cross-timezone matrix. Each case spawns a fresh Node process with the TZ
 * environment variable pinned, so DST and half-hour offsets are genuinely
 * exercised (setting TZ mid-process is unreliable). The probe reproduces
 * serializeDate's formula and the constructor-based addDays, then we assert both
 * the timezone actually took effect (offsets differ) and that the serialised
 * output is timezone-invariant local midnight.
 */
describe("timezone matrix", () => {
    const probe = `
        const offset = (y, m, d) => new Date(y, m, d).getTimezoneOffset();
        const serialize = (dt) =>
            new Date(dt.getTime() - dt.getTimezoneOffset() * 60000).toISOString();
        const addDays = (dt, n) =>
            new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() + n);
        const out = {
            offsetJan: offset(2024, 0, 15),
            offsetMar: offset(2024, 2, 15),
            midMarch: serialize(new Date(2024, 2, 15)),
            // 2024-03-10 is the US spring-forward day; constructor add must land
            // on local midnight of the 10th, not 23:00 of the 9th.
            dstAdd: serialize(addDays(new Date(2024, 2, 9), 1))
        };
        process.stdout.write(JSON.stringify(out));
    `;

    function runIn(tz: string): {
        offsetJan: number;
        offsetMar: number;
        midMarch: string;
        dstAdd: string;
    } {
        const stdout = execFileSync(process.execPath, ["-e", probe], {
            env: { ...process.env, TZ: tz },
            encoding: "utf8"
        });
        return JSON.parse(stdout);
    }

    it("UTC: zero offset year-round, invariant serialisation", () => {
        const r = runIn("UTC");
        expect(r.offsetJan).toBe(0);
        expect(r.offsetMar).toBe(0);
        expect(r.midMarch).toBe("2024-03-15T00:00:00.000Z");
        expect(r.dstAdd).toBe("2024-03-10T00:00:00.000Z");
    });

    it("America/Los_Angeles: DST shifts the offset but not the serialised day", () => {
        const r = runIn("America/Los_Angeles");
        expect(r.offsetJan).toBe(480); // PST
        expect(r.offsetMar).toBe(420); // PDT (DST began 2024-03-10)
        expect(r.midMarch).toBe("2024-03-15T00:00:00.000Z");
        expect(r.dstAdd).toBe("2024-03-10T00:00:00.000Z");
    });

    it("Asia/Kolkata: half-hour offset still serialises to local midnight", () => {
        const r = runIn("Asia/Kolkata");
        expect(r.offsetJan).toBe(-330); // UTC+5:30
        expect(r.offsetMar).toBe(-330);
        expect(r.midMarch).toBe("2024-03-15T00:00:00.000Z");
        expect(r.dstAdd).toBe("2024-03-10T00:00:00.000Z");
    });
});

describe("date arithmetic", () => {
    it("adds days across a month boundary", () => {
        expect(serializeDate(addDays(makeDate(2024, 0, 31), 1)))
            .toBe("2024-02-01T00:00:00.000Z");
    });

    it("adds months, clamping to the last valid day", () => {
        expect(serializeDate(addMonths(makeDate(2024, 0, 31), 1)))
            .toBe("2024-02-29T00:00:00.000Z"); // Jan 31 + 1 month -> Feb 29 (leap)
        expect(serializeDate(addMonths(makeDate(2023, 0, 31), 1)))
            .toBe("2023-02-28T00:00:00.000Z");
    });

    it("computes days in month including leap February", () => {
        expect(daysInMonth(2024, 1)).toBe(29);
        expect(daysInMonth(2023, 1)).toBe(28);
        expect(daysInMonth(2024, 3)).toBe(30);
    });

    it("diffInDays counts whole calendar days", () => {
        expect(diffInDays(makeDate(2024, 2, 1), makeDate(2024, 2, 31))).toBe(30);
        expect(diffInDays(makeDate(2024, 2, 31), makeDate(2024, 2, 1))).toBe(-30);
    });

    it("startOfMonth / startOfNextMonth give half-open month bounds", () => {
        const d = makeDate(2024, 2, 15);
        expect(serializeDate(startOfMonth(d))).toBe("2024-03-01T00:00:00.000Z");
        expect(serializeDate(startOfNextMonth(d))).toBe("2024-04-01T00:00:00.000Z");
    });
});

describe("startOfWeek", () => {
    it("respects the configured week-start day", () => {
        const wed = makeDate(2024, 2, 13); // Wednesday 2024-03-13
        expect(serializeDate(startOfWeek(wed, 0))).toBe("2024-03-10T00:00:00.000Z"); // Sun
        expect(serializeDate(startOfWeek(wed, 1))).toBe("2024-03-11T00:00:00.000Z"); // Mon
        expect(serializeDate(startOfWeek(wed, 6))).toBe("2024-03-09T00:00:00.000Z"); // Sat
    });
});

describe("ISO-8601 week numbers", () => {
    it("matches known reference weeks", () => {
        expect(getISOWeek(makeDate(2024, 0, 1))).toBe(1);   // Mon 2024-01-01
        expect(getISOWeek(makeDate(2024, 11, 30))).toBe(1); // belongs to 2025 W1
        expect(getISOWeekYear(makeDate(2024, 11, 30))).toBe(2025);
        expect(getISOWeek(makeDate(2021, 0, 1))).toBe(53);  // belongs to 2020 W53
        expect(getISOWeekYear(makeDate(2021, 0, 1))).toBe(2020);
        expect(getISOWeek(makeDate(2024, 5, 3))).toBe(23);
    });
});

describe("month grid", () => {
    it("pads with adjacent-month days and covers whole weeks", () => {
        // March 2024 starts on a Friday; Sunday-start grid leads with Feb 25.
        const grid = buildMonthGrid(2024, 2, 0);
        expect(grid[0].length).toBe(7);
        expect(grid.every((week) => week.length === 7)).toBe(true);
        expect(serializeDate(grid[0][0].date)).toBe("2024-02-25T00:00:00.000Z");
        expect(grid[0][0].inMonth).toBe(false);

        const flat = grid.flat();
        const firstInMonth = flat.find((c) => c.inMonth);
        expect(serializeDate(firstInMonth!.date)).toBe("2024-03-01T00:00:00.000Z");
        expect(flat.filter((c) => c.inMonth).length).toBe(31);
        const lastCell = flat[flat.length - 1];
        expect(lastCell.date.getDay()).toBe(6); // Saturday closes a Sunday-start week
    });

    it("uses fewer rows for a short month that fits in five weeks", () => {
        // February 2021 started on a Monday; Monday-start grid needs 4 rows.
        const grid = buildMonthGrid(2021, 1, 1);
        expect(grid.length).toBe(4);
        expect(serializeDate(grid[0][0].date)).toBe("2021-02-01T00:00:00.000Z");
    });
});

describe("fiscal boundaries", () => {
    const weekStart: WeekStart = 0;
    void weekStart;

    it("July fiscal year: Q starts and FY start", () => {
        // Fiscal year starts in July (7). A date in March 2024 is in fiscal Q3
        // (Jan–Mar) of FY2024, which began 2023-07-01.
        const d = makeDate(2024, 2, 15);
        expect(serializeDate(fiscalQuarterStart(d, 7))).toBe("2024-01-01T00:00:00.000Z");
        expect(serializeDate(fiscalYearStart(d, 7))).toBe("2023-07-01T00:00:00.000Z");
    });

    it("January fiscal year equals the calendar year", () => {
        const d = makeDate(2024, 4, 20); // May
        expect(serializeDate(fiscalQuarterStart(d, 1))).toBe("2024-04-01T00:00:00.000Z");
        expect(serializeDate(fiscalYearStart(d, 1))).toBe("2024-01-01T00:00:00.000Z");
    });
});
