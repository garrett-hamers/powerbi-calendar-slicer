/**
 * Mock DataView builder for Atlyn Calendar Slicer tests.
 *
 * Mirrors the shape the visual reads: a categorical DataView with a single
 * `Date` category column and an optional `Values` measure (with optional
 * highlights). The numeric option is retained only for rejection tests.
 */
export interface MockDataInput {
    dates: Array<Date | number | string | null | undefined>;
    values?: Array<number | null | undefined>;
    highlights?: Array<number | null>;
    dateDisplayName?: string;
    dateQueryName?: string;
    /** Column source type. Defaults to a dateTime column. */
    dateType?: "dateTime" | "numeric";
    valueDisplayName?: string;
    objects?: Record<string, Record<string, unknown>>;
}

export function buildMockDataView(input: MockDataInput): any {
    const columnType = input.dateType === "numeric"
        ? { numeric: true }
        : { dateTime: true };

    const categoryColumn = {
        source: {
            displayName: input.dateDisplayName || "Date",
            queryName: input.dateQueryName || "Calendar.Date",
            type: columnType,
            roles: { Date: true }
        },
        values: input.dates
    };

    const categories: any[] = [categoryColumn];
    const valueColumns: any[] = [];

    if (input.values) {
        valueColumns.push({
            source: {
                displayName: input.valueDisplayName || "Amount",
                queryName: "Calendar.Amount",
                type: { numeric: true },
                roles: { Values: true }
            },
            values: input.values,
            highlights: input.highlights
        });
    }

    const columns = [
        ...categories.map((c) => c.source),
        ...valueColumns.map((v) => v.source)
    ];

    return {
        categorical: {
            categories,
            values: valueColumns.length > 0 ? valueColumns : undefined
        },
        metadata: { columns, objects: input.objects }
    };
}

export function buildEmptyDataView(): any {
    return {
        categorical: { categories: [], values: [] },
        metadata: { columns: [] }
    };
}
