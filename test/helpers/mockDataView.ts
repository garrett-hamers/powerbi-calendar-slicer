/**
 * Mock DataView builder for Atlyn Word Cloud tests.
 *
 * Mirrors the shape the visual reads: a categorical DataView with a `Category`
 * column (the words), an optional `Excludes` category column, an optional
 * `Values` measure (with optional highlights), and optional `tooltips` measures.
 */
export interface MockDataInput {
    categories: Array<string | number | boolean | Date | null | undefined>;
    values?: Array<number | null | undefined>;
    highlights?: Array<number | null>;
    sentiment?: Array<number | null | undefined>;
    excludes?: string[];
    categoryDisplayName?: string;
    valueDisplayName?: string;
    sentimentDisplayName?: string;
    objects?: Record<string, Record<string, unknown>>;
    tooltipMeasures?: Array<{
        displayName: string;
        values: Array<string | number | boolean | null | undefined>;
    }>;
}

export function buildMockDataView(input: MockDataInput): any {
    const categoryColumn = {
        source: {
            displayName: input.categoryDisplayName || "Words",
            queryName: "Table.Words",
            type: { text: true },
            roles: { Category: true }
        },
        values: input.categories
    };

    const categories: any[] = [categoryColumn];

    if (input.excludes) {
        categories.push({
            source: {
                displayName: "Exclude",
                queryName: "Table.Exclude",
                type: { text: true },
                roles: { Excludes: true }
            },
            values: input.excludes
        });
    }

    const valueColumns: any[] = [];

    if (input.values) {
        valueColumns.push({
            source: {
                displayName: input.valueDisplayName || "Weight",
                queryName: "Table.Weight",
                roles: { Values: true }
            },
            values: input.values,
            highlights: input.highlights
        });
    }

    if (input.sentiment) {
        valueColumns.push({
            source: {
                displayName: input.sentimentDisplayName || "Sentiment",
                queryName: "Table.Sentiment",
                roles: { Sentiment: true }
            },
            values: input.sentiment
        });
    }

    if (input.tooltipMeasures) {
        input.tooltipMeasures.forEach((measure, index) => {
            valueColumns.push({
                source: {
                    displayName: measure.displayName,
                    queryName: `Table.Tooltip${index + 1}`,
                    roles: { tooltips: true }
                },
                values: measure.values
            });
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
        categorical: {
            categories: [],
            values: []
        },
        metadata: { columns: [] }
    };
}
