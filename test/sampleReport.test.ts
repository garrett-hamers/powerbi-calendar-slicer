import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const root = process.cwd();
const readJson = <T>(path: string): T =>
    JSON.parse(readFileSync(resolve(root, path), "utf8")) as T;

describe("offline sample PBIP", () => {
    it("uses current PBIP/PBIR/TMDL definition contracts", () => {
        const pbip = readJson<{ version: string; artifacts: unknown[] }>(
            "samples/AtlynSample.pbip"
        );
        const pbir = readJson<{ version: string; datasetReference: unknown }>(
            "samples/AtlynSample.Report/definition.pbir"
        );
        const pbism = readJson<{ version: string }>(
            "samples/AtlynSample.SemanticModel/definition.pbism"
        );
        const reportVersion = readJson<{ version: string }>(
            "samples/AtlynSample.Report/definition/version.json"
        );

        expect(pbip.version).toBe("1.0");
        expect(pbip.artifacts).toHaveLength(1);
        expect(pbir.version).toBe("4.0");
        expect(pbir.datasetReference).toEqual({
            byPath: { path: "../AtlynSample.SemanticModel" }
        });
        expect(pbism.version).toBe("4.0");
        expect(reportVersion.version).toBe("2.0.0");
    });

    it("binds the real Date and Values roles to an offline calculated table", () => {
        const visual = readJson<{
            visual: {
                visualType: string;
                query: { queryState: Record<string, unknown> };
            };
        }>(
            "samples/AtlynSample.Report/definition/pages/calendarSamplePage/" +
            "visuals/calendarSampleVisual/visual.json"
        );
        const table = readFileSync(
            resolve(root, "samples/AtlynSample.SemanticModel/definition/tables/CalendarData.tmdl"),
            "utf8"
        );

        expect(visual.visual.visualType).toBe("calendarSlicerATLYN606CC6AF684C4BBA");
        expect(Object.keys(visual.visual.query.queryState).sort()).toEqual([
            "Date", "Values"
        ]);
        expect(table).toContain("partition CalendarData = calculated");
        expect(table).toContain("CALENDAR(DATE(2025, 1, 1), DATE(2026, 12, 31))");
        expect(table).not.toMatch(/\b(?:dataSource|expression)\b/i);
        expect(existsSync(resolve(
            root,
            "samples/AtlynSample.SemanticModel/definition/dataSources.tmdl"
        ))).toBe(false);
    });

    it("declares the embedded private custom visual resource", () => {
        const report = readJson<{
            resourcePackages: Array<{ name: string; type: string }>;
        }>("samples/AtlynSample.Report/definition/report.json");
        expect(report.resourcePackages).toContainEqual(expect.objectContaining({
            name: "calendarSlicerATLYN606CC6AF684C4BBA",
            type: "CustomVisual"
        }));
        expect(existsSync(resolve(
            root,
            "samples/AtlynSample.Report/CustomVisuals/" +
            "calendarSlicerATLYN606CC6AF684C4BBA/package.json"
        ))).toBe(true);
        expect(existsSync(resolve(
            root,
            "samples/AtlynSample.Report/CustomVisuals/" +
            "calendarSlicerATLYN606CC6AF684C4BBA/resources/" +
            "calendarSlicerATLYN606CC6AF684C4BBA.pbiviz.json"
        ))).toBe(true);
    });

    it("includes in-report hints for selection, tooltips, and context menus", () => {
        const pages = readJson<{ pageOrder: string[] }>(
            "samples/AtlynSample.Report/definition/pages/pages.json"
        );
        expect(pages.pageOrder).toEqual(["calendarSamplePage", "hintsTipsPage"]);

        const hintRoot = resolve(
            root,
            "samples/AtlynSample.Report/definition/pages/hintsTipsPage/visuals"
        );
        const hintFiles = [
            "hintsTitle",
            "hintsSelection",
            "hintsHostFeatures",
            "hintsDataShape"
        ].map((name) =>
            readFileSync(resolve(hintRoot, name, "visual.json"), "utf8")
        ).join("\n");
        expect(hintFiles).toContain("hints and tips");
        expect(hintFiles).toContain("Hover a day");
        expect(hintFiles).toContain("Right-click a day");
        expect(hintFiles).toContain("concrete date column");
    });
});
