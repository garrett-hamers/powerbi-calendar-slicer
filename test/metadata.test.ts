import { readdirSync, readFileSync } from "node:fs";
import { relative, resolve } from "node:path";
import { describe, expect, it } from "vitest";

function readJson<T>(fileName: string): T {
    return JSON.parse(readFileSync(resolve(process.cwd(), fileName), "utf8")) as T;
}

function findTypeScriptFiles(directory: string): string[] {
    return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
        const filePath = resolve(directory, entry.name);
        if (entry.isDirectory()) {
            return findTypeScriptFiles(filePath);
        }
        return entry.isFile() && entry.name.endsWith(".ts") ? [filePath] : [];
    });
}

describe("submission metadata", () => {
    it("keeps identity stable and versions aligned", () => {
        const pbiviz = readJson<{
            version: string;
            apiVersion: string;
            visual: {
                name: string;
                displayName: string;
                guid: string;
                supportUrl: string;
                version: string;
            };
            author: { name: string; email: string };
            externalJS: unknown;
            stringResources: string[];
        }>("pbiviz.json");
        const packageJson = readJson<{ name: string; version: string; private: boolean }>(
            "package.json"
        );

        expect(pbiviz.visual.guid).toBe("calendarSlicerATLYN606CC6AF684C4BBA");
        expect(pbiviz.visual.name).toBe("atlynCalendarSlicer");
        expect(pbiviz.visual.displayName).toBe("Atlyn Calendar Slicer");
        expect(pbiviz.visual.version).toBe("1.0.0.3");
        expect(pbiviz.version).toBe("1.0.0.3");
        expect(pbiviz.apiVersion).toBe("5.11.0");
        expect(pbiviz.visual.supportUrl)
            .toBe("https://github.com/garrett-hamers/powerbi-calendar-slicer/issues");
        expect(pbiviz.author.name).toBe("Atlyn");
        expect(pbiviz.author.email).toBe("atlyn.help@gmail.com");
        // externalJS must be empty for certification.
        expect(pbiviz.externalJS).toBeNull();
        expect(pbiviz.stringResources).toEqual([
            "stringResources/en-US/resources.resjson"
        ]);

        expect(packageJson.name).toBe("calendar-slicer-visual");
        expect(packageJson.version).toBe("1.0.0.3");
        expect(packageJson.private).toBe(true);
    });

    it("declares Calendar Slicer data roles and mapping", () => {
        const capabilities = readJson<{
            supportsHighlight?: boolean;
            supportsKeyboardFocus?: boolean;
            supportsSynchronizingFilterState?: boolean;
            supportsLandingPage?: boolean;
            supportsEmptyDataView?: boolean;
            suppressDefaultTitle?: boolean;
            dataRoles: Array<{ name: string; kind: string }>;
            dataViewMappings: Array<{
                conditions: Array<Record<string, { min?: number; max?: number }>>;
                categorical: {
                    categories: {
                        for?: { in: string };
                        dataReductionAlgorithm?: { top?: { count?: number } };
                    };
                    values: { select: Array<Record<string, unknown>> };
                };
            }>;
            privileges: unknown[];
        }>("capabilities.json");

        expect(capabilities.supportsHighlight).toBe(false);
        expect(capabilities.supportsKeyboardFocus).toBe(true);
        expect(capabilities.supportsSynchronizingFilterState).toBe(true);
        expect(capabilities.supportsLandingPage).toBe(true);
        expect(capabilities.supportsEmptyDataView).toBe(true);
        expect(capabilities.suppressDefaultTitle).toBe(true);

        const roles = new Map(capabilities.dataRoles.map((r) => [r.name, r.kind]));
        expect(roles.get("Date")).toBe("Grouping");
        expect(roles.get("Values")).toBe("Measure");

        const mapping = capabilities.dataViewMappings[0];
        expect(mapping.conditions[0]).toMatchObject({
            Date: { min: 1, max: 1 },
            Values: { min: 0, max: 1 }
        });
        expect(mapping.categorical.categories.for).toEqual({ in: "Date" });
        // Data reduction must be high enough to keep multi-year daily date tables.
        expect(mapping.categorical.categories.dataReductionAlgorithm?.top?.count)
            .toBe(30000);
        expect(mapping.categorical.values.select).toContainEqual({
            bind: { to: "Values" }
        });
        expect(capabilities.privileges).toEqual([]);
    });

    it("declares the mandatory general.filter object plus persisted view state", () => {
        const capabilities = readJson<{
            objects: Record<string, { properties: Record<string, {
                type?: Record<string, unknown>;
                filterState?: boolean;
            }> }>;
        }>("capabilities.json");

        const general = capabilities.objects.general;
        expect(general).toBeDefined();
        // Filter property is required for applyJsonFilter to survive bookmarks.
        expect(general.properties.filter.type).toEqual({ filter: true });
        // Non-filter UI state is persisted as filterState so bookmarks restore it.
        expect(general.properties.visibleYear.filterState).toBe(true);
        expect(general.properties.visibleMonth.filterState).toBe(true);
        expect(general.properties.activePreset.filterState).toBe(true);
    });

    it("declares an object for every format card", () => {
        const capabilities = readJson<{
            objects: Record<string, { properties: Record<string, unknown> }>;
        }>("capabilities.json");

        for (const card of [
            "general", "calendar", "cells", "heatmap", "presets", "interaction"
        ]) {
            expect(capabilities.objects[card]).toBeDefined();
        }
        expect(capabilities.objects.calendar.properties.weekStartDay).toBeDefined();
        expect(capabilities.objects.calendar.properties.fiscalYearStartMonth).toBeDefined();
        expect(capabilities.objects.cells.properties.selectedColor).toBeDefined();
        expect(capabilities.objects.heatmap.properties.show).toBeDefined();
        expect(capabilities.objects.presets.properties.show).toBeDefined();
        expect(capabilities.objects.interaction.properties.multiSelect).toBeDefined();
    });

    it("uses a locked local certification toolchain", () => {
        const packageJson = readJson<{
            scripts: Record<string, string>;
            dependencies: Record<string, string>;
            devDependencies: Record<string, string>;
        }>("package.json");
        const gitignore = readFileSync(resolve(process.cwd(), ".gitignore"), "utf8");
        const readme = readFileSync(resolve(process.cwd(), "README.md"), "utf8");
        const license = readFileSync(resolve(process.cwd(), "LICENSE"), "utf8");

        expect(packageJson.scripts.package).toBe("pbiviz package --certification-audit");
        // The eslint script string is mandated verbatim by Microsoft certification.
        expect(packageJson.scripts.eslint).toBe("npx eslint . --ext .js,.jsx,.ts,.tsx");
        expect(packageJson.scripts.certify).toBe(
            "npm run audit && npm run eslint && npm run typecheck && npm run test && npm run package"
        );
        expect(packageJson.scripts["hash:package"]).toBe("node scripts/hash-pbiviz.mjs");
        expect(packageJson.dependencies["powerbi-visuals-api"]).toBe("5.11.0");
        expect(packageJson.devDependencies["powerbi-visuals-tools"]).toBe("7.2.1");
        expect(packageJson.devDependencies["eslint-plugin-powerbi-visuals"]).toBe("1.1.1");
        expect(packageJson.devDependencies.jszip).toBe("3.10.1");
        // No date library may be bundled — date math is hand-rolled.
        for (const banned of ["moment", "date-fns", "luxon", "dayjs"]) {
            expect(packageJson.dependencies[banned]).toBeUndefined();
            expect(packageJson.devDependencies[banned]).toBeUndefined();
        }
        expect(gitignore).toContain("dist/");
        expect(gitignore).toContain(".tmp/");
        expect(readme).toContain("calendarSlicerATLYN606CC6AF684C4BBA.1.0.0.3.pbiviz");
        expect(readme).toContain("Power_BI-API_5.11");
        expect(readme).toContain("npm run certify");
        expect(license).toContain("MIT License");
        expect(license).toContain("Copyright (c) 2026 Atlyn");
    });

    it("keeps the visual source free of certification-blocked APIs", () => {
        const stripComments = (code: string): string =>
            code
                .replace(/\/\*[\s\S]*?\*\//g, "")
                .replace(/\/\/[^\n]*/g, "");

        const sourcePaths = findTypeScriptFiles(resolve(process.cwd(), "src"))
            .map((filePath) => relative(process.cwd(), filePath).replaceAll("\\", "/"))
            .sort();
        expect(sourcePaths).toEqual(expect.arrayContaining([
            "src/visual.ts",
            "src/settings.ts"
        ]));
        const sources = sourcePaths.map((filePath) =>
            stripComments(readFileSync(resolve(process.cwd(), filePath), "utf8"))
        );

        for (const code of sources) {
            expect(code).not.toMatch(/Math\.random/);
            expect(code).not.toMatch(/\.innerHTML/);
            expect(code).not.toMatch(/\beval\s*\(/);
            expect(code).not.toMatch(/new\s+Function\s*\(/);
            expect(code).not.toMatch(/XMLHttpRequest/);
            expect(code).not.toMatch(/WebSocket/);
            expect(code).not.toMatch(/\bfetch\s*\(/);
            expect(code).not.toMatch(/import\s*\(/);
        }
    });

    it("pins the FilterAction workaround to the correct enum values", () => {
        const visualSource = readFileSync(resolve(process.cwd(), "src/visual.ts"), "utf8");
        // The ambient const enum is not inlined by esbuild, so the values are
        // pinned as strict literal-member types. Lock both the names (fleet
        // convention) and the values so the workaround can't silently rot or be
        // "simplified" back into a broken `import { FilterAction }`.
        expect(visualSource).toMatch(/MERGE_FILTER_ACTION:\s*powerbi\.FilterAction\.merge\s*=\s*0/);
        expect(visualSource).toMatch(/REMOVE_FILTER_ACTION:\s*powerbi\.FilterAction\.remove\s*=\s*1/);
        const visualCode = visualSource
            .replace(/\/\*[\s\S]*?\*\//g, "")
            .replace(/\/\/[^\n]*/g, "");
        expect(visualCode).not.toMatch(/import\s*\{[^}]*\bFilterAction\b[^}]*\}/);
    });

    it("does not contain any hosted CI/CD configuration", () => {
        const forbidden = [
            ".github/workflows",
            ".gitlab-ci.yml",
            "azure-pipelines.yml",
            ".circleci",
            "Jenkinsfile",
            ".travis.yml",
            "appveyor.yml"
        ];
        for (const entry of forbidden) {
            let present = true;
            try {
                readdirSync(resolve(process.cwd(), entry));
            } catch {
                try {
                    readFileSync(resolve(process.cwd(), entry));
                } catch {
                    present = false;
                }
            }
            expect(present, `${entry} must not exist (local-only validation)`).toBe(false);
        }
    });
});
