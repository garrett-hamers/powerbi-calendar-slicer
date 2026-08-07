import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import JSZip from "jszip";

const pbiviz = JSON.parse(await readFile(resolve("pbiviz.json"), "utf8"));
const guid = pbiviz.visual.guid;
const version = pbiviz.visual.version;
const packagePath = resolve("dist", `${guid}.${version}.pbiviz`);
const customVisualDirectory = resolve("samples", "AtlynSample.Report", "CustomVisuals", guid);
const manifestPath = resolve(customVisualDirectory, "package.json");
const resourcePath = resolve(customVisualDirectory, "resources", `${guid}.pbiviz.json`);

async function expectedFiles() {
    const archive = await JSZip.loadAsync(await readFile(packagePath));
    const manifest = archive.file("package.json");
    const resource = archive.file(`resources/${guid}.pbiviz.json`);
    if (!manifest || !resource) {
        throw new Error(`Unexpected package structure in ${packagePath}`);
    }
    return {
        manifest: await manifest.async("string"),
        resource: await resource.async("string")
    };
}

async function checkFile(path, expected) {
    const actual = await readFile(path, "utf8");
    if (actual !== expected) {
        throw new Error(
            `${path} does not match ${packagePath}; run npm run sample:report and commit it.`
        );
    }
}

const expected = await expectedFiles();
if (process.argv.includes("--check")) {
    await checkFile(manifestPath, expected.manifest);
    await checkFile(resourcePath, expected.resource);
    console.log("Sample PBIP embeds the current visual package.");
} else {
    await mkdir(resolve(customVisualDirectory, "resources"), { recursive: true });
    await writeFile(manifestPath, expected.manifest);
    await writeFile(resourcePath, expected.resource);
    console.log("Updated the sample PBIP embedded visual.");
}
