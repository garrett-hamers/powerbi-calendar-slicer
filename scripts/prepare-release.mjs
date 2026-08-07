import { execFileSync } from "node:child_process";
import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import { computePackageHashes } from "./hash-pbiviz.mjs";

const STOREFRONT_FILE = "atlynCalendarSlicer.pbiviz";
const CERTIFICATION_REF = "origin/certification";

function git(...args) {
    return execFileSync("git", args, { encoding: "utf8" }).trim();
}

function certificationState(sourceCommit) {
    try {
        const commit = git("rev-parse", CERTIFICATION_REF);
        return {
            ref: CERTIFICATION_REF,
            commit,
            matchesSourceCommit: commit === sourceCommit
        };
    } catch {
        return {
            ref: CERTIFICATION_REF,
            commit: null,
            matchesSourceCommit: false
        };
    }
}

async function main() {
    const pbiviz = JSON.parse(await readFile(resolve("pbiviz.json"), "utf8"));
    const packageFile = `${pbiviz.visual.guid}.${pbiviz.visual.version}.pbiviz`;
    const packagePath = resolve("dist", packageFile);
    const archiveBytes = await readFile(packagePath);
    const hashes = await computePackageHashes(archiveBytes);
    const releaseDirectory = resolve("dist", "release");
    const releasePath = resolve(releaseDirectory, STOREFRONT_FILE);
    const sourceCommit = git("rev-parse", "HEAD");
    const workingTreeClean = git(
        "status",
        "--porcelain",
        "--untracked-files=no"
    ) === "";
    const certification = certificationState(sourceCommit);

    await mkdir(releaseDirectory, { recursive: true });
    await copyFile(packagePath, releasePath);

    const manifest = {
        visualGuid: pbiviz.visual.guid,
        version: pbiviz.visual.version,
        sourcePackage: basename(packagePath),
        uploadFile: STOREFRONT_FILE,
        sizeBytes: archiveBytes.length,
        outerSha256: hashes.outerSha256,
        embeddedContentSha256: hashes.embeddedContentSha256,
        sourceCommit,
        workingTreeClean,
        certification,
        blockers: [
            ...(workingTreeClean
                ? []
                : ["Tracked files differ from the recorded source commit."]),
            ...(certification.matchesSourceCommit
                ? []
                : [`${CERTIFICATION_REF} does not match the recorded source commit.`])
        ]
    };
    await writeFile(
        resolve(releaseDirectory, "release-manifest.json"),
        `${JSON.stringify(manifest, null, 2)}\n`
    );
    console.log(JSON.stringify(manifest, null, 2));
}

await main();
