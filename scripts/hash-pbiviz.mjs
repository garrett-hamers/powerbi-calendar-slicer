import { createHash } from "node:crypto";
import { readdir, readFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import JSZip from "jszip";

export async function computePackageHashes(archiveBytes) {
    const zip = await JSZip.loadAsync(archiveBytes);
    const embeddedHash = createHash("sha256");
    const entryNames = Object.keys(zip.files)
        .filter((name) => !zip.files[name].dir)
        .sort();
    appendLength(embeddedHash, entryNames.length);

    for (const entryName of entryNames) {
        const content = await zip.files[entryName].async("nodebuffer");
        const entryNameBytes = Buffer.from(entryName, "utf8");
        appendLength(embeddedHash, entryNameBytes.length);
        embeddedHash.update(entryNameBytes);
        appendLength(embeddedHash, content.length);
        embeddedHash.update(content);
    }

    return {
        outerSha256: createHash("sha256").update(archiveBytes).digest("hex").toUpperCase(),
        embeddedContentSha256: embeddedHash.digest("hex").toUpperCase(),
        entries: entryNames
    };
}

function appendLength(hash, length) {
    const frame = Buffer.alloc(8);
    frame.writeBigUInt64BE(BigInt(length));
    hash.update(frame);
}

async function findPackagePath(argument) {
    if (argument) return resolve(argument);

    const files = (await readdir(resolve("dist")))
        .filter((name) => name.endsWith(".pbiviz"))
        .sort();
    if (files.length !== 1) {
        throw new Error(`Expected one .pbiviz in dist, found ${files.length}.`);
    }
    return resolve("dist", files[0]);
}

async function main() {
    const packagePath = await findPackagePath(process.argv[2]);
    const hashes = await computePackageHashes(await readFile(packagePath));
    console.log(JSON.stringify({
        file: basename(packagePath),
        ...hashes
    }, null, 2));
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
    await main();
}
