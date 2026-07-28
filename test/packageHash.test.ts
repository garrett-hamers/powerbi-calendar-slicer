import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { computePackageHashes } from "../scripts/hash-pbiviz.mjs";

async function createArchive(date: Date): Promise<Buffer> {
    const zip = new JSZip();
    zip.file("package.json", "{\"version\":\"1.0.0.0\"}", { date });
    zip.file("resources/visual.json", "{\"guid\":\"stable\"}", { date });
    return zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE"
    });
}

describe("PBIVIZ content hashing", () => {
    it("stays stable when only ZIP timestamps change", async () => {
        const first = await computePackageHashes(
            await createArchive(new Date("2026-01-01T00:00:00Z"))
        );
        const second = await computePackageHashes(
            await createArchive(new Date("2026-07-01T00:00:00Z"))
        );

        expect(first.outerSha256).not.toBe(second.outerSha256);
        expect(first.embeddedContentSha256).toBe(second.embeddedContentSha256);
    });

    it("distinguishes archives whose unframed name and content bytes collide", async () => {
        const split = new JSZip();
        split.file("a", "X");
        split.file("b", "Y");
        const combined = new JSZip();
        combined.file("a", Buffer.from("Xb\0Y"));

        const splitHashes = await computePackageHashes(
            await split.generateAsync({ type: "nodebuffer" })
        );
        const combinedHashes = await computePackageHashes(
            await combined.generateAsync({ type: "nodebuffer" })
        );

        expect(splitHashes.embeddedContentSha256)
            .not.toBe(combinedHashes.embeddedContentSha256);
    });

    it("reports sorted entry names alongside both hashes", async () => {
        const hashes = await computePackageHashes(
            await createArchive(new Date("2026-01-01T00:00:00Z"))
        );
        expect(hashes.entries).toEqual(["package.json", "resources/visual.json"]);
        expect(hashes.outerSha256).toMatch(/^[0-9A-F]{64}$/);
        expect(hashes.embeddedContentSha256).toMatch(/^[0-9A-F]{64}$/);
    });
});
