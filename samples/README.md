# Offline sample report

`AtlynSample.pbip` is the text-source project for the Microsoft AppSource sample
report. Its semantic model contains only a DAX calculated calendar table, so it
has no data source, network dependency, credentials, or refresh gateway.

After packaging the visual, run:

```text
npm run sample:report
```

That copies the official packager's manifest and visual resource into the PBIP.
`npm run certify` then checks that the committed embedded bytes still match the
current package.

## Required Desktop conversion

Partner Center requires a `.pbix`, not a `.pbip`. Perform this native validation:

1. Open `AtlynSample.pbip` in Power BI Desktop.
2. Run **Home > Refresh > Schema and data**.
3. Confirm the calendar renders with heat shading and no credential prompt.
4. Use **File > Save As > Power BI report (.pbix)**.
5. Close Desktop, reopen that exact `.pbix`, and confirm the visual and data
   still render.
6. Upload that exact `.pbix` to Partner Center.

Do not commit or hand-edit the PBIX. Its model is a binary Analysis Services
backup image and cannot be produced safely by repository scripts.
