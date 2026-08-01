# Atlyn Calendar Slicer

A free, open-source Power BI custom visual that filters a report by a date column from a **real 7-column month grid** — not a horizontal timeline ribbon. Pick a single day, drag a date range, Ctrl-click individual days, or apply relative presets such as Month-to-Date, Year-to-Date, and Last 7 Days.

![Power BI](https://img.shields.io/badge/Power_BI-API_5.11-yellow)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.0.0.4-blue)

---

## Features

### A real calendar grid
- Seven columns, five to six week rows, with leading and trailing days from the
  adjacent months greyed for context.
- Previous / next month navigation, a **Today** jump, and a configurable
  week-start day (Sunday, Monday, or Saturday).
- Optional ISO-8601 week-number column and weekend shading.

### Produces filters, correctly
- **Single day** click applies a one-day range.
- **Click-drag** or **Shift-click** applies a contiguous range.
- **Ctrl / ⌘-click** toggles individual, non-contiguous days.
- **Relative presets** (Today, Yesterday, This Week, Last 7/14/30 Days, MTD, QTD,
  YTD, Last Month/Quarter/Year) use Power BI relative filters where those
  semantics are exact. Fiscal, to-date, and custom windows are fixed at the
  time the button is applied and remain fixed when restored from a bookmark.
- A visible **Clear** button removes the filter entirely.

Ranges are applied as a **half-open interval** `[start, nextPeriodStart)` using
`GreaterThanOrEqual` + `LessThan`, so fact rows that carry a time component (for
example `2024-03-31T14:30:00`) are never silently dropped. All date serialisation
is routed through a single timezone/DST-safe helper, unit-tested across UTC, a
DST timezone, and a half-hour-offset timezone.

### Optional heat-shading
- Bind an optional **Values** measure to shade each day by magnitude on a
  configurable low → high colour ramp, and optionally grey days that have no data.
- See **Known limits** for the slicer-sync trade-off of binding a measure.

### Accessible
- `role="grid"` / `role="gridcell"`, `aria-selected`, and a per-day label
  ("March 15, 2024").
- Full arrow-key navigation, Enter/Space to select, Delete/Backspace or the Clear
  button to clear, visible focus rings, and right-click context menus. Escape is
  intentionally left for Power BI host focus navigation.
- High-contrast support via the host colour palette, and report-theme awareness.
- Rendering lifecycle events (`renderingStarted` / `renderingFinished` /
  `renderingFailed`) are reported to the host.

### Localized
- Field-well labels, format-pane cards/slices, and messages are localized through
  the host localization manager, with English strings in
  `stringResources/en-US/resources.resjson`. Missing host keys safely fall back to
  these English values. Add a sibling locale folder to translate.

---

## Data Roles

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| **Date** | Grouping | ✅ | One concrete Date/DateTime column to filter by |
| **Values** | Measure | | Optional measure used to heat-shade days by magnitude |

Only a concrete `dateTime` Date/DateTime column is supported. Automatic numeric
levels, date hierarchies, drill, sort, and matrix inputs are intentionally
rejected because they require different semantic filter targets. Add the actual
column to the Date bucket rather than a hierarchy level.

---

## Format Pane Options

| Card | Options |
|------|---------|
| **Calendar** | Week starts on, months shown, mark today, show week numbers, fiscal-year start month |
| **Cells** | Text / header / selected / today colours, weekend shading and colour |
| **Heatmap** | Heat-shade days, low/high colours, grey days without data |
| **Presets** | Show preset buttons |
| **Interaction** | Allow multi-select |

---

## Installation

### From a Locally Built Package
1. Run `npm ci` and `npm run package`
2. In Power BI Desktop → **File → Import → Power BI Visual**
3. Select `dist/calendarSlicerATLYN606CC6AF684C4BBA.1.0.0.4.pbiviz`

### Development

```bash
# Install dependencies
npm install

# Start dev server (requires Power BI developer mode)
npm start

# Type-check, lint, and test
npm run typecheck
npm run lint
npm test

# Regenerate the icon (assets/icon.png)
npm run icon

# Package for distribution (runs the certification audit)
npm run package

# Report the timestamp-dependent outer hash and stable embedded-content hash
npm run hash:package
```

---

## Validation

There is **no hosted CI** for this repository, by policy (see [AGENTS.md](AGENTS.md)).
The single supported validation entry point runs the whole gate locally:

```bash
# audit + eslint + typecheck + tests + package
npm run certify
```

Run it from a clean `npm ci` before considering a change ready. The `.pbiviz`
outer SHA-256 changes between otherwise identical builds because the package ZIP
records build timestamps — recompute it immediately before upload, and use
`npm run hash:package` to compare the stable hash over the archive's embedded
content.

---

## Release / distribution

The packaged file and the storefront download have **two different names** — this
is expected (Word Cloud has the same split), so the rename below is a required
manual handoff step:

1. **Build.** `npm run certify` (or `npm run package`) writes
   `dist/calendarSlicerATLYN606CC6AF684C4BBA.1.0.0.4.pbiviz` — `pbiviz` always
   names the output `{guid}.{version}.pbiviz`.
2. **Rename for the storefront.** Copy it to **`atlynCalendarSlicer.pbiviz`** to
   match `DownloadFileName` in the product catalogue, and upload it to the blob
   path `visuals/calendar-slicer/1.0.0.4/atlynCalendarSlicer.pbiviz`. Renaming
   does not change the bytes, so the SHA-256 is unaffected.
3. **Record the hash.** `npm run hash:package` produces the SHA-256 for the
   Partner Center certification notes. The hash you submit must match the file in
   blob storage **byte-for-byte** — compute it against the exact archive you
   upload (recompute after any rebuild, since the outer ZIP hash is
   timestamp-dependent).

Keep the version pinned at `1.0.0.4` across `package.json`, `package-lock.json`,
`pbiviz.json`, and the
blob path until a packaged-content change warrants a coordinated bump.

---

## Testing

Automated tests run under Vitest with the happy-dom environment.

| Suite | Coverage |
|-------|----------|
| Metadata | Identity, versions, capabilities roles/objects, toolchain, no hosted CI, banned APIs |
| Package Hashing | Stable framed hash and structural-collision resistance |
| Date Math | TZ/DST-safe serialisation matrix, ISO week numbers, month-grid generation, fiscal offsets |
| Date Filters | Half-open range interval, basic multi-day filter, relative-date presets |
| Visual Integration | Landing/empty states, grid render, selection, filter application, ARIA, high contrast |

```bash
npm test
```

---

## Known limits

- **Slicer sync + heat-shading.** `supportsSynchronizingFilterState` supports only
  one bound field at a time. When the optional **Values** measure is bound for
  heat-shading, Power BI disables *Sync slicers* across pages. Heat-shading is
  therefore opt-in (off by default); leave it off if you rely on synced slicers.
- Very large date tables are reduced by the host to at most 30,000 rows before
  rendering. When the cap is reached, the visual discloses that the data may be
  incomplete and disables **Grey days without data** rather than mislabelling
  missing rows. Converting a contiguous range to Ctrl/⌘ discrete values is
  bounded at 5,000 dates and announces the limit instead of sending a huge
  filter payload.
- Today, Yesterday, This Week (when using the host's Sunday week), and Last N
  days/months/years use a host `RelativeDateFilter` and therefore roll with the
  report clock. To-date, custom-week, and fiscal presets use a fixed half-open
  range computed when applied; restoring a bookmark does not silently move that
  saved window. Apply those presets again when a new clock-relative window is
  desired.
- **Non-contiguous multi-select serialisation.** Ctrl/⌘-click selections are applied
  as a `BasicFilter ("In")` whose values are naive local wall-clock strings
  (`2024-11-09T00:00:00`, no `Z`). Range selections use the UTC-relabelled half-open
  form. The split is deliberate: an `In` filter matches by exact equality, so a
  trailing `Z` is read as UTC and converted into the model timezone, never matching
  the local-midnight DateTime stored in the column (this was the v1.0.0.1 fix). A
  single `AdvancedFilter` cannot express a non-contiguous selection — `powerbi-models`
  caps it at two conditions — so `BasicFilter` is the correct primitive here.
- **Static/export contexts.** In PowerPoint export and email subscriptions the host
  reports `allowInteractions = false`; the slicer then renders its current state
  read-only (no clicks, drag, keyboard filtering, focus rings, or hover).
- **Bookmarks.** `pbiviz package` prints a "Bookmarks" warning for this visual. It
  is a **known false positive**: the packager only detects the *SelectionManager*
  bookmark path. This is a **filter** visual, which restores from
  `options.jsonFilters` (plus `general.filter` and `filterState: true` on the
  visible-month/preset state) — the path Microsoft documents for filter visuals.
  Bookmarks work correctly; `registerOnSelectCallback` is intentionally not used.
- **Tooltips.** Not implemented in v1. Even in heat-shading mode the day number and
  cell shading convey the value; a native `tooltipService` hover is a candidate for
  a future release rather than a certification requirement.
- **Validation environments.** Automated tests cover the visual contract, but
  Microsoft Desktop, Service, and mobile host behavior (especially semantic
  filter application, context menus, touch scrolling, bookmarks, and host focus)
  still requires manual validation in those products.

---

## Tech Stack

- **Power BI Visuals API** 5.11.0
- **TypeScript** with hand-rolled, timezone-safe date math (no date library)
- Plain DOM rendering (no runtime charting dependency)
- **Vitest** + happy-dom for testing

---

## License

MIT License — free for personal and commercial use.

---

## Credits

Built by [Atlyn](https://github.com/garrett-hamers). Filter and timezone-handling
patterns follow Microsoft's open-source
[powerbi-visuals-timeline](https://github.com/microsoft/powerbi-visuals-timeline)
reference implementation.
