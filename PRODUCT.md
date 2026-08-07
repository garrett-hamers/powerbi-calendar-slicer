# Product

## Register

product

## Users

Power BI report authors and report consumers who need to filter a report by date. Authors drop a date column into a reusable slicer visual; consumers pick a day, drag a range, toggle individual days, or apply a relative preset with mouse, touch, or keyboard, and expect the rest of the report to filter accordingly.

The Date bucket accepts one concrete Date/DateTime column only. Automatic
numeric hierarchy levels, drill, sort, and matrix inputs are outside the
visual's supported semantic target contract.

## Product Purpose

Atlyn Calendar Slicer provides a free, reviewable, open-source month-grid date slicer for Power BI. Unlike the horizontal timeline ribbon, it presents dates in a familiar seven-column calendar so consumers can reason about weekdays, weekends, and week boundaries directly, while producing correct, bookmarkable filters that respect timezones, daylight saving, and fact rows that carry a time component.

## Brand Personality

Precise, trustworthy, and familiar. The slicer should feel native to Power BI, disappear into the report workflow, and never surprise the user with a filter that silently drops or double-counts a day.

## Anti-references

Avoid off-by-one date errors, timezone drift, inclusive range endpoints that drop timestamped rows, frozen relative dates baked into bookmarks, color-only meaning, nonstandard Power BI interactions, and dense layouts that hide the current selection or navigation.

## Design Principles

- Represent dates exactly: one grid cell is one calendar day in the report's own terms.
- Produce filters through the documented filter API and keep them correct across bookmarks, refresh, and the passage of time.
- Preserve Power BI interaction conventions so the slicer behaves like a native slicer.
- Handle timezone, DST, and half-hour-offset locales through a single tested serialisation path.
- Prefer reviewable, deterministic behavior over hidden automation; no network, no randomness.
- Host-relative Last N presets roll with the report clock. Fiscal and to-date
  presets are fixed at application time so bookmarks do not silently change
  meaning.

## Accessibility & Inclusion

Target WCAG 2.1 AA-compatible interaction within the Power BI host. Support full keyboard grid navigation, meaningful ARIA labels for every day, visible focus, screen-reader context, high-contrast and report themes, non-color selection cues, and resilient layout at small sizes.

## Constraints

- Certification: `privileges` and `externalJS` empty, no `fetch`, `XMLHttpRequest`, `WebSocket`, `innerHTML`, `eval`, `new Function`, dynamic `import()`, or `Math.random` in `src/`. `npm audit` clean at moderate and above.
- No date library is bundled; date math is hand-rolled to keep the audit surface and bundle minimal.
- No hosted CI/CD. All validation runs locally via `npm run certify` (see AGENTS.md).
- The visual GUID `calendarSlicerATLYN606CC6AF684C4BBA` is fixed and never regenerated.

## Known trade-off

Binding the optional Values measure for heat-shading disables Power BI slicer synchronization across pages, because `supportsSynchronizingFilterState` supports only a single bound field. Heat-shading is opt-in and off by default; the trade-off is documented in the README and the format-pane description.

The host may reduce a date category to 30,000 rows; the visual discloses that
the received set may be incomplete. Converting a contiguous range to discrete
values is bounded at 5,000 dates. Desktop, Service, and mobile host behavior
still needs manual validation for release.
