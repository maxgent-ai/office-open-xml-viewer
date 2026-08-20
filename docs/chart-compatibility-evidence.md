# Chart compatibility evidence and scope

This document records the compatibility rules that remain after the local
Office-observation workbooks and render exports have been discarded. It is an
index, not an alternative implementation: the executable rules, safety bounds,
and causal tests live beside the shared parser and renderer code.

## Authority order

1. Authored OOXML properties and schema defaults are authoritative. Classic
   charts follow ECMA-376 / ISO 29500 Part 1, especially DrawingML chart markup
   in §21.2. Microsoft ChartEx behavior follows MS-ODRAWXML.
2. Host-independent parsing belongs to `packages/ooxml-common`; DOCX, XLSX, and
   PPTX provide only relationship, formula/cache, and theme resolution required
   by their packages. The resulting `ChartModel` is painted by
   `packages/core` in all three hosts.
3. Office compatibility rules are used only where the standards deliberately
   leave application layout or automatic choices unspecified. They must name
   their observed input class in code and retain a focused boundary test.
4. Unobserved semantics are not guessed. Unsupported geography/cache scopes and
   unsafe or oversized chart models fail closed with a bounded placeholder.

## Retained Office compatibility observations

The following completed local corpora were used to derive the current rules.
Their raw workbooks, scripts, PDFs, and screenshots are intentionally not part
of the repository.

| Surface | Valid Office observations | Implemented scope |
| --- | ---: | --- |
| Fully automatic linear value axes | 6,354 | Finite classic linear axes; strict 1.2 zero-pin boundary; 1/2/5 ceiling ladder; automatic minor unit is major/5. Twenty-four unstable tiny-offset outputs were not emulated. |
| Explicit min/max with omitted major unit | 297 | Classic linear, non-percent axes. Vertical and horizontal axes use separately observed density classes. Authored units always win. |
| Fractional classic date-axis units | month and year units 1.01, 1.5, 1.9, 1.99, 2.0, 2.01, 2.1, 2.5, and 3.1 with explicit bounds | ECMA-376 retains positive doubles, while MS-OE376 requires Office date-axis units to be at least one. Within the observed `1 <= value < 4` boundary, Excel advances integral `n.0` by `n` calendar units and any observed non-zero fractional part by `floor(n)^2` units. Larger fractional values, values below one, and omitted automatic intervals are not inferred; fractional day units remain elapsed-day intervals. |
| Percent-stacked automatic units | 48 | Horizontal/vertical and positive/signed percentage axes, with the observed 120 pt vertical density boundary. |
| Radar automatic units | 36 | Small, ordinary, and large spoke lengths. |
| Pie Style 2 repeated colors | point counts 1–48 | ECMA-376 Style 2 accent order plus Office-observed repeated-set luminance transforms. Point formatting and `noFill` remain authoritative. Counts above 48 use the same documented repeat-set rule but are not claimed as byte-exact Office observations. |
| Classic Style 2 line up/down bars with empty paint | one standard two-series line group covering rising, falling, and equal values | White up bars, black down bars, and a black outline. The fallback is restricted to legacy Style 2; direct `upBars`/`downBars` shape properties remain authoritative and other legacy styles are unresolved rather than guessed. |
| Classic stock decoration paint omitted from present elements | three- and four-series stock charts; omitted chart style and legacy Styles 1, 2, 10, and 48; rising, falling, equal, missing, and zero-crossing values; absent, empty, `noFill`, and direct-format controls; substituted theme dark-1 color | An absent `hiLowLines` element remains absent. For empty present drop/high-low lines, the observed omitted style and Styles 1, 2, and 10 use a 1 pt theme dark-1 line. Empty up/down bars use the same line plus linear-sRGB dark-1 tints: retained 25%/85% for Style 1 and 5%/95% for omitted Style/Styles 2 and 10. Direct paint wins, linked Chart Style fills only omitted properties, and `noFill` remains authoritative. Style 48 demonstrated a different light/no-bar result and is retained as an exclusion boundary; other legacy styles are unresolved rather than extrapolated. |
| Classic line/area drop lines and interior category-axis crossing | ordinary two-series line and area groups with a zero crossing and an explicit interior numeric crossing | The horizontal category axis, its ticks, `nextTo` labels, and each owning-group drop-line envelope share the same crossing coordinate. Office emits one envelope per category spanning the crossing and all plotted group points, not one coincident line per series. `low`/`high` labels remain attached to the plot edge. |
| Compound chart frames | chart-area, plot-area, and legend rectangles at 2 pt, 4 pt, and 6 pt; solid, preset-dash, and custom-dash lines; omitted pen alignment | Outer-to-inner rail/gap ratios are 1:1:1 (`dbl`), 1:1:3 (`thinThick`), 3:1:1 (`thickThin`), and 1:1:2:1:1 (`tri`). Office placed the observed omitted-alignment envelopes inside the frame edge. Explicit pen alignment, miter-limit geometry, and non-chart shape or connector lines were not observed and are not generalized by this rule. |
| Cartesian 3-D camera | multiple families and view/depth boundaries | One homogeneous camera per chart. Bar/column and line/area use separately observed model-depth occupancy; every wall, axis, mesh, line, and area vertex still passes through the same camera. |
| Region Map omitted world view | Office-produced global maps | Offline country-level rendering only. Omitted projection uses the observed Robinson world view; non-world view contracts and geo-cache identity data fail closed. |

Automatic choices are compatibility policy, not OOXML semantics. The comments
on `planLinearValueAxis`, `automaticPercentMajorUnit`,
`automaticRadarMajorUnit`, `planChartThreeDProjection`, and
`projectRegionMapPoint` define the exact supported domain. Family-local legacy
axis helpers that were no longer reachable were removed so there is one active
linear planner.

## Availability boundaries

- Each numeric tick layer is capped at 512 positions. Automatic plans coarsen
  before allocation; unsafe authored minor plans are skipped rather than
  truncated to one side of the axis.
- Classic Canvas chart input is capped at 10,000 expanded point slots. The
  optional 3-D renderer also applies a cumulative projected-face/stroke budget.
- Chart hierarchy input is capped by both row/segment count and depth before
  tree construction.
- Indexed point and label overrides are resolved into maps before paint loops;
  stacked 3-D primitives are bucketed by category once rather than rescanned
  per category.
- Region Map work is limited to 10,000 source rows and a fixed, checked-in
  Natural Earth 1:110m country asset. Rendering performs no network requests.

These are availability limits. They must not silently select a partial data
prefix or alter authored chart geometry.

## Optional modules and host wiring

Math, 3-D charts, and Region Maps use the same dependency-injection model.
`CoreLoadOptions` defines the optional engines, and DOCX, XLSX, and PPTX pass
them unchanged from viewer/document/workbook/presentation construction to the
shared chart painter. Their implementations live in separate package entries:

- `@silurus/ooxml/math`
- `@silurus/ooxml/three-d`
- `@silurus/ooxml/region-map`

Omitting a renderer keeps its implementation out of the ordinary synchronous
dependency closure. Worker mode identifies only the first-party renderer
objects through an internal registry and reconstructs them in its own realm;
the renderer interfaces themselves contain no worker transport metadata.
Custom function-valued renderers remain main-thread-only and use the documented
fallback in worker mode. Clean build checks verify the optional entry boundary,
the base DOCX/XLSX/PPTX static dependency closures, and the self-contained
production worker assets.

## Known, deliberately limited compatibility surfaces

- OOXML does not define automatic chart layout geometry. Classic title, legend,
  pie-label, and plot-band defaults are bounded Office compatibility policies;
  authored manual layout and explicit text/axis properties take precedence.
- Box-and-whisker's omitted major unit is a narrow ChartEx family policy derived
  from Office vector observations. It does not override an authored unit and is
  not reused by other ChartEx families.
- The 3-D mesh renderer implements the six ST_Shape values as model-space box,
  revolved, or tapered meshes. Office lighting is approximated from face normals;
  no sample-specific screen-space paint patches are permitted.
- Region Map supports deterministic offline country geometry. It does not
  geocode arbitrary localized text or reinterpret an authored non-world view.

When new Office evidence changes one of these policies, update the adjacent
implementation comment and causal boundary test in the same change. Do not
reintroduce archived observation workbooks as permanent repository fixtures.
