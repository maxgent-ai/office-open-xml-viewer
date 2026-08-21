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
| Multiple classic plot groups | reversed bar/line, bar/area, and scatter/bubble group order; standard plus percent-stacked groups on one value axis; column→horizontal-bar and horizontal-bar→column on one primary axis pair; stock plus line; signed-decimal 32-bit axis identifiers; and 2-D plus 3-D controls | Bar/line and bar/area use Office's observed fixed family layering (line above bar; area behind bar) independent of XML order, while scatter/bubble follows group source order. Stack and percentage accumulation remain group-local; a percent-stacked group sharing a raw value axis stays in ratio space. For both observed opposite-direction bar orders, the first group owns the visible axis orientation and the later group retains its own bar/column geometry. This rule is limited to the shared-primary-axis boundary; mixed-direction secondary axes and distinct secondary category axes on line/area groups remain unresolved. Stock followed by line is painted. Although the schema types axis IDs as unsigned integers, observed Office-compatible packages can serialize the same opaque 32-bit cross-reference through a signed decimal view; those bounded IDs are retained without treating them as geometry or allocation sizes. Excel canonicalized the observed 2-D plus 3-D input to one 3-D family; the viewer likewise does not infer mixed geometry. Other unobserved family combinations also fail closed. |
| Bubble 3-D material | series true/false; point false, bare, and omitted; automatic, solid, gradient, pattern, 50%-alpha, and no-fill paints; small/medium/large and overlapping bubbles; plot-edge clipping; labels; legend keys; and negative sizes with `invertIfNegative` omitted/false/true | Current Excel applies one shape-local spherical light envelope over the authored positive fill, retaining gradients, patterns, and fill alpha; no-fill remains outline-only, labels remain above the bubble, and the series legend key carries the same material. Across the measured images the highlight center stayed within 0.42–0.44 of width and 0.33 of height. The renderer uses one bounded five-stop radial overlay at that normalized location, after fill and before outline. Current Excel used the series value for the visible material and ignored point-level `bubble3D`; the shared model instead preserves and consumes the ECMA provenance order point > series > owning group > false without producer-version guessing. Per MS-OE376 §2.1.1504(b), a visible negative bubble is inverted regardless of `invertIfNegative`: the observed automatic flat result is outline-only and the observed 3-D result uses white material, while the point outline remains independent. Host rotation is the existing outer chart-frame transform, not a second camera or Surface-lighting approximation. |
| String-backed bubble X source | one lone bubble series with four cached string X values, explicit per-point fills, and `varyColors=false` | Current Excel maps the string X values to one-based ordinal positions, reserves one empty ordinal slot at each axis end (four points => 0..5), and lists the string values as point legend entries. Numeric X sources retain ordinary numeric positioning and series-driven legends. The source grammar is retained explicitly; the renderer does not infer this mode from cached label text. |
| Cartesian 3-D camera | multiple families and view/depth boundaries | One homogeneous camera per chart. Bar/column and line/area use separately observed model-depth occupancy; every wall, axis, mesh, line, and area vertex still passes through the same camera. |
| Classic 3-D `CT_Surface` thickness | floor, side wall, and back wall at 0%, 10%, 25%, and 50%, with distinct authored face paints | ECMA-376 §21.2.2.206 thickness is applied outward in model space as a percentage of the largest plot-volume dimension. The base cuboid and all three bounded slabs are uniformly refit once and pass through the same camera. The observation establishes geometry only: material-dependent face shading, picture stacking, and extension of grid rules across the exterior slab faces are not inferred. |
| Classic 3-D `CT_Surface` picture options | Excel/PDF probes for floor, side wall, and back wall at 0% and 25% thickness; isolated `applyToFront`, `applyToSides`, and `applyToEnd` true/false controls for stretch and positive-thickness plain `stack`; normal and reversed value axes; back-wall camera controls at rotY 20/340 and rotX 20/-20 for stretch and planar `stack`; stretch with an asymmetric four-color image, positive and negative 25% left source/destination rectangles, and a visible border; planar `stack` with source aspect ratios from 1:2 through 8:1 on the back wall plus square/default/8:1 floor and side-wall controls, 48/96/192 authored DPI, and value spans 5/10/20; `stackScale` with unit 2; `pictureFormat=stretch` tile controls on all three surfaces at 0% and 25% thickness with 100%/50% scale, top-left/center alignment, nonzero offsets, x/y/xy flips, a reversed value axis, rotY 340, and positive/negative 25% source rectangles; image-without-options and options-without-image controls | On zero-thickness surfaces, stretch maps the complete source rectangle onto the projected face. A positive `srcRect` crop removes the authored source edge and maps the remaining intersection across the same full face; a negative `srcRect` outset retains the unavailable source portion as transparent destination space. A positive `fillRect` inset leaves that portion of the face empty and maps the complete source across the remaining projective rectangle; a negative `fillRect` outset expands the destination beyond the face and the existing face clip removes the excess. The same face-local rule holds independently on the measured surfaces and positive-thickness slab faces. Plain `stack` derives one repetition fraction from the projected plot-face aspect and source aspect, shares it across back wall, floor, and side wall, anchors at the lower edge, and repeats toward the upper edge; at positive thickness the same mapping applies independently to front and lateral side faces, while each end face maps one complete source because it has no repetition extent. Authored DPI and numeric value span do not change that mapping. A tiled fill first resolves the DrawingML physical tile size, alignment, offset, and alternating row/column flip in each face's local two-dimensional metric, applies the signed `srcRect` independently inside every tile, then projectively maps the completed grid to that face. The same rule holds for the measured planar and positive-thickness front, side, and end faces, and remains face-local under the measured camera and value-axis reversal. Tile repetitions, scratch dimensions, and aggregate scratch area are bounded before visible paint. The DrawingML tile/stretch choice does not permit a stretch-only `fillRect` on a tiled fill; malformed public models that combine them remain fail-closed rather than inventing a destination rule. `applyToFront=false` suppresses the back-wall picture and `applyToSides=false` suppresses the floor/side-wall picture. `stackScale` repeats by value-axis units on planar back/side walls; Excel ignores `pictureStackUnit` on floor as documented by MS-OE376 §2.1.1543(c). Reversing the value axis reverses data/axis placement but leaves the observed stretch, plain `stack`, and `stackScale` texture ordering unchanged. At 25% thickness, stretch maps the inner face to front, the lateral joining faces to sides, and the alternating end faces to end; omitted face elements remain enabled. For `stackScale`, repetition continues across front and side faces while an end face receives one complete source because it has no value-axis extent; the floor continues to ignore the stack unit. The asymmetric quadrants stay screen-upright across the measured positive/negative camera rotations and value-axis directions. |
| Classic 3-D line/area direct point formatting | a seven-point line with direct solid, dashed, and no-line dPt styles; and an Excel-converted 3-D area whose direct point fills were read back before PDF export | A 3-D line dPt owns the segment ending at that point: the preceding segment takes the point line paint/geometry, the following segment returns to the series style, and direct no-line removes only that incoming segment. Excel retained the 3-D area point fill in its object model but rendered the area body entirely with the series paint, so the viewer does not invent per-point area-face segmentation. |
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
- Each visible `bubble3D` mark or key adds exactly one five-stop material
  gradient to the chart-wide marker-paint budget; no material cache or mesh is
  allocated.
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
