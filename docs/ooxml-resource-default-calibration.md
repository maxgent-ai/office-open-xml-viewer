# OOXML resource-default calibration

Recorded: 2026-08-05. Scope: 0.76 resource-policy convergence.

The standard admission policy is 128 MiB for one inflated archive entry,
256 MiB for distinct inflated bytes visited in one package session, and 4,096
archive entries. These
are deliberately comfortable browser-safety defaults, not a model of exact
JavaScript, WASM, renderer, GPU, tab, or process memory.

## Public-sample observations

The three redistributable demo documents were opened through the bounded Node
sessions backed by the same Rust `ResourceGovernor` and generated WASM used by
browser workers. Values below are parser counters and therefore do not depend on
Node allocator behavior.

| Format | compressed source | entries | declared inflated | largest actual entry | distinct actual total | result |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| DOCX | 458,245 B | 23 | 579,822 B | 66,859 B | 143,059 B | 6 pages |
| PPTX | 1,991,164 B | 80 | 2,190,705 B | 40,160 B | 204,408 B | 9 slides |
| XLSX | 49,747 B | 43 | 192,655 B | 18,050 B | 93,168 B | first sheet, 2 pull units |

The local regression corpus contained 36 DOCX, 8 XLSX, and 11 PPTX packages.
Its largest archive contained 185 entries; the public-sample maximum was 80.
The 4,096-entry public default therefore leaves more than 22 times headroom over
the observed maximum while remaining independently stricter than the
non-disableable 20,000-entry implementation ceiling.

The aggregate is reproducible without recording filenames or document content:

```sh
pnpm measure:archive-entries -- <corpus-root> [additional-corpus-root...]
```

This small public corpus is not claimed to represent all real Office files. It
does establish that the defaults leave orders of magnitude of headroom over the
repository's ordinary examples. The per-entry default intentionally rejects the
roughly 268 MiB inflated worksheet class reported in Issue #1102 before the full
model/serialization amplification can reach Window.

## Boundary evidence

Synthetic tests cover the properties that a benign corpus cannot:

- exact limit succeeds and limit + 1 produces a typed, poisoned session;
- forged small ZIP declarations are stopped by actual decompressor output;
- distinct entries add to the total while repeat reads do not double-charge it;
- central-directory identity, entry count, and metadata have separate hard
  ceilings;
- DOCX body, PPTX slide/dependency, and XLSX row/model/JSON stages enforce their
  format-owned hard quotas before unbounded serialization or retention;
- main, worker, and in-process Node transports carry the same usage shape.

Primary commands:

```sh
cargo test -p ooxml-common
cargo test -p docx-parser
cargo test -p pptx-parser
cargo test -p xlsx-parser
pnpm test
```

## Decision and adjustment guidance

128 MiB / 256 MiB / 4,096 entries is retained because it is simple, gives
normal documents large headroom, and turns known pathological classes into a deterministic
`OoxmlResourceLimitError`. Applications with a known corpus should collect
`debug: true` reports in representative testing, choose headroom above both
`largestInflatedEntryBytes` and `distinctInflatedBytes`, and validate on their
lowest-memory supported browser/device.

Increasing these limits accepts more OOXML input but does not reserve memory or
make OOM catchable. Lowering them provides earlier rejection but increases false
rejections. Setting a field to `null` disables only that public admission limit;
internal non-configurable safety ceilings still apply.

Decoded browser images use separate shared hard guards rather than additional
public tuning parameters: 32 megapixels / 128 MiB RGBA for one raster, 128 MiB
of decoded ownership per document cache or active render pass, and two
simultaneous decodes per document. These values are deliberately conservative
implementation ceilings. They prevent measured image amplification from being
silently omitted or left to an uncontrolled allocation, but are not added to the
archive counters and do not model canvas, GPU, decoder, or process overhead.
Browser-managed SVG/vector storage is count-bounded separately; it cannot be
reliably expressed as decoded RGBA bytes or explicitly closed by the library.
