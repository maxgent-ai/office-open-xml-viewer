// Live demos for the per-format detail pages. Each demo mirrors one Storybook
// story (Demo / ScrollView / ThumbnailGrid / MasterDetail) using the real API.
// Demos are mounted lazily (on scroll) so a page with several of them doesn't
// parse the same file many times at once.
import { PptxPresentation, PptxViewer } from '@silurus/ooxml-pptx';
import { DocxDocument, DocxViewer } from '@silurus/ooxml-docx';
import { XlsxViewer } from '@silurus/ooxml-xlsx';

export type Format = 'pptx' | 'docx' | 'xlsx';
export type DemoKind = 'demo' | 'scroll' | 'thumbnails' | 'masterdetail' | 'sheet';

const DPR = () => Math.min(typeof window !== 'undefined' ? window.devicePixelRatio : 1, 2);

// ── headless engine adapter (pptx slides / docx pages) ──────────────
type Doc = { count: number; render: (c: HTMLCanvasElement, i: number, width: number) => Promise<void> };

async function loadDoc(format: Format, url: string): Promise<Doc> {
  if (format === 'pptx') {
    const d = await PptxPresentation.load(url, { useGoogleFonts: true });
    return { count: d.slideCount, render: (c, i, width) => d.renderSlide(c, i, { width, dpr: DPR() }) };
  }
  const d = await DocxDocument.load(url, { useGoogleFonts: true });
  return { count: d.pageCount, render: (c, i, width) => d.renderPage(c, i, { width, dpr: DPR() }) };
}

// ── viewer adapter (pptx / docx built-in viewers) ───────────────────
type ViewerCtl = {
  load: (url: string) => Promise<void>;
  go: (i: number) => Promise<void>;
  next: () => Promise<void>;
  prev: () => Promise<void>;
  index: () => number;
  count: () => number;
};

function makeViewer(format: Format, canvas: HTMLCanvasElement, width: number): ViewerCtl {
  // NB: no enableTextSelection here. The demo canvases are downscaled with CSS
  // (.demo-page width:100%/height:auto) to fit the card, but the viewer's text
  // overlay is sized to the un-scaled page — leaving it on would inflate the
  // scroll area and add a big empty gap below the page. Text selection is shown
  // in the API reference instead.
  if (format === 'pptx') {
    const v = new PptxViewer(canvas, { width, useGoogleFonts: true });
    return {
      load: (u) => v.load(u), go: (i) => v.goToSlide(i), next: () => v.nextSlide(),
      prev: () => v.prevSlide(), index: () => v.slideIndex, count: () => v.slideCount,
    };
  }
  const v = new DocxViewer(canvas, { width, dpr: DPR(), useGoogleFonts: true });
  return {
    load: (u) => v.load(u), go: (i) => v.goToPage(i), next: () => v.nextPage(),
    prev: () => v.prevPage(), index: () => v.currentPage, count: () => v.pageCount,
  };
}

const UNIT = (f: Format) => (f === 'pptx' ? 'Slide' : 'Page');

// ── public entry ────────────────────────────────────────────────────
export function mountDemoInto(el: HTMLElement, kind: DemoKind, format: Format, url: string): void {
  el.innerHTML = '';
  if (format === 'xlsx') return mountSheet(el, url);
  switch (kind) {
    case 'scroll': return mountScroll(el, format, url);
    case 'thumbnails': return mountThumbnails(el, format, url);
    case 'masterdetail': return mountMasterDetail(el, format, url);
    default: return mountDemo(el, format, url);
  }
}

function status(el: HTMLElement, text: string): HTMLDivElement {
  const d = document.createElement('div');
  d.className = 'demo-status';
  d.setAttribute('role', 'status');
  d.setAttribute('aria-live', 'polite');
  const circle = document.createElement('span');
  circle.className = 'demo-progress-circle';
  circle.setAttribute('aria-hidden', 'true');
  const label = document.createElement('span');
  label.textContent = text;
  d.append(circle, label);
  el.appendChild(d);
  return d;
}

// Demo — single viewer with built-in navigation
function mountDemo(el: HTMLElement, format: Format, url: string): void {
  const bar = document.createElement('div');
  bar.className = 'demo-bar';
  const prev = button('‹');
  const next = button('›');
  const info = document.createElement('span');
  info.className = 'demo-info';
  info.textContent = 'Loading…';
  bar.append(prev, info, next);

  const stage = document.createElement('div');
  stage.className = 'demo-stage';
  const canvas = document.createElement('canvas');
  canvas.className = 'demo-page';
  canvas.hidden = true;
  stage.appendChild(canvas);
  const st = status(stage, 'Parsing document…');
  el.append(bar, stage);

  const v = makeViewer(format, canvas, 960);
  const viewerWrapper = canvas.parentElement as HTMLDivElement;
  viewerWrapper.hidden = true;
  const sync = () => {
    const n = v.count();
    info.textContent = n ? `${UNIT(format)} ${v.index() + 1} / ${n}` : 'Loading…';
    prev.disabled = v.index() <= 0;
    next.disabled = v.index() >= n - 1;
  };
  prev.addEventListener('click', () => void v.prev().then(sync));
  next.addEventListener('click', () => void v.next().then(sync));
  v.load(url).then(() => {
    canvas.hidden = false;
    viewerWrapper.hidden = false;
    st.remove();
    sync();
  }).catch((e) => {
    st.textContent = err(e);
    info.textContent = 'Failed';
  });
}

// ScrollView — every page stacked on a backdrop
function mountScroll(el: HTMLElement, format: Format, url: string): void {
  const sc = document.createElement('div');
  sc.className = 'demo-scroll';
  el.appendChild(sc);
  const st = status(sc, 'Parsing…');
  loadDoc(format, url).then(async (doc) => {
    st.remove();
    for (let i = 0; i < doc.count; i++) {
      const c = document.createElement('canvas');
      c.className = 'demo-page';
      sc.appendChild(c);
      await doc.render(c, i, 1100);
    }
  }).catch((e) => { st.textContent = err(e); });
}

// ThumbnailGrid — every page at a glance
function mountThumbnails(el: HTMLElement, format: Format, url: string): void {
  const grid = document.createElement('div');
  grid.className = 'demo-grid';
  el.appendChild(grid);
  const st = status(grid, 'Rendering thumbnails…');
  loadDoc(format, url).then(async (doc) => {
    for (let i = 0; i < doc.count; i++) {
      const cell = document.createElement('div');
      cell.className = 'demo-cell';
      const c = document.createElement('canvas');
      c.className = 'demo-page';
      const cap = document.createElement('span');
      cap.className = 'demo-cap';
      cap.textContent = `${UNIT(format)} ${i + 1}`;
      cell.append(c, cap);
      grid.appendChild(cell);
      await doc.render(c, i, 320);
    }
    st.remove();
  }).catch((e) => { st.textContent = err(e); });
}

// MasterDetail — thumbnail rail + large preview
function mountMasterDetail(el: HTMLElement, format: Format, url: string): void {
  const layout = document.createElement('div');
  layout.className = 'demo-md';
  const rail = document.createElement('div');
  rail.className = 'demo-rail';
  const detail = document.createElement('div');
  detail.className = 'demo-detail';
  const detailCanvas = document.createElement('canvas');
  detailCanvas.className = 'demo-page';
  detailCanvas.hidden = true;
  detail.appendChild(detailCanvas);
  const st = status(detail, 'Parsing document…');
  layout.append(rail, detail);
  el.appendChild(layout);

  const viewer = makeViewer(format, detailCanvas, 960);
  const detailViewerWrapper = detailCanvas.parentElement as HTMLDivElement;
  detailViewerWrapper.hidden = true;
  Promise.all([loadDoc(format, url), viewer.load(url)])
    .then(async ([doc]) => {
      detailCanvas.hidden = false;
      detailViewerWrapper.hidden = false;
      st.remove();
      const cells: HTMLDivElement[] = [];
      const select = async (i: number) => {
        cells.forEach((c, k) => c.classList.toggle('active', k === i));
        await viewer.go(i);
        // Reset scroll so the new page is shown from its top, not wherever the
        // previous page was scrolled to.
        detail.scrollTop = 0;
      };
      for (let i = 0; i < doc.count; i++) {
        const cell = document.createElement('div');
        cell.className = 'demo-rail-cell';
        const c = document.createElement('canvas');
        c.className = 'demo-page';
        cell.appendChild(c);
        cell.addEventListener('click', () => void select(i));
        rail.appendChild(cell);
        cells.push(cell);
        await doc.render(c, i, 200);
      }
      cells[0]?.classList.add('active');
    })
    .catch((e) => { st.textContent = err(e); });
}

// Excel — the full viewer (sheets + selection + zoom)
function mountSheet(el: HTMLElement, url: string): void {
  const host = document.createElement('div');
  host.className = 'demo-xlsx';
  el.appendChild(host);
  const viewer = new XlsxViewer(host, { useGoogleFonts: true, showZoomSlider: true });
  viewer.load(url).catch(() => { /* viewer surfaces its own errors */ });
}

function button(label: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.className = 'demo-btn';
  b.textContent = label;
  b.disabled = true;
  return b;
}
function err(e: unknown): string {
  return `Failed: ${e instanceof Error ? e.message : String(e)}`;
}
