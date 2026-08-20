export interface ChartFrameRail {
  /** Distance from the outer frame edge to the rail center, in CSS pixels. */
  center: number;
  width: number;
}

/**
 * Split one DrawingML chart-frame width into the authored compound rails.
 *
 * ECMA-376 Part 1 §20.1.10.15 defines the rail ordering but not their
 * relative widths. Excel vector output establishes the bounded chart-frame
 * compatibility ratios used here: equal rails/gap for `dbl`, 1:1:3 and
 * 3:1:1 for the asymmetric pairs, and 1:1:2:1:1 for `tri`.
 */
export function chartFrameRails(totalWidth: number, compound?: string | null): ChartFrameRail[] {
  if (!Number.isFinite(totalWidth) || totalWidth <= 0) return [];
  const bands = compound === 'dbl' ? [1, 1, 1]
    : compound === 'thinThick' ? [1, 1, 3]
      : compound === 'thickThin' ? [3, 1, 1]
        : compound === 'tri' ? [1, 1, 2, 1, 1]
          : [1];
  const unit = totalWidth / bands.reduce((sum, band) => sum + band, 0);
  const rails: ChartFrameRail[] = [];
  let offset = 0;
  for (let index = 0; index < bands.length; index += 2) {
    const width = bands[index] * unit;
    rails.push({ center: offset + width / 2, width });
    offset += width + (bands[index + 1] ?? 0) * unit;
  }
  return rails;
}

function roundedRectPath(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  radius: number,
): void {
  const r = Math.max(0, Math.min(radius, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

/** Stroke a rectangular chart frame, keeping every compound rail inside it. */
export function strokeChartFrameRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  totalWidth: number,
  compound?: string | null,
  cornerRadius = 0,
): void {
  if (!(w > 0 && h > 0)) return;
  for (const rail of chartFrameRails(totalWidth, compound)) {
    const railWidth = Math.max(0, w - rail.center * 2);
    const railHeight = Math.max(0, h - rail.center * 2);
    if (!(railWidth > 0 && railHeight > 0)) continue;
    ctx.lineWidth = rail.width;
    if (cornerRadius > 0) {
      roundedRectPath(
        ctx,
        x + rail.center,
        y + rail.center,
        railWidth,
        railHeight,
        Math.max(0, cornerRadius - rail.center),
      );
      ctx.stroke();
    } else {
      ctx.strokeRect(x + rail.center, y + rail.center, railWidth, railHeight);
    }
  }
}
