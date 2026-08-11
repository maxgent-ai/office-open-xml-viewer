type ReflectionCanvas = HTMLCanvasElement | OffscreenCanvas;
type ReflectionContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

export interface ReflectionBlurBounds {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface ReflectionBlurBand {
  y: number;
  h: number;
  radius: number;
}

const TARGET_RADIUS_STEP_PX = 0.5;

/**
 * PowerPoint keeps a floor reflection sharp where it meets the source and
 * progressively increases the authored blur toward the far edge. ECMA-376
 * §20.1.8.50 defines the authored blur radius but does not prescribe a Canvas
 * rasterisation algorithm. Treat that value as the far-edge radius and
 * approximate the observed Office result with narrow monotonic bands. The
 * source is painted once; only bitmap copies are repeated, keeping glyph
 * shaping out of the band loop.
 */
function reflectionBlurBands(
  bounds: ReflectionBlurBounds,
  maxBlur: number,
): ReflectionBlurBand[] {
  if (!(maxBlur > 0) || !(bounds.h > 0)) {
    return [{ y: bounds.y, h: Math.max(0, bounds.h), radius: 0 }];
  }

  // A half-device-pixel radius step is already below the visible filter change
  // for text at normal display scales. Finer sampling only repeats full bitmap
  // copies without improving the edge. Bound the pass count because one slide
  // may contain many reflected runs.
  const count = Math.max(
    4,
    Math.min(24, Math.ceil(maxBlur / TARGET_RADIUS_STEP_PX) + 1),
  );
  const bottom = bounds.y + bounds.h;
  const bands: ReflectionBlurBand[] = [];
  for (let index = 0; index < count; index++) {
    // Quantise the quadratic radius curve by equal radius increments, then
    // derive non-uniform spatial cells around those samples. Keeping radius
    // deltas uniform avoids the large far-edge jumps (and visible horizontal
    // seams) produced by evaluating distance² over uniformly-spaced bands.
    const sample = Math.sqrt(index / (count - 1));
    const previous = index === 0
      ? 0
      : Math.sqrt((index - 1) / (count - 1));
    const next = index === count - 1
      ? 1
      : Math.sqrt((index + 1) / (count - 1));
    const near = index === 0 ? 0 : (previous + sample) / 2;
    const far = index === count - 1 ? 1 : (sample + next) / 2;
    const y = bottom - far * bounds.h;
    bands.push({
      y,
      h: (far - near) * bounds.h,
      // The first band touches the source and must remain sharp. The last band
      // reaches the authored blur radius at the far edge. PowerPoint's floor
      // reflection keeps the near field legible longer than a linear radius
      // ramp; a quadratic ramp matches that observed depth-of-field falloff.
      radius: maxBlur * index / (count - 1),
    });
  }
  return bands;
}

/** Paint a sharp reflection source with blur that grows away from its bottom edge. */
export function paintDistanceAwareReflectionBlur(
  target: ReflectionContext,
  source: ReflectionCanvas,
  bounds: ReflectionBlurBounds,
  maxBlur: number,
  canvasWidth: number,
): void {
  for (const band of reflectionBlurBands(bounds, maxBlur)) {
    target.save();
    target.beginPath();
    target.rect(0, band.y, canvasWidth, band.h);
    target.clip();
    target.filter = band.radius > 0 ? `blur(${band.radius}px)` : 'none';
    target.drawImage(source as CanvasImageSource, 0, 0);
    target.restore();
  }
}
