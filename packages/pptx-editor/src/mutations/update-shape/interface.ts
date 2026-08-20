import type { Fill, Stroke } from '@maxgent/ooxml/pptx';

export interface ShapePatch {
  /** Horizontal position in English Metric Units (EMU). */
  readonly x?: number;
  /** Vertical position in EMU. */
  readonly y?: number;
  /** Width in EMU. */
  readonly width?: number;
  /** Height in EMU. */
  readonly height?: number;
  /** Clockwise rotation in degrees. */
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** Shape fill. Use `{ fillType: 'none' }` to remove the fill. */
  readonly fill?: Fill | null;
  /** Shape outline. Use `null` to remove the outline. */
  readonly stroke?: Stroke | null;
}
