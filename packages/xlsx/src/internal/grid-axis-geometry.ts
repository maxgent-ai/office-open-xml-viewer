/** Sparse cumulative geometry for one worksheet row/column axis. */
export class GridAxisGeometry {
  private readonly indices: number[];
  private readonly cumulativeDelta: number[];
  private readonly customPx: number[];
  private readonly defaultPx: number;

  constructor(
    customs: Record<number, number>,
    defaultPx: number,
    toPx: (raw: number) => number,
    private readonly maxIndex: number,
    sortedPixels?: readonly { index: number; px: number }[],
  ) {
    this.defaultPx = Number.isFinite(defaultPx) && defaultPx >= 0 ? defaultPx : 0;
    this.indices = sortedPixels
      ? sortedPixels.map((entry) => entry.index)
      : Object.keys(customs)
          .map(Number)
          .filter((value) => value >= 1 && value <= maxIndex)
          .sort((a, b) => a - b);
    this.cumulativeDelta = new Array(this.indices.length);
    this.customPx = new Array(this.indices.length);
    let accumulated = 0;
    for (let index = 0; index < this.indices.length; index++) {
      const measured = sortedPixels?.[index]?.px ?? toPx(customs[this.indices[index]]);
      const px = Number.isFinite(measured) && measured >= 0 ? measured : this.defaultPx;
      this.customPx[index] = px;
      accumulated += px - this.defaultPx;
      this.cumulativeDelta[index] = accumulated;
    }
  }

  private deltaBefore(index: number): number {
    let low = 0;
    let high = this.indices.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (this.indices[middle] < index) low = middle + 1;
      else high = middle;
    }
    return low === 0 ? 0 : this.cumulativeDelta[low - 1];
  }

  offsetOf(index: number): number {
    return (index - 1) * this.defaultPx + this.deltaBefore(index);
  }

  indexAt(offset: number): { index: number; partial: number } {
    if (offset < 0) return { index: 1, partial: 0 };
    let low = 1;
    let high = this.maxIndex;
    while (low < high) {
      const middle = (low + high + 1) >> 1;
      if (this.offsetOf(middle) <= offset) low = middle;
      else high = middle - 1;
    }
    return { index: low, partial: offset - this.offsetOf(low) };
  }

  scrollableIndexAt(content: number, firstScrollable: number): number | null {
    const absoluteOffset = content + this.offsetOf(firstScrollable);
    if (absoluteOffset >= this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex)) return null;
    return this.indexAt(absoluteOffset).index;
  }

  sizeOf(index: number): number {
    return this.offsetOf(index + 1) - this.offsetOf(index);
  }

  scaled(scale: number): GridAxisGeometry {
    return new GridAxisGeometry(
      {},
      Math.round(this.defaultPx * scale),
      (value) => value,
      this.maxIndex,
      this.indices.map((index, position) => ({
        index,
        px: Math.round(this.customPx[position] * scale),
      })),
    );
  }

  countToCover(index: number, distance: number): number {
    if (index > this.maxIndex || distance <= 0) return 0;
    const target = this.offsetOf(index) + distance;
    const end = this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex);
    if (target >= end) return this.maxIndex - index + 1;
    const located = this.indexAt(target);
    return located.index - index + (located.partial > 0 ? 1 : 0);
  }

  /** Materialize only positive-width bands, jumping across zero-sized runs by
   * cumulative offset. Work is bounded by visible pixels, not sheet ordinals. */
  bandsToCover(
    startIndex: number,
    endIndex: number,
    distance = Number.POSITIVE_INFINITY,
  ): ReadonlyArray<{ index: number; size: number }> {
    const start = Math.max(1, startIndex);
    const end = Math.min(this.maxIndex, endIndex);
    if (start > end || distance <= 0) return [];
    const bands: Array<{ index: number; size: number }> = [];
    let covered = 0;
    let index = Math.max(start, this.indexAt(this.offsetOf(start)).index);
    while (index <= end && covered < distance) {
      const size = this.sizeOf(index);
      if (Number.isFinite(size) && size > 0) {
        bands.push({ index, size });
        covered += size;
      }
      if (index >= end) break;
      const nextOffset = this.offsetOf(index + 1);
      const located = this.indexAt(nextOffset).index;
      index = Math.max(index + 1, located);
    }
    return bands;
  }
}
