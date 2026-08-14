import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartexRegionMap } from '../types/chart.js';
import {
  projectRegionMapPoint,
  regionMapColorScale,
  renderRegionMapChart,
  resolveRegionMapFeature,
} from './region-map-renderer.js';

interface RecordedMapContext {
  ctx: CanvasRenderingContext2D;
  fills: string[];
  texts: string[];
  nonFiniteArguments: number[];
  beginPathCount: number;
}

function recordingContext(): RecordedMapContext {
  const fills: string[] = [];
  const texts: string[] = [];
  const nonFiniteArguments: number[] = [];
  let beginPathCount = 0;
  const state: Record<string, unknown> = { fillStyle: '#000000', strokeStyle: '#000000' };
  const gradient = { addColorStop() {} };
  const ctx = new Proxy({} as CanvasRenderingContext2D, {
    get(_target, property) {
      if (property in state) return state[String(property)];
      if (property === 'measureText') return (value: string) => ({ width: [...String(value)].length * 6 });
      if (property === 'createLinearGradient') return () => gradient;
      if (property === 'fill') return () => fills.push(String(state.fillStyle));
      if (property === 'fillRect') return (...args: number[]) => {
        for (const value of args) if (!Number.isFinite(value)) nonFiniteArguments.push(value);
        fills.push(String(state.fillStyle));
      };
      if (property === 'fillText') return (value: string, ...args: number[]) => {
        for (const number of args) if (!Number.isFinite(number)) nonFiniteArguments.push(number);
        texts.push(String(value));
      };
      if (property === 'beginPath') return () => { beginPathCount++; };
      return (...args: unknown[]) => {
        for (const value of args) if (typeof value === 'number' && !Number.isFinite(value)) nonFiniteArguments.push(value);
      };
    },
    set(_target, property, value) {
      state[String(property)] = value;
      return true;
    },
  });
  return {
    ctx,
    fills,
    texts,
    nonFiniteArguments,
    get beginPathCount() { return beginPathCount; },
  };
}

function chart(regionMap: ChartexRegionMap): ChartModel {
  return {
    chartType: 'regionMap',
    categories: regionMap.rows.map((row) => row.label),
    series: [{ name: 'Sales', values: regionMap.rows.map((row) => row.value ?? null) }],
    showLegend: true,
    chartexRegionMap: regionMap,
  } as ChartModel;
}

describe('offline ChartEx Region Map renderer', () => {
  it('resolves ISO identities and the names used by the country boundary corpus', () => {
    const labels = [
      'United States', 'Canada', 'Mexico', 'Brazil', 'United Kingdom', 'Germany',
      'France', 'China', 'Japan', 'Australia', 'India',
    ];
    for (const label of labels) expect(resolveRegionMapFeature(label)?.n, label).toBe(label);
    expect(resolveRegionMapFeature('CN')?.n).toBe('China');
    expect(resolveRegionMapFeature('ignored', 'USA')?.n).toBe('United States');
    expect(resolveRegionMapFeature('Czechia')?.n).toBe('Czech Republic');
    expect(resolveRegionMapFeature('South Korea')?.n).toBe('Republic of Korea');
    expect(resolveRegionMapFeature('Atlantis')).toBeUndefined();
  });

  it('projects every supported MS-ODRAWXML projection to finite, distinct coordinates', () => {
    const point: readonly [number, number] = [139.69, 35.68];
    const coordinates = ['mercator', 'miller', 'robinson', 'albers'].map((projection) =>
      projectRegionMapPoint(point, projection));
    for (const coordinate of coordinates) {
      expect(Number.isFinite(coordinate.x)).toBe(true);
      expect(Number.isFinite(coordinate.y)).toBe(true);
    }
    expect(new Set(coordinates.map(({ x, y }) => `${x.toFixed(6)},${y.toFixed(6)}`)).size).toBe(4);
    expect(projectRegionMapPoint(point, 'unknown')).toEqual(projectRegionMapPoint(point, 'robinson'));
    expect(projectRegionMapPoint(point, undefined)).toEqual(projectRegionMapPoint(point, 'robinson'));
  });

  it('keeps two- and three-stop scales finite at equal and opposite finite extremes', () => {
    const extreme = regionMapColorScale(
      [-Number.MAX_VALUE, Number.MAX_VALUE],
      null,
      '156082',
    );
    expect(extreme.color(-Number.MAX_VALUE)).toBe('#C1E5F5');
    expect(extreme.color(Number.MAX_VALUE)).toBe('#104862');
    const alternateTheme = regionMapColorScale([0, 1], null, 'C00000');
    expect(alternateTheme.minColor).toBe('#FFBFBF');
    expect(alternateTheme.maxColor).toBe('#900000');
    const equal = regionMapColorScale([7, 7]);
    expect(equal.min).toBe(7);
    expect(equal.max).toBe(7);
    expect(equal.color(7)).toMatch(/^#[0-9A-F]{6}$/);
    const three = regionMapColorScale([0, 100], {
      stopCount: 3,
      minColor: '000000', midColor: 'FF0000', maxColor: 'FFFFFF',
      minPosition: { kind: 'number', value: 0 },
      midPosition: { kind: 'percent', value: 50 },
      maxPosition: { kind: 'number', value: 100 },
    });
    expect(three.color(50)).toBe('#FF0000');
    const defaultTwo = regionMapColorScale([0, 100], {
      minColor: '000000', midColor: 'FF0000', maxColor: 'FFFFFF',
    });
    expect(defaultTwo.midColor).toBeUndefined();
    expect(defaultTwo.color(50)).toBe('#808080');
  });

  it('draws fixed offline geometry and differentiates resolved from neutral countries', () => {
    const rec = recordingContext();
    const handled = renderRegionMapChart(rec.ctx, chart({
      rows: [
        { label: 'United States', value: 850 },
        { label: 'Canada', value: 420 },
        { label: 'Atlantis', value: 999 },
      ],
      regionLabelLayout: 'showAll',
      geography: { projectionType: 'robinson', cachePresent: false },
    }), { x: 0, y: 0, w: 640, h: 360 }, 4 / 3);
    expect(handled).toBe(true);
    expect(rec.beginPathCount).toBe(177);
    expect(rec.fills).toContain('#E0E0E0');
    expect(rec.fills.some((fill) => /^#[0-9A-F]{6}$/.test(fill) && fill !== '#E0E0E0')).toBe(true);
    expect(rec.texts).toContain('United States');
    expect(rec.texts).not.toContain('Atlantis');
    expect(rec.nonFiniteArguments).toEqual([]);
  });

  it('derives the rendered automatic ramp from the workbook accent instead of fixed blue', () => {
    const rec = recordingContext();
    const model = chart({
      rows: [
        { label: 'Canada', value: 0 },
        { label: 'United States', value: 1 },
      ],
      geography: { projectionType: 'robinson', cachePresent: false },
    });
    model.chartexAccents = ['C00000'];

    renderRegionMapChart(rec.ctx, model, { x: 0, y: 0, w: 640, h: 360 }, 1);

    expect(rec.fills).toContain('#FFBFBF');
    expect(rec.fills).toContain('#900000');
    expect(rec.fills).not.toContain('#C1E5F5');
    expect(rec.fills).not.toContain('#104862');
  });

  it('rejects oversized public-model rows before allocating paths or resolving geometry', () => {
    const rec = recordingContext();
    const rows = Array.from({ length: 10_001 }, (_, index) => ({ label: `Country ${index}`, value: index }));
    expect(renderRegionMapChart(rec.ctx, chart({ rows }), { x: 0, y: 0, w: 300, h: 180 }, 1)).toBe(true);
    expect(rec.beginPathCount).toBe(0);
    expect(rec.texts).toEqual(['(chart values exceed rendering limit)']);
  });

  it('fails closed for sub-country mapping levels that the country asset cannot represent', () => {
    const rec = recordingContext();
    expect(renderRegionMapChart(rec.ctx, chart({
      rows: [{ label: 'Washington', value: 1 }],
      geography: { viewedRegionType: 'state', cachePresent: false },
    }), { x: 0, y: 0, w: 300, h: 180 }, 1)).toBe(true);
    expect(rec.beginPathCount).toBe(0);
    expect(rec.texts).toEqual(['(region map detail is unavailable offline)']);
  });

  it.each(['dataOnly', 'countryRegion', 'countryRegionList'])('%s fails closed until its authored viewport is modeled', viewedRegionType => {
    const rec = recordingContext();
    expect(renderRegionMapChart(rec.ctx, chart({
      rows: [{ label: 'Canada', value: 1 }],
      geography: { viewedRegionType, cachePresent: false },
    }), { x: 0, y: 0, w: 300, h: 180 }, 1)).toBe(true);
    expect(rec.beginPathCount).toBe(0);
    expect(rec.texts).toEqual(['(region map detail is unavailable offline)']);
  });

  it('fails closed when an opaque geoCache owns localized region identities', () => {
    const rec = recordingContext();
    expect(renderRegionMapChart(rec.ctx, chart({
      rows: [{ label: '中国', value: 1 }],
      geography: { viewedRegionType: 'world', cachePresent: true, cacheProvider: 'Office' },
    }), { x: 0, y: 0, w: 300, h: 180 }, 1)).toBe(true);
    expect(rec.beginPathCount).toBe(0);
    expect(rec.texts).toEqual(['(region map cache is unavailable offline)']);
  });

  it('formats the continuous legend endpoints with the authored color-value format', () => {
    const rec = recordingContext();
    const model = chart({
      rows: [
        { label: 'Canada', value: .25 },
        { label: 'United States', value: .75 },
      ],
    });
    model.series[0].valFormatCode = '0.0%';
    renderRegionMapChart(rec.ctx, model, { x: 0, y: 0, w: 640, h: 360 }, 1);
    expect(rec.texts).toContain('25.0%');
    expect(rec.texts).toContain('75.0%');
  });
});
