// Opt-in offline ChartEx Region Map renderer entry point.

import {
  renderRegionMapChart,
} from '../packages/core/src/chart/region-map-renderer.js';
import type { ChartRegionMapRenderer } from '../packages/core/src/chart/region-map-contract.js';
import { registerBuiltinWorkerRenderer } from '../packages/core/src/worker/renderer-module-contract.js';

export const regionMap: ChartRegionMapRenderer = registerBuiltinWorkerRenderer({
  render: renderRegionMapChart,
}, 'regionMap');

export type { ChartRegionMapRenderer } from '../packages/core/src/chart/region-map-contract.js';
