// Opt-in 3-D chart renderer entry point: `@silurus/ooxml/three-d`.
//
// Mesh construction, the homogeneous camera and the painter are deliberately
// absent from the base DOCX/XLSX/PPTX entries. Consumers that need authored
// 3-D charts explicitly inject this synchronous renderer at viewer load time:
//
//   import { XlsxViewer } from '@silurus/ooxml/xlsx';
//   import { threeD } from '@silurus/ooxml/three-d';
//   new XlsxViewer(container, { threeD });

import {
  renderSimpleThreeDChart,
} from '../packages/core/src/chart/three-d-renderer.js';
import type { ChartThreeDRenderer } from '../packages/core/src/chart/three-d-contract.js';
import { registerBuiltinWorkerRenderer } from '../packages/core/src/worker/renderer-module-contract.js';

/** The optional model-space mesh/camera/material chart renderer. */
export const threeD: ChartThreeDRenderer = registerBuiltinWorkerRenderer({
  render: renderSimpleThreeDChart,
}, 'threeD');

export type { ChartThreeDRenderer } from '../packages/core/src/chart/three-d-contract.js';
