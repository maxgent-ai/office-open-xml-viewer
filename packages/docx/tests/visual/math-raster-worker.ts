import { rasterizeMathSvg } from '@silurus/ooxml-core';
import { loadWorkerRenderers, type WorkerRendererDescriptor } from '@silurus/ooxml-core/worker';

interface Request {
  readonly mathml: string;
  readonly renderer: WorkerRendererDescriptor;
}

self.onmessage = async ({ data }: MessageEvent<Request>) => {
  try {
    const renderers = await loadWorkerRenderers({ math: data.renderer });
    if (!renderers.math) throw new Error('worker math renderer was not reconstructed');
    await renderers.math.loadMathJax();
    const output = await renderers.math.mathMLToSvg(data.mathml);
    const raster = await rasterizeMathSvg(output, '#111827');
    const canvas = new OffscreenCanvas(900, 240);
    const context = canvas.getContext('2d');
    if (!context) throw new Error('worker Canvas 2D context unavailable');
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);
    const scale = Math.min(800 / raster.widthPx, 180 / raster.heightPx, 1);
    const width = raster.widthPx * scale;
    const height = raster.heightPx * scale;
    context.drawImage(
      raster.source,
      (canvas.width - width) / 2,
      (canvas.height - height) / 2,
      width,
      height,
    );
    const bitmap = canvas.transferToImageBitmap();
    self.postMessage({ ok: true, bitmap }, { transfer: [bitmap] });
  } catch (error) {
    self.postMessage({
      ok: false,
      message: error instanceof Error ? error.message : String(error),
    });
  }
};
