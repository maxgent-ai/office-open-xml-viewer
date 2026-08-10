import { createSignal } from 'solid-js';
import { createOfficeViewer, type OfficeFormat, type OfficeSource } from './createOfficeViewer';

interface SelectedDocument { format: OfficeFormat; source: OfficeSource; name: string }

export function App() {
  const [target, setTarget] = createSignal<HTMLElement | null>(null);
  const [selected, setSelected] = createSignal<SelectedDocument | null>(null);
  const viewer = createOfficeViewer({
    target,
    format: () => selected()?.format ?? null,
    source: () => selected()?.source ?? null,
  });
  const chooseFile = async (event: Event) => {
    const file = (event.currentTarget as HTMLInputElement).files?.[0];
    const format = file?.name.split('.').pop()?.toLowerCase();
    if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
    setSelected({ format, source: await file.arrayBuffer(), name: file.name });
  };

  return (
    <main class="app">
      <div class="toolbar">
        <button onClick={viewer.zoomOut} aria-label="Zoom out">−</button>
        <button onClick={viewer.fitWidth}>Fit width</button>
        <button onClick={viewer.zoomIn} aria-label="Zoom in">+</button>
        <button onClick={viewer.reload}>Reload</button>
        <input class="file-input" type="file" accept=".docx,.xlsx,.pptx" onChange={chooseFile} />
        <span classList={{ status: true, error: Boolean(viewer.error()) }}>
          {viewer.error()?.message ?? (selected() ? `${selected()?.name} · ${viewer.status()}` : 'Choose an Office file')}
        </span>
      </div>
      <div class="stage" ref={setTarget} />
    </main>
  );
}
