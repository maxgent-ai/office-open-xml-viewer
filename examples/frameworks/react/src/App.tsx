import { useCallback, useState, type ChangeEvent } from 'react';
import { useOfficeViewer, type OfficeFormat, type OfficeSource } from './useOfficeViewer';

interface SelectedDocument { format: OfficeFormat; source: OfficeSource; name: string }

export function App() {
  const [selected, setSelected] = useState<SelectedDocument | null>(null);
  const viewer = useOfficeViewer({
    format: selected?.format ?? null,
    source: selected?.source ?? null,
  });
  const chooseFile = useCallback(async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    const format = file?.name.split('.').pop()?.toLowerCase();
    if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
    setSelected({ format, source: await file.arrayBuffer(), name: file.name });
  }, []);

  return (
    <main className="app">
      <div className="toolbar">
        <button onClick={viewer.zoomOut} aria-label="Zoom out">−</button>
        <button onClick={viewer.fitWidth}>Fit width</button>
        <button onClick={viewer.zoomIn} aria-label="Zoom in">+</button>
        <button onClick={viewer.reload}>Reload</button>
        <input className="file-input" type="file" accept=".docx,.xlsx,.pptx" onChange={chooseFile} />
        <span className={viewer.error ? 'status error' : 'status'}>
          {viewer.error?.message ?? (selected ? `${selected.name} · ${viewer.status}` : 'Choose an Office file')}
        </span>
      </div>
      {viewer.renderOfficeViewer({ className: 'stage' })}
    </main>
  );
}
