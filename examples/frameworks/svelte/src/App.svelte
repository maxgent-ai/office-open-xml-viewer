<script lang="ts">
  import { writable } from 'svelte/store';
  import { createOfficeViewer, type OfficeFormat, type OfficeSource } from './createOfficeViewer';

  type SelectedDocument = { format: OfficeFormat; source: OfficeSource; name: string };
  const selected = writable<SelectedDocument | null>(null);
  const viewer = createOfficeViewer();
  const { status, error } = viewer;
  const config = $derived({
    format: $selected?.format ?? null,
    source: $selected?.source ?? null,
  });
  const chooseFile = async (event: Event) => {
    const file = (event.currentTarget as HTMLInputElement).files?.[0];
    const format = file?.name.split('.').pop()?.toLowerCase();
    if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
    selected.set({ format, source: await file.arrayBuffer(), name: file.name });
  };
</script>

<main class="app">
  <div class="toolbar">
    <button aria-label="Zoom out" onclick={viewer.zoomOut}>−</button>
    <button onclick={viewer.fitWidth}>Fit width</button>
    <button aria-label="Zoom in" onclick={viewer.zoomIn}>+</button>
    <button onclick={viewer.reload}>Reload</button>
    <input class="file-input" type="file" accept=".docx,.xlsx,.pptx" onchange={chooseFile} />
    <span class:error={$error} class="status">
      {$error?.message ?? ($selected ? `${$selected.name} · ${$status}` : 'Choose an Office file')}
    </span>
  </div>
  <div class="stage" use:viewer.action={config}></div>
</main>
