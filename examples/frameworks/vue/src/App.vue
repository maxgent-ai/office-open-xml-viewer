<script setup lang="ts">
import { computed, shallowRef, useTemplateRef } from 'vue';
import { useOfficeViewer, type OfficeFormat, type OfficeSource } from './useOfficeViewer';

const document = shallowRef<{ format: OfficeFormat; source: OfficeSource; name: string } | null>(null);
const target = useTemplateRef<HTMLDivElement>('target');
const viewer = useOfficeViewer({
  target,
  format: computed(() => document.value?.format ?? null),
  source: computed(() => document.value?.source ?? null),
});
const chooseFile = async (event: Event) => {
  const file = (event.currentTarget as HTMLInputElement).files?.[0];
  const format = file?.name.split('.').pop()?.toLowerCase();
  if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
  document.value = { format, source: await file.arrayBuffer(), name: file.name };
};
</script>

<template>
  <main class="app">
    <div class="toolbar">
      <button aria-label="Zoom out" @click="viewer.zoomOut">−</button>
      <button @click="viewer.fitWidth">Fit width</button>
      <button aria-label="Zoom in" @click="viewer.zoomIn">+</button>
      <button @click="viewer.reload">Reload</button>
      <input class="file-input" type="file" accept=".docx,.xlsx,.pptx" @change="chooseFile" />
      <span :class="['status', { error: viewer.error.value }]">
        {{ viewer.error.value?.message ?? (document ? `${document.name} · ${viewer.status.value}` : 'Choose an Office file') }}
      </span>
    </div>
    <div ref="target" class="stage"></div>
  </main>
</template>
