export type FrameworkId = 'react' | 'vue' | 'svelte' | 'solid';

export interface FrameworkGuide {
  id: FrameworkId;
  name: string;
  title: string;
  description: string;
  integrationName: string;
  stackBlitzUrl: string;
  stackBlitzEmbedUrl: string;
}

const repository = 'yukiyokotani/office-open-xml-viewer';

function stackBlitz(framework: FrameworkId, file: string): string {
  const params = new URLSearchParams({ file, startScript: 'dev' });
  return `https://stackblitz.com/github/${repository}/tree/main/examples/frameworks/${framework}?${params}`;
}

function stackBlitzEmbed(framework: FrameworkId, file: string): string {
  const params = new URLSearchParams({
    embed: '1',
    file,
    view: 'editor',
    showSidebar: '1',
    hideNavigation: '1',
  });
  return `https://stackblitz.com/github/${repository}/tree/main/examples/frameworks/${framework}?${params}`;
}

export const frameworkGuides: FrameworkGuide[] = [
  {
    id: 'react',
    name: 'React',
    title: 'How to render Office files in the browser with React | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in a React and TypeScript application with a reusable custom hook, Canvas viewer controls, and correct effect cleanup.',
    integrationName: 'useOfficeViewer',
    stackBlitzUrl: stackBlitz('react', 'src/useOfficeViewer.tsx'),
    stackBlitzEmbedUrl: stackBlitzEmbed('react', 'src/useOfficeViewer.tsx'),
  },
  {
    id: 'vue',
    name: 'Vue',
    title: 'How to render Office files in the browser with Vue | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in Vue 3 with TypeScript, a reusable Composition API composable, template refs, viewer controls, and automatic cleanup.',
    integrationName: 'useOfficeViewer',
    stackBlitzUrl: stackBlitz('vue', 'src/useOfficeViewer.ts'),
    stackBlitzEmbedUrl: stackBlitzEmbed('vue', 'src/useOfficeViewer.ts'),
  },
  {
    id: 'svelte',
    name: 'Svelte',
    title: 'How to render Office files in the browser with Svelte | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in Svelte with TypeScript, a reusable action, readable status stores, viewer controls, and reliable teardown.',
    integrationName: 'createOfficeViewer',
    stackBlitzUrl: stackBlitz('svelte', 'src/createOfficeViewer.ts'),
    stackBlitzEmbedUrl: stackBlitzEmbed('svelte', 'src/createOfficeViewer.ts'),
  },
  {
    id: 'solid',
    name: 'Solid',
    title: 'How to render Office files in the browser with Solid | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in Solid with TypeScript, a reusable reactive primitive, signal-based status, viewer controls, and cleanup.',
    integrationName: 'createOfficeViewer',
    stackBlitzUrl: stackBlitz('solid', 'src/createOfficeViewer.ts'),
    stackBlitzEmbedUrl: stackBlitzEmbed('solid', 'src/createOfficeViewer.ts'),
  },
];

export function getFrameworkGuide(id: FrameworkId): FrameworkGuide {
  const guide = frameworkGuides.find((candidate) => candidate.id === id);
  if (!guide) throw new Error(`Unknown framework guide: ${id}`);
  return guide;
}
