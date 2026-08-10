export const siteMetadata = {
  name: '@silurus/ooxml',
  alternateName: 'OOXML Viewer',
  origin: 'https://ooxml.silurus.dev',
  title: 'OOXML Viewer | JavaScript Office File Viewer',
  description:
    'Open-source JavaScript and TypeScript Office file viewer library for rendering DOCX, XLSX and PPTX directly in the browser with WebAssembly and Canvas. No upload or server-side conversion.',
} as const;

export const formatSeo = {
  docx: {
    title: 'JavaScript DOCX Viewer Library | @silurus/ooxml',
    description:
      'Open-source JavaScript and TypeScript DOCX viewer library for rendering Word documents directly in the browser with WebAssembly and Canvas.',
    heading: 'Render DOCX files in the browser with JavaScript.',
  },
  xlsx: {
    title: 'JavaScript XLSX Viewer Library | @silurus/ooxml',
    description:
      'Open-source JavaScript and TypeScript XLSX viewer library for rendering Excel workbooks directly in the browser with WebAssembly and Canvas.',
    heading: 'Render XLSX files in the browser with JavaScript.',
  },
  pptx: {
    title: 'JavaScript PPTX Viewer Library | @silurus/ooxml',
    description:
      'Open-source JavaScript and TypeScript PPTX viewer library for rendering PowerPoint presentations directly in the browser with WebAssembly and Canvas.',
    heading: 'Render PPTX files in the browser with JavaScript.',
  },
} as const;

const filePathPattern = /\/[^/]+\.[^/]+$/;

export function canonicalPageUrl(path: string): string {
  const absolutePath = path.startsWith('/') ? path : `/${path}`;
  const normalizedPath =
    absolutePath === '/' || absolutePath.endsWith('/') || filePathPattern.test(absolutePath)
      ? absolutePath
      : `${absolutePath}/`;
  return new URL(normalizedPath, siteMetadata.origin).href;
}
