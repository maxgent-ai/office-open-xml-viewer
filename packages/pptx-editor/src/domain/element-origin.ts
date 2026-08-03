export const ELEMENT_ORIGINS = Object.freeze({
  MASTER: 'master',
  LAYOUT: 'layout',
  SLIDE: 'slide',
} as const);

export type ElementOrigin = (typeof ELEMENT_ORIGINS)[keyof typeof ELEMENT_ORIGINS];
