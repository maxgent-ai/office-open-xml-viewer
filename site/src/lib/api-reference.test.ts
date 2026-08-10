import { describe, expect, it } from 'vitest';
import { apiReference } from './api-reference.js';

describe('official-site API reference', () => {
  it('documents the shared resource controls on every browser API class', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const optionNames = apiClass.options?.map(({ name }) => name) ?? [];
        expect(optionNames, apiClass.name).toEqual(expect.arrayContaining([
          'resourceLimits',
          'onResourceMetrics',
          'debug',
        ]));
      }
    }
  });

  it('keeps every semantic emphasis synchronized with its description', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        for (const item of [...(apiClass.options ?? []), ...apiClass.methods]) {
          if (item.emphasis) {
            expect(item.desc, `${apiClass.name}: ${'name' in item ? item.name : item.sig}`)
              .toContain(item.emphasis);
          }
        }
      }
    }
  });

  it('documents the Viewer error-delivery contract and typed resource failures', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes.filter(({ name }) => name.endsWith('Viewer'))) {
        const onError = apiClass.options?.find(({ name }) => name === 'onError');
        expect(onError, apiClass.name).toBeDefined();
        expect(onError?.desc, apiClass.name).toContain('load(), navigation, and other awaitable operations reject');
        expect(onError?.desc, apiClass.name).toContain('the same failure is never delivered twice');
        expect(onError?.desc, apiClass.name).toContain('OoxmlResourceLimitError');
        expect(onError?.desc, apiClass.name).toContain('OoxmlDecodedImageLimitError');
        expect(onError?.desc, apiClass.name).toContain('message text is not a stable discriminator');
        expect(onError?.detailsHref, apiClass.name).toBe('/errors#delivery');
      }
    }
  });

  it('links resource-limit options to their typed error fields', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const resourceLimits = apiClass.options?.find(({ name }) => name === 'resourceLimits');
        expect(resourceLimits?.detailsHref, apiClass.name).toBe('/errors#ooxml-resource-limit-error');
      }
    }
  });

  it('documents password on every self-loading Viewer and engine', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const password = apiClass.options?.find(({ name }) => name === 'password');
        expect(password?.type, apiClass.name).toBe('string');
        expect(password?.desc, apiClass.name).toContain('borrowed');
      }
    }
  });

  it('documents the common native context-menu handoff on every Viewer', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes.filter(({ name }) => name.endsWith('Viewer'))) {
        const option = apiClass.options?.find(({ name }) => name === 'onContextMenu');
        expect(option?.type, apiClass.name).toContain('ViewerContextMenuEvent<');
        expect(option?.desc, apiClass.name).toContain('originalEvent.preventDefault()');
        expect(option?.desc, apiClass.name).toContain('getContext()');
        expect(option?.desc, apiClass.name).toContain('native browser behavior unchanged');
      }
    }
  });
});
