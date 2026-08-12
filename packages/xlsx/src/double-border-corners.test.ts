import { describe, expect, it } from 'vitest';
import { doubleBorderRailEndpoints } from './renderer.js';

describe('doubleBorderRailEndpoints', () => {
  it('keeps both rails full length for a standalone double underline', () => {
    expect(doubleBorderRailEndpoints(10, 90, false, false)).toEqual({
      outerStart: 10,
      outerEnd: 90,
      innerStart: 10,
      innerEnd: 90,
    });
  });

  it('extends the outer rail and trims the inner rail only at joined corners', () => {
    expect(doubleBorderRailEndpoints(10, 90, true, true)).toEqual({
      outerStart: 9,
      outerEnd: 91,
      innerStart: 11,
      innerEnd: 89,
    });
    expect(doubleBorderRailEndpoints(10, 90, true, false)).toEqual({
      outerStart: 9,
      outerEnd: 90,
      innerStart: 11,
      innerEnd: 90,
    });
  });
});
