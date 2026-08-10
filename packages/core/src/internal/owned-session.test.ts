import { describe, expect, it, vi } from 'vitest';
import { usingOwnedSession } from './owned-session.js';

describe('usingOwnedSession', () => {
  it('returns the operation result and closes once', async () => {
    const close = vi.fn(async () => undefined);
    await expect(usingOwnedSession(async () => ({ close }), async () => 42)).resolves.toBe(42);
    expect(close).toHaveBeenCalledOnce();
  });

  it('keeps the operation failure primary over cleanup failure', async () => {
    const primary = new Error('operation');
    const cleanup = new Error('cleanup');
    await expect(usingOwnedSession(
      async () => ({ close: async () => { throw cleanup; } }),
      async () => { throw primary; },
    )).rejects.toBe(primary);
  });

  it('surfaces cleanup failure after a successful operation', async () => {
    const cleanup = new Error('cleanup');
    await expect(usingOwnedSession(
      async () => ({ close: async () => { throw cleanup; } }),
      async () => 'done',
    )).rejects.toBe(cleanup);
  });
});
