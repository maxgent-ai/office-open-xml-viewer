import { describe, expect, it, vi } from 'vitest';
import { createLazyWasmModule } from './wasm-loader.ts';

describe('createLazyWasmModule', () => {
  it('resolves and compiles only on first use, then reuses the immutable module', () => {
    const module = {} as WebAssembly.Module;
    const resolvePath = vi.fn(() => '/parser.wasm');
    const compile = vi.fn(() => module);
    const getModule = createLazyWasmModule(resolvePath, compile);

    expect(resolvePath).not.toHaveBeenCalled();
    expect(compile).not.toHaveBeenCalled();

    expect(getModule()).toBe(module);
    expect(getModule()).toBe(module);
    expect(resolvePath).toHaveBeenCalledOnce();
    expect(compile).toHaveBeenCalledOnce();
    expect(compile).toHaveBeenCalledWith('/parser.wasm');
  });
});
