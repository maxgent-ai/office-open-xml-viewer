import { afterEach, describe, expect, it } from 'vitest';
import { SelectionContextRegistry } from '../selectionContextRegistry';
import {
  ActiveContextBridge,
  ACTIVE_CONTEXT_BRIDGE_PORT_ENV,
  ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV,
} from './selectionBridge';

const bridges: ActiveContextBridge[] = [];

afterEach(async () => {
  await Promise.all(bridges.splice(0).map((bridge) => bridge.close()));
});

describe('ActiveContextBridge', () => {
  it('serves the active selection only to a bearer-authenticated loopback client', async () => {
    const registry = new SelectionContextRegistry();
    const handle = registry.track({ active: true }, {
      format: 'pptx',
      name: 'deck.pptx',
      path: '/tmp/deck.pptx',
    });
    handle.update({
      format: 'pptx',
      kind: 'element',
      slideIndex: 2,
      elementIndex: 4,
      origin: 'slide',
      elementType: 'shape',
      point: { x: 10, y: 20 },
      bounds: {
        x: 0,
        y: 0,
        width: 100,
        height: 200,
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      shapeId: '7',
      text: 'Explain this',
      truncated: false,
      truncationReasons: [],
      textCharacters: 12,
      maxTextCharacters: 16_384,
    });
    const bridge = new ActiveContextBridge(registry);
    bridges.push(bridge);
    const env = await bridge.environment();
    const url = `http://127.0.0.1:${env[ACTIVE_CONTEXT_BRIDGE_PORT_ENV]}/v1/context`;

    const unauthenticated = await fetch(url);
    expect(unauthenticated.status).toBe(401);
    expect(unauthenticated.headers.get('cache-control')).toBe('no-store');

    const response = await fetch(url, {
      headers: { authorization: `Bearer ${env[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]}` },
    });
    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toBe('application/json; charset=utf-8');
    expect(await response.json()).toMatchObject({
      schemaVersion: 1,
      available: true,
      context: {
        document: { path: '/tmp/deck.pptx', format: 'pptx' },
        selection: {
          format: 'pptx',
          kind: 'element',
          slideIndex: 2,
          shapeId: '7',
        },
      },
    });
  });

  it('returns an explicit null when no OOXML preview is active', async () => {
    const bridge = new ActiveContextBridge(new SelectionContextRegistry());
    bridges.push(bridge);
    const env = await bridge.environment();
    const response = await fetch(
      `http://127.0.0.1:${env[ACTIVE_CONTEXT_BRIDGE_PORT_ENV]}/v1/context`,
      { headers: { authorization: `Bearer ${env[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]}` } },
    );

    expect(await response.json()).toEqual({
      schemaVersion: 1,
      available: true,
      context: null,
    });
  });

  it('reuses one endpoint and rejects unsupported methods and paths', async () => {
    const bridge = new ActiveContextBridge(new SelectionContextRegistry());
    bridges.push(bridge);
    const first = await bridge.environment();
    const second = await bridge.environment();
    expect(second).toEqual(first);
    const base = `http://127.0.0.1:${first[ACTIVE_CONTEXT_BRIDGE_PORT_ENV]}`;
    const headers = { authorization: `Bearer ${first[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]}` };

    expect((await fetch(`${base}/v1/other`, { headers })).status).toBe(404);
    expect((await fetch(`${base}/v1/context`, { method: 'POST', headers })).status).toBe(405);
  });

  it('cannot finish starting after disposal', async () => {
    const bridge = new ActiveContextBridge(new SelectionContextRegistry());
    const starting = bridge.environment();
    await bridge.close();
    await expect(starting).rejects.toThrow('disposed');
  });

  it('stops exposure when MCP is disabled and rotates credentials on restart', async () => {
    const bridge = new ActiveContextBridge(new SelectionContextRegistry());
    bridges.push(bridge);
    const first = await bridge.environment();
    const firstUrl = `http://127.0.0.1:${first[ACTIVE_CONTEXT_BRIDGE_PORT_ENV]}/v1/context`;
    const stopping = bridge.stop();
    const restarting = bridge.environment();
    await stopping;
    await expect(fetch(firstUrl, {
      headers: { authorization: `Bearer ${first[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]}` },
    })).rejects.toThrow();

    const second = await restarting;
    expect(second[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]).not.toBe(first[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]);
    const response = await fetch(
      `http://127.0.0.1:${second[ACTIVE_CONTEXT_BRIDGE_PORT_ENV]}/v1/context`,
      { headers: { authorization: `Bearer ${second[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]}` } },
    );
    expect(response.status).toBe(200);
  });

  it('invalidates and restarts after a listening server error without throwing', async () => {
    const bridge = new ActiveContextBridge(new SelectionContextRegistry());
    bridges.push(bridge);
    const first = await bridge.environment();
    const internalServer = (bridge as unknown as {
      server: import('node:http').Server | null;
    }).server;
    expect(internalServer).not.toBeNull();

    internalServer!.emit('error', new Error('accept failed'));
    const second = await bridge.environment();

    expect(second[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]).not.toBe(first[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]);
    const response = await fetch(
      `http://127.0.0.1:${second[ACTIVE_CONTEXT_BRIDGE_PORT_ENV]}/v1/context`,
      { headers: { authorization: `Bearer ${second[ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]}` } },
    );
    expect(response.status).toBe(200);
  });
});
