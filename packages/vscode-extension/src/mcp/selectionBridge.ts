import { randomBytes, timingSafeEqual } from 'node:crypto';
import { createServer, type Server } from 'node:http';
import type { SelectionContextRegistry } from '../selectionContextRegistry';

export const ACTIVE_CONTEXT_BRIDGE_PORT_ENV = 'OOXML_ACTIVE_CONTEXT_BRIDGE_PORT';
export const ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV = 'OOXML_ACTIVE_CONTEXT_BRIDGE_TOKEN';

const LOOPBACK_HOST = '127.0.0.1';
const ACTIVE_CONTEXT_PATH = '/v1/context';

function sendJson(
  response: import('node:http').ServerResponse,
  status: number,
  value: unknown,
): void {
  const body = JSON.stringify(value);
  response.writeHead(status, {
    'cache-control': 'no-store',
    'content-length': Buffer.byteLength(body),
    'content-type': 'application/json; charset=utf-8',
    'x-content-type-options': 'nosniff',
  });
  response.end(body);
}

function tokenMatches(header: string | undefined, token: string): boolean {
  if (!header) return false;
  const supplied = Buffer.from(header, 'utf8');
  const expected = Buffer.from(`Bearer ${token}`, 'utf8');
  return supplied.length === expected.length && timingSafeEqual(supplied, expected);
}

/**
 * Process-local, authenticated bridge from the VS Code extension host to the
 * stdio MCP child. It never binds a non-loopback interface or persists content.
 */
export class ActiveContextBridge {
  private server: Server | null = null;
  private environmentValue: Readonly<Record<string, string>> | null = null;
  private disposed = false;
  private operationTail: Promise<void> = Promise.resolve();

  constructor(private readonly registry: SelectionContextRegistry) {}

  environment(): Promise<Readonly<Record<string, string>>> {
    if (this.disposed) return Promise.reject(new Error('Selection bridge is disposed.'));
    return this.enqueue(async () => {
      if (this.disposed) throw new Error('Selection bridge is disposed.');
      if (this.environmentValue) return this.environmentValue;
      return await this.start();
    });
  }

  private enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.operationTail.then(operation, operation);
    this.operationTail = result.then(() => undefined, () => undefined);
    return result;
  }

  private start(): Promise<Readonly<Record<string, string>>> {
    return new Promise((resolve, reject) => {
      let starting = true;
      const token = randomBytes(32).toString('hex');
      const server = createServer((request, response) => {
        if (request.url !== ACTIVE_CONTEXT_PATH) {
          sendJson(response, 404, { error: 'not_found' });
          return;
        }
        if (request.method !== 'GET') {
          response.setHeader('allow', 'GET');
          sendJson(response, 405, { error: 'method_not_allowed' });
          return;
        }
        if (!tokenMatches(request.headers.authorization, token)) {
          sendJson(response, 401, { error: 'unauthorized' });
          return;
        }
        sendJson(response, 200, {
          schemaVersion: 1,
          available: true,
          context: this.registry.getActiveContext(),
        });
      });
      server.maxHeadersCount = 16;
      server.headersTimeout = 2_000;
      server.requestTimeout = 2_000;
      server.keepAliveTimeout = 1_000;
      server.on('clientError', (_error, socket) => socket.destroy());
      server.on('error', (error) => {
        if (starting) {
          starting = false;
          reject(error);
          return;
        }
        this.invalidateServer(server);
      });
      server.listen(0, LOOPBACK_HOST, () => {
        if (!starting) return;
        starting = false;
        if (this.disposed) {
          server.close();
          reject(new Error('Selection bridge is disposed.'));
          return;
        }
        const address = server.address();
        if (!address || typeof address === 'string') {
          server.close();
          reject(new Error('Selection bridge did not receive a TCP port.'));
          return;
        }
        this.server = server;
        this.environmentValue = Object.freeze({
          [ACTIVE_CONTEXT_BRIDGE_PORT_ENV]: String(address.port),
          [ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]: token,
        });
        server.unref();
        resolve(this.environmentValue);
      });
    });
  }

  private invalidateServer(server: Server): void {
    void this.enqueue(async () => {
      if (this.server !== server) return;
      try {
        await this.stopCurrentServer();
      } catch {
        // The server has already failed. State and credentials were cleared by
        // stopCurrentServer; a later environment() call can create a new endpoint.
      }
    }).catch(() => undefined);
  }

  /** Stop serving selection context while keeping the bridge restartable. */
  async stop(): Promise<void> {
    await this.enqueue(async () => this.stopCurrentServer());
  }

  private async stopCurrentServer(): Promise<void> {
    const server = this.server;
    this.server = null;
    this.environmentValue = null;
    if (!server) return;
    await new Promise<void>((resolve, reject) => {
      server.close((error) => error ? reject(error) : resolve());
      // Do not let a slow or idle loopback client keep disabled selection data reachable.
      server.closeAllConnections();
    });
  }

  async close(): Promise<void> {
    if (this.disposed) {
      await this.operationTail;
      return;
    }
    this.disposed = true;
    await this.enqueue(async () => this.stopCurrentServer());
  }

  dispose(): void {
    void this.close().catch(() => undefined);
  }
}
