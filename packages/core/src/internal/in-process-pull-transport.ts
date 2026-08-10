import type {
  WorkerBridgeTransport,
  WorkerRequestOptions,
} from '../worker/index.js';

export type InProcessPullDispatch<Response> = (
  command: unknown,
  respond: (response: Response) => void,
) => void | Promise<void>;

/** In-realm adapter for the same correlated pull protocol used by workers. */
export class InProcessPullTransport<Response> implements WorkerBridgeTransport<Response> {
  private nextRequestId = 1;
  private terminated = false;

  constructor(
    private readonly dispatch: InProcessPullDispatch<Response>,
    private readonly terminateHost: () => void,
  ) {}

  async request(
    build: (id: number) => unknown,
    _transfer?: Transferable[],
    options?: WorkerRequestOptions,
  ): Promise<Response> {
    if (this.terminated) throw new Error('pull transport terminated');
    if (options?.signal?.aborted) {
      options.onCancel?.(this.nextRequestId, 'abort');
      const error = new Error('worker request aborted');
      error.name = 'AbortError';
      throw error;
    }
    const command = build(this.nextRequestId++);
    let response: Response | undefined;
    await this.dispatch(command, (value) => { response = value; });
    if (response === undefined) throw new Error('in-process pull host did not respond');
    return response;
  }

  forgetOrphaned(_ids: Iterable<number>): void {}

  terminate(): void {
    if (this.terminated) return;
    this.terminated = true;
    this.terminateHost();
  }
}
