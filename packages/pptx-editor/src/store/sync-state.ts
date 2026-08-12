export const EDITOR_SYNC_STATUSES = Object.freeze({
  READY: 'ready',
  HALTED: 'halted',
} as const);

export interface ReadyEditorSyncState {
  readonly status: typeof EDITOR_SYNC_STATUSES.READY;
}

export interface HaltedEditorSyncState {
  readonly status: typeof EDITOR_SYNC_STATUSES.HALTED;
  readonly blockedByCommandId: string;
  readonly cause: unknown;
}

export type EditorSyncState = ReadyEditorSyncState | HaltedEditorSyncState;

export const READY_EDITOR_SYNC_STATE: ReadyEditorSyncState = Object.freeze({
  status: EDITOR_SYNC_STATUSES.READY,
});
