import { EDITOR_STORE_CHANGE_REASONS } from '../store/types';

export const EDITOR_SESSION_CHANGE_REASONS = Object.freeze({
  ...EDITOR_STORE_CHANGE_REASONS,
  HISTORY_CHANGED: 'history.changed',
} as const);
