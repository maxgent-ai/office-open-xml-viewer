import type { ShapeElement } from '@maxgent/ooxml/pptx';

import type { ElementRef } from '../domain/mutation';
import type { PptxEditorSession } from '../session/pptx-editor-session';
import type { EDITOR_SELECTION_CHANGE_REASONS } from './constants';

export interface ClientPoint {
  readonly clientX: number;
  readonly clientY: number;
}

export interface SlidePoint {
  readonly x: number;
  readonly y: number;
}

export interface ShapeHitTestOptions {
  /** Extra hit area in slide EMUs, useful for zero-width lines and connectors. */
  readonly hitSlop?: number;
}

export interface PptxEditorShapeSelection {
  readonly target: ElementRef;
  readonly slideIndex: number;
  readonly presentationElementIndex: number;
  readonly element: ShapeElement;
  /** False when the parser did not expose a stable numeric OfficeCLI shape id. */
  readonly isOfficeCliTargetable: boolean;
}

export type PptxEditorSelectionChangeReason =
  (typeof EDITOR_SELECTION_CHANGE_REASONS)[keyof typeof EDITOR_SELECTION_CHANGE_REASONS];

export interface PptxEditorSelectionSnapshot {
  readonly selection: PptxEditorShapeSelection | null;
}

export interface PptxEditorSelectionChange {
  readonly reason: PptxEditorSelectionChangeReason;
  readonly snapshot: PptxEditorSelectionSnapshot;
}

export type PptxEditorSelectionListener = (change: PptxEditorSelectionChange) => void;

export type PptxEditorSelectionListenerErrorHandler = (
  cause: unknown,
  change: PptxEditorSelectionChange,
) => void;

/** Minimal interaction surface already implemented by PptxViewer. */
export interface PptxEditorInteractionHost {
  readonly canvasElement: HTMLCanvasElement;
  readonly slideIndex: number;
}

export interface PptxEditorSelectionControllerOptions {
  readonly session: PptxEditorSession;
  readonly host: PptxEditorInteractionHost;
  /** Pointer tolerance in CSS pixels. Defaults to 4. */
  readonly hitSlopPx?: number;
  readonly onListenerError?: PptxEditorSelectionListenerErrorHandler;
}
