import * as vscode from 'vscode';
import { BaseEditorProvider } from './baseEditor';
import { SelectionContextRegistry } from '../selectionContextRegistry';

export class PptxEditorProvider extends BaseEditorProvider {
  static readonly viewType = 'ooxmlViewer.pptxEditor';
  protected readonly fileType = 'pptx' as const;

  static register(
    context: vscode.ExtensionContext,
    selectionContexts: SelectionContextRegistry,
  ): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      PptxEditorProvider.viewType,
      new PptxEditorProvider(context, selectionContexts),
    );
  }
}
