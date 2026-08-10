import * as vscode from 'vscode';
import { BaseEditorProvider } from './baseEditor';
import { SelectionContextRegistry } from '../selectionContextRegistry';

export class DocxEditorProvider extends BaseEditorProvider {
  static readonly viewType = 'ooxmlViewer.docxEditor';
  protected readonly fileType = 'docx' as const;

  static register(
    context: vscode.ExtensionContext,
    selectionContexts: SelectionContextRegistry,
  ): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      DocxEditorProvider.viewType,
      new DocxEditorProvider(context, selectionContexts),
    );
  }
}
