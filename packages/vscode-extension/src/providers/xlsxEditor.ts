import * as vscode from 'vscode';
import { BaseEditorProvider } from './baseEditor';
import { SelectionContextRegistry } from '../selectionContextRegistry';

export class XlsxEditorProvider extends BaseEditorProvider {
  static readonly viewType = 'ooxmlViewer.xlsxEditor';
  protected readonly fileType = 'xlsx' as const;

  static register(
    context: vscode.ExtensionContext,
    selectionContexts: SelectionContextRegistry,
  ): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      XlsxEditorProvider.viewType,
      new XlsxEditorProvider(context, selectionContexts),
    );
  }
}
