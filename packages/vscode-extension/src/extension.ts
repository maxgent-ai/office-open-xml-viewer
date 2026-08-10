import * as vscode from 'vscode';
import { DocxEditorProvider } from './providers/docxEditor';
import { XlsxEditorProvider } from './providers/xlsxEditor';
import { PptxEditorProvider } from './providers/pptxEditor';
import { refreshAllWebviews, showFindInActiveWebview } from './providers/baseEditor';
import { USE_GOOGLE_FONTS_CONFIG_ID } from './config';
import { activateMcp } from './mcp';
import { SelectionContextRegistry } from './selectionContextRegistry';

export function activate(context: vscode.ExtensionContext): void {
  const selectionContexts = new SelectionContextRegistry();
  context.subscriptions.push(
    DocxEditorProvider.register(context, selectionContexts),
    XlsxEditorProvider.register(context, selectionContexts),
    PptxEditorProvider.register(context, selectionContexts),
    vscode.commands.registerCommand('ooxmlViewer.find', showFindInActiveWebview),
  );

  // The Google Fonts opt-in is baked into the webview CSP, so changing it (or
  // gaining workspace trust, which can flip the effective value) requires
  // regenerating the HTML of every open OOXML preview rather than a soft
  // re-render. Re-setting `webview.html` reloads the bootstrap with the new flag.
  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration((e) => {
      if (e.affectsConfiguration(USE_GOOGLE_FONTS_CONFIG_ID)) {
        refreshAllWebviews();
      }
    }),
    vscode.workspace.onDidGrantWorkspaceTrust(() => {
      refreshAllWebviews();
    }),
  );

  activateMcp(context, selectionContexts);
}

export function deactivate(): void {
  // nothing to clean up
}
