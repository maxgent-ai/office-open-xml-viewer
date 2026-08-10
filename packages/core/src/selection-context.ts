/** Resource bounds for browser-native text selection context snapshots. */
export interface TextSelectionContextOptions {
  /** Maximum selected UTF-16 code units returned. Default and hard maximum 65,536. */
  readonly maxTextCharacters?: number;
  /** Maximum intersected rendered-run locators returned. Default and hard maximum 1,024. */
  readonly maxRunLocators?: number;
}

/** A browser context-menu event paired with the Viewer context at its target. */
export interface ViewerContextMenuEvent<TContext> {
  /** The original, synchronous browser event. Call `preventDefault()` here to replace the native menu. */
  readonly originalEvent: MouseEvent;
  /** Start the target lookup on first call and return the same memoized Promise thereafter. */
  getContext(): Promise<TContext | null>;
}
