/** Whether an xdr:twoCellAnchor saved as `editAs="oneCell"` has the complete
 * positive native extent required to ignore its `to` marker. Both dimensions
 * form one DrawingML size; a partial/malformed extent falls back as a unit so
 * culling and paint always derive the same rectangle. */
export function usesNativeOneCellExtent(anchor: {
  readonly editAs?: string;
  readonly nativeExtCx?: number;
  readonly nativeExtCy?: number;
}): boolean {
  return anchor.editAs === 'oneCell'
    && (anchor.nativeExtCx ?? 0) > 0
    && (anchor.nativeExtCy ?? 0) > 0;
}
