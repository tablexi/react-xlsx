import type { XlsxOfficeLayout } from "./types";

type DigitWidths = Pick<XlsxOfficeLayout, "columnDigitWidthPt" | "columnDigitWidthPx">;

export const MIN_COL_WIDTH_PX = 30;

/**
 * The digit widths for the `columnWidths` option: a whole number of points
 * (the default layout) and, for "pixels", a whole number of pixels as Excel
 * for Windows and the web use. `measuredPt` is the font's widest digit at its
 * point size, unrounded.
 */
export function resolveColumnDigitWidths(measuredPt: number, columnWidths: "points" | "pixels" = "points"): DigitWidths {
  return {
    columnDigitWidthPt: Math.max(1, Math.round(measuredPt)),
    columnDigitWidthPx: columnWidths === "pixels" ? Math.max(1, Math.round(measuredPt * (4 / 3))) : undefined
  };
}

/** Pixel width of a column stored as `width` characters. */
export function resolveOfficeColumnWidthPixels(width: number, digits: DigitWidths) {
  const digitPx = digits.columnDigitWidthPx;
  if (digitPx) {
    return Math.max(0, Math.floor(((256 * width + Math.floor(128 / digitPx)) / 256) * digitPx));
  }
  const digitPt = digits.columnDigitWidthPt ?? 6;
  return Math.max(0, Math.floor(((256 * width + Math.floor(128 / digitPt)) / 256) * digitPt) * (4 / 3));
}

/** The stored width, in characters, of a column `widthPx` pixels wide. */
export function pxToOfficeColumnWidth(widthPx: number, digits: DigitWidths) {
  const digitPx = digits.columnDigitWidthPx;
  if (digitPx) {
    return (Math.round(Math.max(widthPx, MIN_COL_WIDTH_PX)) + 0.5) / digitPx - Math.floor(128 / digitPx) / 256;
  }
  const digitPt = digits.columnDigitWidthPt ?? 6;
  const points = Math.round(Math.max(widthPx, MIN_COL_WIDTH_PX) * (3 / 4));
  return (points + 0.5) / digitPt - Math.floor(128 / digitPt) / 256;
}

/** Pixel width of a column with no stored width, from the sheet's base width in characters. */
export function resolveOfficeDefaultColumnPixels(baseWidth: number, digits: DigitWidths) {
  const digitPx = digits.columnDigitWidthPx;
  if (digitPx) {
    // Excel rounds the default width up to a multiple of 8 pixels.
    return Math.ceil((baseWidth * digitPx + 5) / 8) * 8;
  }
  return (baseWidth * (digits.columnDigitWidthPt ?? 6) + 5) * (4 / 3);
}
