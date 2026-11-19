import assert from "node:assert/strict";
import test from "node:test";
import {
  pxToOfficeColumnWidth,
  resolveColumnDigitWidths,
  resolveOfficeColumnWidthPixels,
  resolveOfficeDefaultColumnPixels
} from "./column-widths.ts";

// Calibri 11's widest digit, unrounded, in points.
const CALIBRI_11_DIGIT_PT = 5.575;
const points = resolveColumnDigitWidths(CALIBRI_11_DIGIT_PT);
const pixels = resolveColumnDigitWidths(CALIBRI_11_DIGIT_PT, "pixels");

test("digit widths round to whole points, and to whole pixels when asked", () => {
  assert.deepEqual(points, { columnDigitWidthPt: 6, columnDigitWidthPx: undefined });
  assert.deepEqual(pixels, { columnDigitWidthPt: 6, columnDigitWidthPx: 7 });
});

test("a default Calibri 11 column is 64px in pixels, as in Excel for Windows and the web", () => {
  assert.equal(resolveOfficeColumnWidthPixels(9.140625, pixels), 64);
  assert.equal(resolveOfficeDefaultColumnPixels(8, pixels), 64);
});

test("points keep the existing layout", () => {
  assert.ok(Math.abs(resolveOfficeColumnWidthPixels(9.140625, points) - 73.33) < 0.01);
  assert.ok(Math.abs(resolveOfficeDefaultColumnPixels(8, points) - 70.67) < 0.01);
});

test("a resized column's stored width gives back the same pixel width", () => {
  for (const widthPx of [30, 64, 100, 151]) {
    assert.equal(resolveOfficeColumnWidthPixels(pxToOfficeColumnWidth(widthPx, pixels), pixels), widthPx, `${widthPx}px`);
  }
  for (const widthPx of [40, 73.33333333333333, 100]) {
    const stored = resolveOfficeColumnWidthPixels(pxToOfficeColumnWidth(widthPx, points), points);
    assert.ok(Math.abs(stored - widthPx) <= 4 / 3, `${widthPx}px gave ${stored}px`);
  }
});

test("narrow columns keep the minimum width", () => {
  assert.equal(resolveOfficeColumnWidthPixels(pxToOfficeColumnWidth(10, pixels), pixels), 30);
});
