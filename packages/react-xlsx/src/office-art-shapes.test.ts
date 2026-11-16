import assert from "node:assert/strict";
import test from "node:test";
import { gunzipSync } from "node:zlib";
import { initSync, renderShape } from "./office-art/generated/office_art.js";
import payload from "./office-art/generated/payload.js";

initSync({ module: new WebAssembly.Module(gunzipSync(Buffer.from(payload, "base64"))) });

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const black = () => JSON.stringify({ color: "#000000", opacity: 1 });

function render(properties: string, style: string, width: number, height: number, styles: (kind: string, index: number, xml: string) => string | undefined) {
  const output = renderShape(properties, style, width, height, black, styles);
  assert.ok(output, "the shape rendered");
  return JSON.parse(output) as { svg: string; paintBounds: { xPx: number; yPx: number; widthPx: number; heightPx: number } };
}

test("a shape turned a quarter turn is drawn at its own size, then rotated", () => {
  // A tall, narrow "delay" shape (flat side and half-ellipse) turned 90 degrees.
  const properties = `<xdr:spPr ${NS}><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="952500" cy="3810000"/></a:xfrm><a:prstGeom prst="flowChartDelay"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr>`;
  const { svg, paintBounds } = render(properties, "", 100, 400, () => undefined);

  assert.match(svg, /<path d="M0 0 L50 0 A50 200 0 0 1 50 400 L0 400 Z/);
  assert.ok(paintBounds.widthPx > paintBounds.heightPx, "the rotated shape is wide");
  assert.ok(Math.abs(paintBounds.xPx - -150) <= 1 && Math.abs(paintBounds.yPx - 150) <= 1, "rotated about its centre");
});

test("a shape's outline comes from the theme line style its lnRef names", () => {
  const properties = `<xdr:spPr ${NS}><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr>`;
  const style = `<xdr:style ${NS}><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="dk1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="dk1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>`;
  const requests: Array<[string, number]> = [];
  const { svg } = render(properties, style, 100, 0, (kind, index, xml) => {
    requests.push([kind, index]);
    assert.match(xml, /<a:lnRef idx="1"><a:schemeClr val="dk1"\/><\/a:lnRef>/);
    return `<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`;
  });

  assert.deepEqual(requests, [["line", 1]]);
  // 0.5pt is two-thirds of a pixel.
  assert.match(svg, /stroke-width="0\.666/);
});
