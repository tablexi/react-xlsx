import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { strToU8, zipSync } from "fflate";
import { conditionalRuleFormula, withConditionalFormulaMatches } from "./conditional-format-formulas.ts";
import { initSync, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";
import type { XlsxConditionalStyledRule } from "./types.ts";

initSync({
  module: readFileSync(
    new URL("./duke-sheets/generated/duke_sheets_wasm_bg.wasm", import.meta.url),
  ),
});

// A1:A3 = 5, 15, 25; B2 = "x"; C1 = 10.
function workbook() {
  const parts = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    </Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Values" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    </Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1"><v>5</v></c><c r="C1"><v>10</v></c></row>
      <row r="2"><c r="A2"><v>15</v></c><c r="B2" t="inlineStr"><is><t>x</t></is></c></row>
      <row r="3"><c r="A3"><v>25</v></c></row>
    </sheetData></worksheet>`,
  };
  return Workbook.fromBytesForViewer(
    zipSync(Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)]))),
  );
}

function rule(ruleType: string, formulas: string[], range: string, operator?: string): XlsxConditionalStyledRule {
  const [start, end] = range.split(":").map((ref) => ({ col: ref.charCodeAt(0) - 65, row: Number(ref.slice(1)) - 1 }));
  return { formulas, kind: "styled", operator, priority: 1, ranges: [{ start: start!, end: end! }], ruleType, style: {} };
}

function matches(rules: XlsxConditionalStyledRule[]) {
  const book = workbook();
  const sheet = book.getSheet(0);
  try {
    return withConditionalFormulaMatches(rules, sheet).map((result) => (result.kind === "styled" ? result.matchedCells : undefined));
  } finally {
    sheet.free();
    book.free();
  }
}

test("a cell value rule compares each cell with its formula operand", () => {
  assert.equal(conditionalRuleFormula(rule("cellIs", ["$C$1+2"], "A1:A3", "greaterThan")), "A1>($C$1+2)");
  assert.equal(conditionalRuleFormula(rule("cellIs", ["10", "20"], "B2:B3", "notBetween")), "NOT(AND(B2>=(10),B2<=(20)))");
  assert.deepEqual(matches([rule("cellIs", ["$C$1+2"], "A1:A3", "greaterThan")]), [[1, 0, 2, 0]]);
  assert.deepEqual(matches([rule("cellIs", ["10", "20"], "A1:A3", "between")]), [[1, 0]]);
});

test("an expression rule shifts relative references for each cell", () => {
  assert.deepEqual(matches([rule("expression", ["$A1<$C$1"], "A1:B3")]), [[0, 0, 0, 1]]);
});

test("a blanks rule uses the formula Excel saves with it", () => {
  assert.deepEqual(matches([rule("containsBlanks", ["LEN(TRIM(B1))=0"], "B1:B3")]), [[0, 1, 2, 1]]);
});

test("rules ranked across their range are left to the viewer", () => {
  assert.equal(conditionalRuleFormula(rule("top10", [], "A1:A3")), null);
  assert.deepEqual(matches([rule("top10", [], "A1:A3")]), [undefined]);
});
