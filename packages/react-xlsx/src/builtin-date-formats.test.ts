import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { strToU8, zipSync } from "fflate";
import { initSync, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";

initSync({
  module: readFileSync(
    new URL("./duke-sheets/generated/duke_sheets_wasm_bg.wasm", import.meta.url),
  ),
});

// Cells using the built-in short date (14) and short date and time (22)
// formats, which have no numFmt entry of their own.
function builtinDateWorkbook(): Uint8Array {
  const parts = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
    "xl/styles.xml": `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="3">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1" s="1"><v>44941</v></c></row>
      <row r="2"><c r="A2" s="1"><v>45631</v></c></row>
      <row r="3"><c r="A3" s="2"><v>44941.5625</v></c></row>
    </sheetData></worksheet>`,
  };
  return zipSync(
    Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])),
  );
}

for (const [mode, load] of [
  ["viewer", Workbook.fromBytesForViewer],
  ["editable", Workbook.fromBytes],
] as const) {
  test(`${mode} shows built-in short dates the way Excel does in en-US`, () => {
    const workbook = load(builtinDateWorkbook());
    const sheet = workbook.getSheet(0);
    try {
      assert.equal(sheet.getFormattedValue("A1"), "1/15/2023");
      assert.equal(sheet.getFormattedValue("A2"), "12/5/2024");
      assert.equal(sheet.getFormattedValue("A3"), "1/15/2023 13:30");
    } finally {
      sheet.free();
      workbook.free();
    }
  });
}
