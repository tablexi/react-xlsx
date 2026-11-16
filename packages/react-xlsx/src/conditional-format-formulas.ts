import type { XlsxConditionalFormatRule, XlsxConditionalStyledRule } from "./types";

type FormulaMatcher = {
  conditionalFormulaMatches(formula: string, ranges: Uint32Array): Uint32Array;
};

const CELL_IS_OPERATORS: Record<string, string> = {
  equal: "=",
  greaterThan: ">",
  greaterThanOrEqual: ">=",
  lessThan: "<",
  lessThanOrEqual: "<=",
  notEqual: "<>"
};

function columnLetters(col: number) {
  let letters = "";
  for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    letters = String.fromCharCode(65 + ((value - 1) % 26)) + letters;
  }
  return letters;
}

/**
 * The formula that decides whether a styled rule applies, written for the
 * top-left cell of its first range as Excel does. Excel saves one with every
 * expression, text, blank, error and date rule; a "cell value" rule becomes a
 * comparison against the cell. Rules ranked across their range (top/bottom,
 * average, duplicate/unique) have none.
 */
export function conditionalRuleFormula(rule: XlsxConditionalStyledRule): string | null {
  const [first, second] = rule.formulas;
  if (rule.ruleType !== "cellIs") {
    return first ?? null;
  }

  const start = rule.ranges[0]?.start;
  if (!start || first === undefined) {
    return null;
  }
  const cell = `${columnLetters(start.col)}${start.row + 1}`;
  if (rule.operator === "between" || rule.operator === "notBetween") {
    if (second === undefined) {
      return null;
    }
    const between = `AND(${cell}>=(${first}),${cell}<=(${second}))`;
    return rule.operator === "between" ? between : `NOT(${between})`;
  }
  const operator = CELL_IS_OPERATORS[rule.operator ?? "equal"];
  return operator ? `${cell}${operator}(${first})` : null;
}

/**
 * Attach the matching cells to each formula-based styled rule, evaluated by
 * the parser against the whole workbook, so the viewer can apply them without
 * the cells the formulas refer to.
 */
export function withConditionalFormulaMatches(
  rules: XlsxConditionalFormatRule[],
  worksheet: FormulaMatcher
): XlsxConditionalFormatRule[] {
  return rules.map((rule) => {
    const formula = rule.kind === "styled" ? conditionalRuleFormula(rule) : null;
    if (!formula) {
      return rule;
    }
    const ranges = Uint32Array.from(rule.ranges.flatMap(({ start, end }) => [start.row, start.col, end.row, end.col]));
    return { ...rule, matchedCells: Array.from(worksheet.conditionalFormulaMatches(formula, ranges)) };
  });
}
