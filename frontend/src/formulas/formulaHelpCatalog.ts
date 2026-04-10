import { DATAGRID_DEFAULT_FORMULA_FUNCTIONS } from '@affino/datagrid-formula-engine'

export type FormulaHelpCategory = 'Numeric' | 'Logic' | 'Text' | 'Date' | 'Advanced'
export type FormulaHelpAvailability = 'grid' | 'workbook'

export interface FormulaHelpEntry {
  name: string
  category: FormulaHelpCategory
  signature: string
  summary: string
  example: string
  availability: FormulaHelpAvailability
}

export interface FormulaReferencePattern {
  label: string
  example: string
  note: string
  availability?: FormulaHelpAvailability
}

const CATEGORY_ORDER: FormulaHelpCategory[] = ['Numeric', 'Logic', 'Text', 'Date', 'Advanced']

const FORMULA_HELP_METADATA: Record<
  string,
  Omit<FormulaHelpEntry, 'name'>
> = {
  ABS: { category: 'Numeric', signature: 'ABS(number)', summary: 'Returns the absolute value.', example: '=ABS(-42)', availability: 'grid' },
  AVG: { category: 'Numeric', signature: 'AVG(value1, value2, ...)', summary: 'Averages the supplied values.', example: '=AVG([Column 1]1, [Column 1]2, [Column 1]3)', availability: 'grid' },
  AVERAGE: { category: 'Numeric', signature: 'AVERAGE(value1, value2, ...)', summary: 'Alias of AVG for mean values.', example: '=AVERAGE([Column 1]1:[Column 1]5)', availability: 'grid' },
  AVGW: { category: 'Numeric', signature: 'AVGW(values, weights)', summary: 'Calculates a weighted average.', example: '=AVGW([Score]1:[Score]3, [Weight]1:[Weight]3)', availability: 'grid' },
  CEIL: { category: 'Numeric', signature: 'CEIL(number, significance?)', summary: 'Rounds up to the next multiple.', example: '=CEIL(12.2, 0.5)', availability: 'grid' },
  CEILING: { category: 'Numeric', signature: 'CEILING(number, significance?)', summary: 'Alias of CEIL.', example: '=CEILING([Budget]1, 100)', availability: 'grid' },
  CHAR: { category: 'Numeric', signature: 'CHAR(code)', summary: 'Returns the character for an ASCII code.', example: '=CHAR(65)', availability: 'grid' },
  COUNT: { category: 'Numeric', signature: 'COUNT(value1, value2, ...)', summary: 'Counts non-empty values.', example: '=COUNT([Column 1]1:[Column 1]10)', availability: 'grid' },
  COUNTM: { category: 'Numeric', signature: 'COUNTM(value1, value2, ...)', summary: 'Counts present values, including mixed types.', example: '=COUNTM([Owner]1:[Owner]10)', availability: 'grid' },
  DECTOHEX: { category: 'Numeric', signature: 'DECTOHEX(number)', summary: 'Converts decimal to hexadecimal.', example: '=DECTOHEX(255)', availability: 'grid' },
  FLOOR: { category: 'Numeric', signature: 'FLOOR(number, significance?)', summary: 'Rounds down to the previous multiple.', example: '=FLOOR(12.8, 0.5)', availability: 'grid' },
  HEXTODEC: { category: 'Numeric', signature: 'HEXTODEC(hex)', summary: 'Converts hexadecimal text to decimal.', example: '=HEXTODEC("FF")', availability: 'grid' },
  INT: { category: 'Numeric', signature: 'INT(number)', summary: 'Truncates a number to an integer.', example: '=INT(12.99)', availability: 'grid' },
  LARGE: { category: 'Numeric', signature: 'LARGE(values, rank)', summary: 'Returns the N-th largest value.', example: '=LARGE([Amount]1:[Amount]10, 2)', availability: 'grid' },
  MAX: { category: 'Numeric', signature: 'MAX(value1, value2, ...)', summary: 'Returns the largest numeric value.', example: '=MAX([Amount]1:[Amount]10)', availability: 'grid' },
  MEDIAN: { category: 'Numeric', signature: 'MEDIAN(value1, value2, ...)', summary: 'Returns the median value.', example: '=MEDIAN([Amount]1:[Amount]7)', availability: 'grid' },
  MIN: { category: 'Numeric', signature: 'MIN(value1, value2, ...)', summary: 'Returns the smallest numeric value.', example: '=MIN([Amount]1:[Amount]10)', availability: 'grid' },
  MOD: { category: 'Numeric', signature: 'MOD(number, divisor)', summary: 'Returns the remainder after division.', example: '=MOD(17, 5)', availability: 'grid' },
  MROUND: { category: 'Numeric', signature: 'MROUND(number, multiple)', summary: 'Rounds to the nearest multiple.', example: '=MROUND([Budget]1, 25)', availability: 'grid' },
  NPV: { category: 'Numeric', signature: 'NPV(rate, cashflow1, cashflow2, ...)', summary: 'Computes the net present value of a cashflow series.', example: '=NPV(0.08, [Cashflow]1:[Cashflow]5)', availability: 'grid' },
  PERCENTILE: { category: 'Numeric', signature: 'PERCENTILE(values, percentile)', summary: 'Returns the value at a percentile.', example: '=PERCENTILE([Amount]1:[Amount]10, 0.9)', availability: 'grid' },
  POW: { category: 'Numeric', signature: 'POW(base, exponent)', summary: 'Raises a number to a power.', example: '=POW(2, 8)', availability: 'grid' },
  RANKAVG: { category: 'Numeric', signature: 'RANKAVG(value, values, order?)', summary: 'Returns the average rank for duplicate values.', example: '=RANKAVG([Score]1, [Score]1:[Score]10)', availability: 'grid' },
  RANKEQ: { category: 'Numeric', signature: 'RANKEQ(value, values, order?)', summary: 'Returns the rank of a value.', example: '=RANKEQ([Score]1, [Score]1:[Score]10)', availability: 'grid' },
  ROUND: { category: 'Numeric', signature: 'ROUND(number, digits?)', summary: 'Rounds to a given number of digits.', example: '=ROUND(12.3456, 2)', availability: 'grid' },
  ROUNDDOWN: { category: 'Numeric', signature: 'ROUNDDOWN(number, digits?)', summary: 'Rounds toward zero.', example: '=ROUNDDOWN(12.987, 2)', availability: 'grid' },
  ROUNDUP: { category: 'Numeric', signature: 'ROUNDUP(number, digits?)', summary: 'Rounds away from zero.', example: '=ROUNDUP(12.123, 2)', availability: 'grid' },
  SMALL: { category: 'Numeric', signature: 'SMALL(values, rank)', summary: 'Returns the N-th smallest value.', example: '=SMALL([Amount]1:[Amount]10, 2)', availability: 'grid' },
  STDEVA: { category: 'Numeric', signature: 'STDEVA(value1, value2, ...)', summary: 'Sample standard deviation.', example: '=STDEVA([Score]1:[Score]10)', availability: 'grid' },
  STDEVP: { category: 'Numeric', signature: 'STDEVP(value1, value2, ...)', summary: 'Population standard deviation.', example: '=STDEVP([Score]1:[Score]10)', availability: 'grid' },
  STDEVPA: { category: 'Numeric', signature: 'STDEVPA(value1, value2, ...)', summary: 'Population standard deviation including mixed values.', example: '=STDEVPA([Score]1:[Score]10)', availability: 'grid' },
  STDEVS: { category: 'Numeric', signature: 'STDEVS(value1, value2, ...)', summary: 'Sample standard deviation alias.', example: '=STDEVS([Score]1:[Score]10)', availability: 'grid' },
  SUM: { category: 'Numeric', signature: 'SUM(value1, value2, ...)', summary: 'Adds numbers and ranges together.', example: '=SUM([Column 1]1:[Column 1]5)', availability: 'grid' },
  UNICHAR: { category: 'Numeric', signature: 'UNICHAR(codePoint)', summary: 'Returns the character for a Unicode code point.', example: '=UNICHAR(10003)', availability: 'grid' },

  AND: { category: 'Logic', signature: 'AND(condition1, condition2, ...)', summary: 'Returns TRUE only if all conditions are true.', example: '=AND([Done]1 = 1, [Approved]1 = 1)', availability: 'grid' },
  COALESCE: { category: 'Logic', signature: 'COALESCE(value1, value2, ...)', summary: 'Returns the first non-blank value.', example: '=COALESCE([Owner]1, [Backup Owner]1, "Unassigned")', availability: 'grid' },
  CONTAINS: { category: 'Logic', signature: 'CONTAINS(search, text)', summary: 'Checks whether text contains a substring.', example: '=CONTAINS("pro", [Title]1)', availability: 'grid' },
  COUNTIF: { category: 'Logic', signature: 'COUNTIF(range, criterion)', summary: 'Counts values matching a criterion.', example: '=COUNTIF([Status]1:[Status]10, "Done")', availability: 'grid' },
  HAS: { category: 'Logic', signature: 'HAS(values, target)', summary: 'Checks whether a list contains a value.', example: '=HAS(RANGE("Red", "Green", "Blue"), "Green")', availability: 'grid' },
  IF: { category: 'Logic', signature: 'IF(condition, whenTrue, whenFalse)', summary: 'Returns one value when a condition is true and another when false.', example: '=IF([Amount]1 > 1000, "Large", "Small")', availability: 'grid' },
  IFERROR: { category: 'Logic', signature: 'IFERROR(value, fallback)', summary: 'Replaces a formula error with a fallback.', example: '=IFERROR([Amount]1 / [Count]1, 0)', availability: 'grid' },
  IFS: { category: 'Logic', signature: 'IFS(test1, value1, test2, value2, ...)', summary: 'Returns the first value whose test is true.', example: '=IFS([Score]1 >= 90, "A", [Score]1 >= 80, "B", TRUE, "C")', availability: 'grid' },
  IN: { category: 'Logic', signature: 'IN(value, option1, option2, ...)', summary: 'Checks whether a value matches any option.', example: '=IN([Status]1, "Draft", "Ready", "Done")', availability: 'grid' },
  ISBLANK: { category: 'Logic', signature: 'ISBLANK(value)', summary: 'Checks whether a value is blank.', example: '=ISBLANK([Notes]1)', availability: 'grid' },
  ISBOOLEAN: { category: 'Logic', signature: 'ISBOOLEAN(value)', summary: 'Checks whether a value is boolean.', example: '=ISBOOLEAN(TRUE)', availability: 'grid' },
  ISDATE: { category: 'Logic', signature: 'ISDATE(value)', summary: 'Checks whether a value is a date.', example: '=ISDATE([Due]1)', availability: 'grid' },
  ISERROR: { category: 'Logic', signature: 'ISERROR(value)', summary: 'Checks whether a value is a formula error.', example: '=ISERROR([Amount]1 / [Count]1)', availability: 'grid' },
  ISEVEN: { category: 'Logic', signature: 'ISEVEN(number)', summary: 'Checks whether a number is even.', example: '=ISEVEN([Column 1]1)', availability: 'grid' },
  ISNUMBER: { category: 'Logic', signature: 'ISNUMBER(value)', summary: 'Checks whether a value is numeric.', example: '=ISNUMBER([Amount]1)', availability: 'grid' },
  ISODD: { category: 'Logic', signature: 'ISODD(number)', summary: 'Checks whether a number is odd.', example: '=ISODD([Column 1]1)', availability: 'grid' },
  ISTEXT: { category: 'Logic', signature: 'ISTEXT(value)', summary: 'Checks whether a value is text.', example: '=ISTEXT([Title]1)', availability: 'grid' },
  NOT: { category: 'Logic', signature: 'NOT(condition)', summary: 'Negates a boolean condition.', example: '=NOT([Done]1 = 1)', availability: 'grid' },
  OR: { category: 'Logic', signature: 'OR(condition1, condition2, ...)', summary: 'Returns TRUE if any condition is true.', example: '=OR([Blocked]1 = 1, [At Risk]1 = 1)', availability: 'grid' },

  CONCAT: { category: 'Text', signature: 'CONCAT(value1, value2, ...)', summary: 'Concatenates values without a delimiter.', example: '=CONCAT([First Name]1, " ", [Last Name]1)', availability: 'grid' },
  FIND: { category: 'Text', signature: 'FIND(search, text, start?)', summary: 'Returns the position of a substring.', example: '=FIND("-", [Order ID]1)', availability: 'grid' },
  JOIN: { category: 'Text', signature: 'JOIN(values, delimiter?)', summary: 'Joins values with a delimiter.', example: '=JOIN(RANGE("UI", "API", "QA"), ", ")', availability: 'grid' },
  LEFT: { category: 'Text', signature: 'LEFT(text, count?)', summary: 'Returns characters from the start of text.', example: '=LEFT([Order ID]1, 3)', availability: 'grid' },
  LEN: { category: 'Text', signature: 'LEN(text)', summary: 'Returns text length.', example: '=LEN([Title]1)', availability: 'grid' },
  LOWER: { category: 'Text', signature: 'LOWER(text)', summary: 'Converts text to lowercase.', example: '=LOWER([Email]1)', availability: 'grid' },
  MID: { category: 'Text', signature: 'MID(text, start, count)', summary: 'Returns a substring from the middle of text.', example: '=MID([Order ID]1, 4, 3)', availability: 'grid' },
  REPLACE: { category: 'Text', signature: 'REPLACE(text, start, count, newText)', summary: 'Replaces a substring at a given position.', example: '=REPLACE([Code]1, 1, 3, "NEW")', availability: 'grid' },
  RIGHT: { category: 'Text', signature: 'RIGHT(text, count?)', summary: 'Returns characters from the end of text.', example: '=RIGHT([Order ID]1, 4)', availability: 'grid' },
  SUBSTITUTE: { category: 'Text', signature: 'SUBSTITUTE(text, oldText, newText, instance?)', summary: 'Replaces matching text.', example: '=SUBSTITUTE([Title]1, "Draft", "Ready")', availability: 'grid' },
  TRIM: { category: 'Text', signature: 'TRIM(text)', summary: 'Trims whitespace around text.', example: '=TRIM([Client Name]1)', availability: 'grid' },
  UPPER: { category: 'Text', signature: 'UPPER(text)', summary: 'Converts text to uppercase.', example: '=UPPER([Code]1)', availability: 'grid' },
  VALUE: { category: 'Text', signature: 'VALUE(text)', summary: 'Converts text into a number when possible.', example: '=VALUE("42.5")', availability: 'grid' },

  DATE: { category: 'Date', signature: 'DATE(year, month, day)', summary: 'Builds a UTC date from parts.', example: '=DATE(2026, 4, 9)', availability: 'grid' },
  DATEONLY: { category: 'Date', signature: 'DATEONLY(date)', summary: 'Strips the time part from a date.', example: '=DATEONLY(TODAY())', availability: 'grid' },
  DAY: { category: 'Date', signature: 'DAY(date)', summary: 'Returns the day of the month.', example: '=DAY([Due Date]1)', availability: 'grid' },
  MONTH: { category: 'Date', signature: 'MONTH(date)', summary: 'Returns the month number.', example: '=MONTH([Due Date]1)', availability: 'grid' },
  NETDAYS: { category: 'Date', signature: 'NETDAYS(startDate, endDate)', summary: 'Returns the number of days between two dates.', example: '=NETDAYS([Start]1, [End]1)', availability: 'grid' },
  NETWORKDAY: { category: 'Date', signature: 'NETWORKDAY(startDate, endDate, holidays?)', summary: 'Counts working days between two dates.', example: '=NETWORKDAY([Start]1, [End]1)', availability: 'grid' },
  NETWORKDAYS: { category: 'Date', signature: 'NETWORKDAYS(startDate, endDate, holidays?)', summary: 'Alias of NETWORKDAY.', example: '=NETWORKDAYS([Start]1, [End]1)', availability: 'grid' },
  TIME: { category: 'Date', signature: 'TIME(value)', summary: 'Parses a time-like value.', example: '=TIME("14:30")', availability: 'grid' },
  TODAY: { category: 'Date', signature: 'TODAY(offset?)', summary: 'Returns today, optionally shifted by days.', example: '=TODAY(7)', availability: 'grid' },
  WEEKDAY: { category: 'Date', signature: 'WEEKDAY(date)', summary: 'Returns the weekday index.', example: '=WEEKDAY([Due Date]1)', availability: 'grid' },
  WEEKNUMBER: { category: 'Date', signature: 'WEEKNUMBER(date)', summary: 'Returns the week number in the year.', example: '=WEEKNUMBER([Due Date]1)', availability: 'grid' },
  WORKDAY: { category: 'Date', signature: 'WORKDAY(startDate, offset, holidays?)', summary: 'Moves a date by working days only.', example: '=WORKDAY([Start]1, 5)', availability: 'grid' },
  YEAR: { category: 'Date', signature: 'YEAR(date)', summary: 'Returns the year number.', example: '=YEAR([Due Date]1)', availability: 'grid' },
  YEARDAY: { category: 'Date', signature: 'YEARDAY(date)', summary: 'Returns the day number within the year.', example: '=YEARDAY([Due Date]1)', availability: 'grid' },

  ARRAY: { category: 'Advanced', signature: 'ARRAY(value1, value2, ...)', summary: 'Builds an array from explicit values.', example: '=ARRAY(1, 2, 3)', availability: 'grid' },
  AVERAGEIF: { category: 'Advanced', signature: 'AVERAGEIF(criteriaRange, criterion, averageRange?)', summary: 'Averages values that match a condition.', example: '=AVERAGEIF([Status]1:[Status]10, "Done", [Amount]1:[Amount]10)', availability: 'grid' },
  CHOOSE: { category: 'Advanced', signature: 'CHOOSE(index, option1, option2, ...)', summary: 'Returns the N-th option.', example: '=CHOOSE(2, "Draft", "Active", "Done")', availability: 'grid' },
  COLLECT: { category: 'Advanced', signature: 'COLLECT(values, criteriaRange1, criterion1, ...)', summary: 'Collects values that match one or more criteria.', example: '=COLLECT([Amount]1:[Amount]10, [Status]1:[Status]10, "Done")', availability: 'grid' },
  COUNTIFS: { category: 'Advanced', signature: 'COUNTIFS(range1, criterion1, range2, criterion2, ...)', summary: 'Counts rows that satisfy multiple criteria.', example: '=COUNTIFS([Status]1:[Status]10, "Done", [Owner]1:[Owner]10, "Ana")', availability: 'grid' },
  DISTINCT: { category: 'Advanced', signature: 'DISTINCT(values)', summary: 'Returns unique values from a range.', example: '=DISTINCT([Owner]1:[Owner]10)', availability: 'grid' },
  INDEX: { category: 'Advanced', signature: 'INDEX(values, index, fallback?)', summary: 'Returns a value by 1-based position.', example: '=INDEX([Amount]1:[Amount]10, 3, 0)', availability: 'grid' },
  MATCH: { category: 'Advanced', signature: 'MATCH(value, values, matchMode?)', summary: 'Returns the index of a matching value.', example: '=MATCH("Done", [Status]1:[Status]10, 0)', availability: 'grid' },
  RANGE: { category: 'Advanced', signature: 'RANGE(value1, value2, ...)', summary: 'Builds a range-like array inline.', example: '=RANGE([Column 1]1, [Column 1]2, [Column 1]3)', availability: 'grid' },
  SUMIF: { category: 'Advanced', signature: 'SUMIF(criteriaRange, criterion, sumRange?)', summary: 'Sums values that match a condition.', example: '=SUMIF([Status]1:[Status]10, "Done", [Amount]1:[Amount]10)', availability: 'grid' },
  SUMIFS: { category: 'Advanced', signature: 'SUMIFS(sumRange, criteriaRange1, criterion1, ...)', summary: 'Sums values that satisfy multiple criteria.', example: '=SUMIFS([Amount]1:[Amount]10, [Status]1:[Status]10, "Done", [Owner]1:[Owner]10, "Ana")', availability: 'grid' },
  TABLE: { category: 'Advanced', signature: "TABLE('sheet', 'column')", summary: 'Returns values from another sheet column.', example: "=TABLE('Orders', 'Total')", availability: 'workbook' },
  RELATED: { category: 'Advanced', signature: "RELATED('sheet', sourceValue, lookupColumn, returnColumn, fallback?)", summary: 'Finds a related value in another sheet.', example: "=RELATED('Clients', [Client ID]@row, 'ID', 'Name', '')", availability: 'workbook' },
  ROLLUP: { category: 'Advanced', signature: "ROLLUP('sheet', matchColumn, sourceValue, valueColumn, method, fallback?)", summary: 'Aggregates related rows from another sheet.', example: "=ROLLUP('Orders', 'Client ID', [Client ID]@row, 'Total', 'sum', 0)", availability: 'workbook' },
  VLOOKUP: { category: 'Advanced', signature: 'VLOOKUP(value, values, columnNumber, exact?)', summary: 'Legacy compatibility lookup over a single return column.', example: '=VLOOKUP("Done", RANGE("Draft", "Done", "Blocked"), 1, 0)', availability: 'grid' },
  XLOOKUP: { category: 'Advanced', signature: 'XLOOKUP(value, lookupValues, returnValues, fallback?, matchMode?)', summary: 'Looks up a value and returns a paired result.', example: '=XLOOKUP("Done", RANGE("Draft", "Done", "Blocked"), RANGE(1, 2, 3), 0, 0)', availability: 'grid' },
}

export const FORMULA_REFERENCE_PATTERNS: readonly FormulaReferencePattern[] = [
  {
    label: 'Absolute cell',
    example: '=[Column 1]1 + [Column 1]2',
    note: 'Use row numbers for fixed references.',
    availability: 'grid',
  },
  {
    label: 'Current row',
    example: '=[Qty]@row * [Price]@row',
    note: 'Use @row when the formula should follow the current row.',
    availability: 'grid',
  },
  {
    label: 'Range',
    example: '=SUM([Amount]1:[Amount]10)',
    note: 'Same-column ranges work well with aggregate functions.',
    availability: 'grid',
  },
  {
    label: 'Cross sheet',
    example: "=TABLE('Orders', 'Total')",
    note: 'Workbook context is required for cross-sheet functions.',
    availability: 'workbook',
  },
] as const

function compareEntries(left: FormulaHelpEntry, right: FormulaHelpEntry) {
  const categoryDiff =
    CATEGORY_ORDER.indexOf(left.category) - CATEGORY_ORDER.indexOf(right.category)
  if (categoryDiff !== 0) {
    return categoryDiff
  }

  return left.name.localeCompare(right.name)
}

export const formulaHelpCatalog: readonly FormulaHelpEntry[] = Object.keys(
  DATAGRID_DEFAULT_FORMULA_FUNCTIONS,
)
  .map((name) => {
    const metadata = FORMULA_HELP_METADATA[name]
    if (metadata) {
      return {
        name,
        ...metadata,
      }
    }

    return {
      name,
      category: 'Advanced' as const,
      signature: `${name}(...)`,
      summary: 'Available in the current formula engine.',
      example: `=${name}(...)`,
      availability: 'grid' as const,
    }
  })
  .sort(compareEntries)

export const formulaHelpCategories = CATEGORY_ORDER
