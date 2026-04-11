import type { FormulaHelpAvailability, FormulaHelpEntry } from '@/formulas/formulaHelpCatalog'

export interface FormulaCaretSelection {
  start: number
  end: number
}

export interface FormulaAutocompleteMatch {
  query: string
  replacementStart: number
  replacementEnd: number
}

export interface FormulaAutocompleteSuggestion {
  name: string
  signature: string
  summary: string
  example: string
  availability: FormulaHelpAvailability
  argumentLabels: string[]
}

export interface FormulaFunctionTemplate {
  text: string
  argumentRanges: Array<{ start: number; end: number }>
}

export interface FormulaSignatureHintSegment {
  id: string
  text: string
  kind: 'function' | 'argument' | 'separator' | 'punctuation'
  active: boolean
}

export interface FormulaSignatureHint {
  name: string
  segments: FormulaSignatureHintSegment[]
}

export interface FormulaFunctionContext {
  name: string
  openIndex: number
  closeIndex: number | null
  argumentRanges: Array<{ start: number; end: number }>
  activeArgumentIndex: number | null
}

interface FormulaFunctionStackEntry {
  kind: 'function' | 'group'
  openIndex: number
  name?: string
  currentArgumentStart?: number
  argumentRanges?: Array<{ start: number; end: number }>
}

export function isFormulaFunctionIdentifierCharacter(character: string | undefined) {
  return typeof character === 'string' && /^[A-Za-z0-9_]$/.test(character)
}

function findPreviousMeaningfulCharacter(text: string, fromIndex: number): string | null {
  for (let index = fromIndex - 1; index >= 0; index -= 1) {
    const character = text[index]
    if (character && !/\s/.test(character)) {
      return character
    }
  }

  return null
}

function isFormulaFunctionContextCharacter(character: string | null) {
  return (
    character === '=' ||
    character === '(' ||
    character === ',' ||
    character === '+' ||
    character === '-' ||
    character === '*' ||
    character === '/' ||
    character === '<' ||
    character === '>' ||
    character === '&'
  )
}

export function resolveFormulaAutocompleteMatch(
  rawInput: string,
  selection: FormulaCaretSelection,
): FormulaAutocompleteMatch | null {
  if (!rawInput.startsWith('=') || selection.start !== selection.end) {
    return null
  }

  const caret = selection.end
  if (caret < 1) {
    return null
  }

  let replacementStart = caret
  while (
    replacementStart > 0 &&
    isFormulaFunctionIdentifierCharacter(rawInput[replacementStart - 1])
  ) {
    replacementStart -= 1
  }

  let replacementEnd = caret
  while (
    replacementEnd < rawInput.length &&
    isFormulaFunctionIdentifierCharacter(rawInput[replacementEnd])
  ) {
    replacementEnd += 1
  }

  const previousCharacter = findPreviousMeaningfulCharacter(rawInput, replacementStart)
  if (!isFormulaFunctionContextCharacter(previousCharacter)) {
    return null
  }

  return {
    query: rawInput.slice(replacementStart, caret),
    replacementStart,
    replacementEnd,
  }
}

export function buildFormulaAutocompleteSuggestions(
  entries: readonly FormulaHelpEntry[],
  match: FormulaAutocompleteMatch | null,
) {
  if (!match) {
    return []
  }

  const normalizedQuery = match.query.trim().toUpperCase()
  if (!normalizedQuery) {
    return []
  }

  const suggestions = entries.map((entry) => ({
    name: entry.name.trim().toUpperCase(),
    signature: entry.signature,
    summary: entry.summary,
    example: entry.example,
    availability: entry.availability,
    argumentLabels: extractSignatureArgumentLabels(entry.signature),
  }))
  const rankMatches = (items: readonly FormulaAutocompleteSuggestion[]) =>
    [...items].sort((left, right) => {
      const leftLengthDelta = Math.abs(left.name.length - normalizedQuery.length)
      const rightLengthDelta = Math.abs(right.name.length - normalizedQuery.length)
      if (leftLengthDelta !== rightLengthDelta) {
        return leftLengthDelta - rightLengthDelta
      }

      return left.name.localeCompare(right.name)
    })

  const startsWithMatches = rankMatches(
    suggestions.filter((item) => item.name.startsWith(normalizedQuery)),
  )
  const includesMatches =
    normalizedQuery.length === 0
      ? []
      : rankMatches(
          suggestions.filter(
            (item) =>
              !item.name.startsWith(normalizedQuery) && item.name.includes(normalizedQuery),
          ),
        )

  return [...startsWithMatches, ...includesMatches].slice(0, 8)
}

export function buildFormulaFunctionTemplate(
  suggestion: Pick<FormulaAutocompleteSuggestion, 'name' | 'argumentLabels'>,
): FormulaFunctionTemplate {
  if (suggestion.argumentLabels.length === 0) {
    return {
      text: `${suggestion.name}()`,
      argumentRanges: [],
    }
  }

  const text = `${suggestion.name}()`
  const caret = `${suggestion.name}(`.length

  return {
    text,
    argumentRanges: [{ start: caret, end: caret }],
  }
}

export function buildFormulaSignatureHint(
  signature: string,
  activeArgumentIndex: number | null,
): FormulaSignatureHint | null {
  const openIndex = signature.indexOf('(')
  const closeIndex = signature.lastIndexOf(')')
  if (openIndex < 0 || closeIndex <= openIndex) {
    return null
  }

  const name = signature.slice(0, openIndex).trim().toUpperCase()
  const argsSource = signature.slice(openIndex + 1, closeIndex)
  const rawArgs = argsSource
    .split(',')
    .map((part) => part.trim())
    .filter((part) => part.length > 0)

  const hasVariadicTail = rawArgs.some((arg) => arg === '...' || arg === '…' || /\.\.\.$/.test(arg))
  const args = rawArgs.filter((arg) => arg !== '...' && arg !== '…' && !/\.\.\.$/.test(arg))
  const visibleArgumentCount = hasVariadicTail
    ? Math.max(args.length + 1, (activeArgumentIndex ?? 0) + 1)
    : args.length
  const visibleArgs = Array.from({ length: visibleArgumentCount }, (_, index) => {
    if (index < args.length) {
      return args[index] as string
    }

    return buildVariadicArgumentLabel(args, index)
  })

  const segments: FormulaSignatureHintSegment[] = [
    {
      id: 'fn-name',
      text: name,
      kind: 'function',
      active: false,
    },
    {
      id: 'fn-open',
      text: '(',
      kind: 'punctuation',
      active: false,
    },
  ]

  visibleArgs.forEach((arg, index) => {
    if (index > 0) {
      segments.push({
        id: `sep-${index}`,
        text: ', ',
        kind: 'separator',
        active: false,
      })
    }

    segments.push({
      id: `arg-${index}`,
      text: arg,
      kind: 'argument',
      active: activeArgumentIndex === index,
    })
  })

  if (hasVariadicTail) {
    if (visibleArgs.length > 0) {
      segments.push({
        id: 'sep-ellipsis',
        text: ', ',
        kind: 'separator',
        active: false,
      })
    }

    segments.push({
      id: 'arg-ellipsis',
      text: '...',
      kind: 'argument',
      active: false,
    })
  }

  segments.push({
    id: 'fn-close',
    text: ')',
    kind: 'punctuation',
    active: false,
  })

  return {
    name,
    segments,
  }
}

function buildVariadicArgumentLabel(baseArgs: string[], index: number) {
  const lastLabel = baseArgs[baseArgs.length - 1] ?? 'value1'
  const match = /^(.*?)(\d+)$/.exec(lastLabel)
  if (!match) {
    return `${lastLabel}${index + 1}`
  }

  const [, prefix, numericSuffix] = match
  const startIndex = Number(numericSuffix)
  return `${prefix}${startIndex + (index - baseArgs.length) + 1}`
}

export function resolveFormulaFunctionContext(
  rawInput: string,
  selection: FormulaCaretSelection,
): FormulaFunctionContext | null {
  if (!rawInput.startsWith('=')) {
    return null
  }

  const caret = selection.end
  const stack: FormulaFunctionStackEntry[] = []
  let candidate: FormulaFunctionContext | null = null
  let bracketDepth = 0
  let quotedBy: '"' | "'" | null = null

  const commitFunctionContext = (
    entry: FormulaFunctionStackEntry & {
      kind: 'function'
      name: string
      currentArgumentStart: number
      argumentRanges: Array<{ start: number; end: number }>
    },
    closeIndex: number | null,
    endIndex: number,
  ) => {
    const nextRanges =
      entry.argumentRanges.length > 0 || endIndex > entry.currentArgumentStart
        ? [
            ...entry.argumentRanges,
            trimFormulaArgumentRange(rawInput, entry.currentArgumentStart, endIndex),
          ].filter((range) => range.end >= range.start)
        : []

    if (caret < entry.openIndex + 1 || (closeIndex !== null && caret > closeIndex + 1)) {
      return
    }

    if (candidate && candidate.openIndex > entry.openIndex) {
      return
    }

    candidate = {
      name: entry.name,
      openIndex: entry.openIndex,
      closeIndex,
      argumentRanges: nextRanges,
      activeArgumentIndex: resolveActiveArgumentIndex(nextRanges, selection),
    }
  }

  for (let index = 0; index < rawInput.length; index += 1) {
    const character = rawInput[index]
    const previousCharacter = rawInput[index - 1] ?? null

    if (quotedBy) {
      if (character === quotedBy && previousCharacter !== '\\') {
        quotedBy = null
      }
      continue
    }

    if (character === '"' || character === "'") {
      quotedBy = character
      continue
    }

    if (character === '[') {
      bracketDepth += 1
      continue
    }

    if (character === ']') {
      bracketDepth = Math.max(0, bracketDepth - 1)
      continue
    }

    if (bracketDepth > 0) {
      continue
    }

    if (character === '(') {
      let nameStart = index
      while (nameStart > 0 && isFormulaFunctionIdentifierCharacter(rawInput[nameStart - 1])) {
        nameStart -= 1
      }

      const name = rawInput.slice(nameStart, index).trim().toUpperCase()
      const isFunction =
        name.length > 0 &&
        isFormulaFunctionContextCharacter(findPreviousMeaningfulCharacter(rawInput, nameStart))

      if (isFunction) {
        stack.push({
          kind: 'function',
          name,
          openIndex: index,
          currentArgumentStart: index + 1,
          argumentRanges: [],
        })
      } else {
        stack.push({
          kind: 'group',
          openIndex: index,
        })
      }

      continue
    }

    const currentEntry = stack[stack.length - 1] ?? null
    if (character === ',' && currentEntry?.kind === 'function') {
      currentEntry.argumentRanges?.push(
        trimFormulaArgumentRange(
          rawInput,
          currentEntry.currentArgumentStart ?? index,
          index,
        ),
      )
      currentEntry.currentArgumentStart = index + 1
      continue
    }

    if (character === ')') {
      const currentEntry = stack.pop() ?? null
      if (currentEntry?.kind === 'function') {
        commitFunctionContext(
          currentEntry as FormulaFunctionStackEntry & {
            kind: 'function'
            name: string
            currentArgumentStart: number
            argumentRanges: Array<{ start: number; end: number }>
          },
          index,
          index,
        )
      }
    }
  }

  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const currentEntry = stack[index]
    if (currentEntry?.kind !== 'function') {
      continue
    }

    commitFunctionContext(
      currentEntry as FormulaFunctionStackEntry & {
        kind: 'function'
        name: string
        currentArgumentStart: number
        argumentRanges: Array<{ start: number; end: number }>
      },
      null,
      rawInput.length,
    )
    break
  }

  return candidate
}

export function resolveNextFormulaArgumentSelection(
  rawInput: string,
  selection: FormulaCaretSelection,
  direction: 1 | -1 = 1,
) {
  const context = resolveFormulaFunctionContext(rawInput, selection)
  if (!context || context.argumentRanges.length === 0) {
    return null
  }

  const ranges = context.argumentRanges
  const currentIndex =
    context.activeArgumentIndex ??
    (direction === 1
      ? ranges.findIndex((range) => selection.end < range.end)
      : findLastIndex(ranges, (range) => selection.start > range.start))

  const nextIndex =
    currentIndex === null || currentIndex < 0
      ? direction === 1
        ? 0
        : ranges.length - 1
      : currentIndex + direction

  if (nextIndex < 0 || nextIndex >= ranges.length) {
    return null
  }

  return ranges[nextIndex] ?? null
}

function extractSignatureArgumentLabels(signature: string) {
  const openIndex = signature.indexOf('(')
  const closeIndex = signature.lastIndexOf(')')
  if (openIndex < 0 || closeIndex <= openIndex) {
    return []
  }

  return signature
    .slice(openIndex + 1, closeIndex)
    .split(',')
    .map((part) => part.trim())
    .filter((part) => part.length > 0 && part !== '...' && part !== '…')
    .map((part) => part.replace(/\?$/, '').replace(/\s*\.\.\.$/, '').trim())
    .filter((part) => part.length > 0)
}

function trimFormulaArgumentRange(rawInput: string, start: number, end: number) {
  let nextStart = start
  let nextEnd = end

  while (nextStart < nextEnd && /\s/.test(rawInput[nextStart] ?? '')) {
    nextStart += 1
  }

  while (nextEnd > nextStart && /\s/.test(rawInput[nextEnd - 1] ?? '')) {
    nextEnd -= 1
  }

  return {
    start: nextStart,
    end: nextEnd,
  }
}

function resolveActiveArgumentIndex(
  ranges: Array<{ start: number; end: number }>,
  selection: FormulaCaretSelection,
) {
  const overlapIndex = ranges.findIndex((range) =>
    selection.start <= range.end && selection.end >= range.start,
  )
  if (overlapIndex >= 0) {
    return overlapIndex
  }

  const containingIndex = ranges.findIndex(
    (range) => selection.end >= range.start && selection.end <= range.end,
  )
  if (containingIndex >= 0) {
    return containingIndex
  }

  return null
}

function findLastIndex<T>(
  values: readonly T[],
  predicate: (value: T, index: number) => boolean,
) {
  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (predicate(values[index] as T, index)) {
      return index
    }
  }

  return -1
}
