import { FilterOperator, IRollupFilter } from '../models/IPageRollupProps';

/* eslint-disable @typescript-eslint/no-explicit-any */

const decoder = typeof document !== 'undefined' ? document.createElement('textarea') : undefined;

/** Strips tags and decodes entities from rich text (Note) fields. */
export function stripHtml(value: string): string {
  if (!value || (value.indexOf('<') === -1 && value.indexOf('&') === -1)) {
    return value;
  }
  const noTags = value.replace(/<br\s*\/?>/gi, ' ').replace(/<\/(p|div|li)>/gi, ' ').replace(/<[^>]+>/g, '');
  if (decoder) {
    decoder.innerHTML = noTags;
    return decoder.value.replace(/\s+/g, ' ').trim();
  }
  return noTags.replace(/\s+/g, ' ').trim();
}

function objectToText(obj: any): string {
  if (obj === null || obj === undefined) {
    return '';
  }
  if (typeof obj !== 'object') {
    return String(obj);
  }
  // Managed metadata, lookup, person, URL and image values as returned by RenderListDataAsStream
  return obj.Label || obj.lookupValue || obj.title || obj.Title || obj.desc || obj.Description || obj.value || '';
}

/**
 * Returns every individual value a row holds for a field, as plain text.
 * Multi-value columns (choice, managed metadata, lookup, people) give one entry per value.
 */
export function getFieldValues(row: any, field: string): string[] {
  const raw = row ? row[field] : undefined;
  if (raw === null || raw === undefined || raw === '') {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw.map(objectToText).map((v) => stripHtml(String(v)).trim()).filter((v) => v.length > 0);
  }
  if (typeof raw === 'object') {
    const text = objectToText(raw);
    return text ? [stripHtml(String(text))] : [];
  }
  const text = String(raw);
  // Multi-choice values can arrive as ";#A;#B;#"
  if (text.indexOf(';#') !== -1) {
    return text.split(';#').map((v) => v.trim()).filter((v) => v.length > 0);
  }
  // URL / hyperlink fields: prefer the description if present
  const desc = row[field + '.desc'];
  if (typeof desc === 'string' && desc.trim()) {
    return [desc.trim()];
  }
  const plain = stripHtml(text).trim();
  return plain ? [plain] : [];
}

/** Like getFieldText but keeps paragraph and line breaks from rich text, for the hover details panel. */
export function getFieldTextMultiline(row: any, field: string): string {
  const raw = row ? row[field] : undefined;
  if (typeof raw !== 'string' || (raw.indexOf('<') === -1 && raw.indexOf('\n') === -1)) {
    return getFieldText(row, field);
  }
  const withBreaks = raw
    .replace(/\r\n?/g, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|h[1-6])>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<[^>]+>/g, '');
  let text = withBreaks;
  if (decoder) {
    decoder.innerHTML = withBreaks;
    text = decoder.value;
  }
  return text
    .split('\n')
    .map((line) => line.replace(/[ \t\u00a0]+/g, ' ').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function getFieldText(row: any, field: string): string {
  return getFieldValues(row, field).join(', ');
}

/** Splits a filter value box such as "Lesson Learning; Safety" into separate values. */
export function parseFilterValues(input: string): string[] {
  return (input || '')
    .split(/[;,|]/)
    .map((v) => v.trim())
    .filter((v) => v.length > 0);
}

function matchesFilter(row: any, filter: IRollupFilter): boolean {
  const values = getFieldValues(row, filter.field).map((v) => v.toLowerCase());
  const wanted = filter.values.map((v) => v.toLowerCase());
  const op: FilterOperator = filter.operator;

  switch (op) {
    case 'empty':
      return values.length === 0;
    case 'notEmpty':
      return values.length > 0;
    case 'neq':
      return !values.some((v) => wanted.indexOf(v) !== -1);
    case 'contains':
      return values.some((v) => wanted.some((w) => v.indexOf(w) !== -1));
    case 'eq':
    default:
      return values.some((v) => wanted.indexOf(v) !== -1);
  }
}

/** Applies the column filters to a row. Filters with no value (except empty/not empty) are ignored. */
export function rowMatches(row: any, filters: IRollupFilter[], match: 'all' | 'any'): boolean {
  const active = filters.filter((f) => f.field && (f.values.length > 0 || f.operator === 'empty' || f.operator === 'notEmpty'));
  if (active.length === 0) {
    return true;
  }
  return match === 'any' ? active.some((f) => matchesFilter(row, f)) : active.every((f) => matchesFilter(row, f));
}

/** 1234 -> "1.2K" */
export function formatCount(n: number): string {
  if (n >= 1000000) {
    return `${(n / 1000000).toFixed(n >= 10000000 ? 0 : 1).replace(/\.0$/, '')}M`;
  }
  if (n >= 1000) {
    return `${(n / 1000).toFixed(n >= 10000 ? 0 : 1).replace(/\.0$/, '')}K`;
  }
  return String(n);
}
