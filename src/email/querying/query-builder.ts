import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.ts';

type FieldOperator = {
  $any?: string[];
  $all?: string[];
  $none?: string[];
};

type DateOperator = {
  $gte?: string;
  $lt?: string;
};

type FieldQuery = {
  from?: FieldOperator | string;
  to?: FieldOperator | string;
  cc?: FieldOperator | string;
  bcc?: FieldOperator | string;
  subject?: FieldOperator | string;
  text?: FieldOperator | string;
  body?: FieldOperator | string;
  categories?: FieldOperator | string;
  label?: FieldOperator | string;
};

/**
 * Outlook category mappings - case insensitive input to exact system categories
 * Note: These are user-defined categories in Outlook, not system categories like Gmail
 */
export const OUTLOOK_CATEGORIES = {
  personal: 'Personal',
  work: 'Work',
  family: 'Family',
  travel: 'Travel',
  important: 'Important',
  urgent: 'Urgent',
} as const;

/**
 * Validate and map category name to Outlook system category
 * Throws error for invalid categories (fail fast principle)
 */
export function mapOutlookCategory(category: string): string {
  // Input validation - fail fast on invalid input
  if (!category || typeof category !== 'string') {
    throw new Error(`Invalid category: expected non-empty string, got ${typeof category}`);
  }

  const trimmed = category.trim();
  if (trimmed === '') {
    throw new Error('Invalid category: empty string after trimming');
  }

  // Fail fast on unknown categories
  const normalizedCategory = trimmed.toLowerCase();
  const systemCategory = OUTLOOK_CATEGORIES[normalizedCategory as keyof typeof OUTLOOK_CATEGORIES];

  if (!systemCategory) {
    throw new Error(`Invalid Outlook category: "${category}". Valid categories: ${Object.keys(OUTLOOK_CATEGORIES).join(', ')}`);
  }

  return systemCategory;
}

export interface ODataFilterResult {
  search: string | null;
  filter: string | null;
  requireBodyClientFilter: boolean;
  hasFullText: boolean;
}

export function toGraphFilter(query: QueryNode): ODataFilterResult {
  if (query.kqlQuery) {
    return {
      search: query.kqlQuery,
      filter: null,
      requireBodyClientFilter: false,
      hasFullText: true,
    };
  }

  if (hasFullTextIntent(query)) {
    const searchCollect = collectSearchTerms(query);
    const search = searchCollect.terms.length ? searchCollect.terms.join(' OR ') : null;
    return {
      search,
      filter: null,
      requireBodyClientFilter: searchCollect.usesBody || searchCollect.usesText,
      hasFullText: true,
    };
  }

  const filterStr = buildQueryFilter(query);
  const cleanedFilter = filterStr && typeof filterStr === 'string' && filterStr.trim().length ? filterStr.trim() : null;
  return {
    search: null,
    filter: cleanedFilter,
    requireBodyClientFilter: false,
    hasFullText: false,
  };
}

function collectSearchTerms(query: QueryNode): { terms: string[]; usesText: boolean; usesBody: boolean; usesExact: boolean } {
  const state = {
    terms: [] as string[],
    usesText: false,
    usesBody: false,
    usesExact: false,
  };

  walkNode(query);
  return state;

  function walkNode(node: QueryNode | undefined) {
    if (!node || typeof node !== 'object') return;

    if (node.$and) {
      node.$and.forEach((child) => walkNode(child));
      return;
    }
    if (node.$or) {
      node.$or.forEach((child) => walkNode(child));
      return;
    }
    if (node.$not) {
      walkNode(node.$not);
      return;
    }

    if ('exactPhrase' in node && node.exactPhrase) {
      state.usesExact = true;
      pushExactPhrase(node.exactPhrase);
    }

    pushOperatorTerms(node.subject);

    if ('text' in node && node.text) {
      state.usesText = true;
      pushOperatorTerms(node.text);
    }
    if ('body' in node && node.body) {
      state.usesBody = true;
      pushOperatorTerms(node.body);
    }

    // Other fields (hasAttachment, importance, date) are handled via OData filters (emit())
    // and should not create search terms with colon prefixes that break Graph $search.
  }

  function pushExactPhrase(value: string) {
    const phrase = String(value);
    if (!phrase.trim()) return;
    pushAnyTerms([`"${phrase}"`]);
  }

  function pushOperatorTerms(value: FieldOperator | string | undefined) {
    if (!value) return;
    if (typeof value === 'string') {
      const clause = buildClause(null, value);
      if (clause) state.terms.push(clause);
      return;
    }

    if (Array.isArray(value.$any) && value.$any.length > 0) {
      pushAnyTerms(value.$any);
    }
    if (Array.isArray(value.$all) && value.$all.length > 0) {
      pushAllTerms(value.$all);
    }
    if (Array.isArray(value.$none) && value.$none.length > 0) {
      const clauses = value.$none.map((term) => buildClause(null, String(term ?? ''))).filter(Boolean);
      if (clauses.length === 1) {
        state.terms.push(`NOT ${clauses[0]}`);
      } else if (clauses.length > 1) {
        state.terms.push(`NOT (${clauses.join(' OR ')})`);
      }
    }
  }

  function pushAnyTerms(values: string[]) {
    const clauses = values.map((val) => buildClause(null, val)).filter(Boolean);
    if (!clauses.length) return;
    state.terms.push(clauses.length === 1 ? clauses[0] : `(${clauses.join(' OR ')})`);
  }

  function pushAllTerms(values: string[]) {
    const clauses = values.map((val) => buildClause(null, val)).filter(Boolean);
    if (!clauses.length) return;
    state.terms.push(clauses.length === 1 ? clauses[0] : `(${clauses.join(' AND ')})`);
  }

  function escapeKQL(value: string): string {
    return value
      .replace(/\\/g, '\\\\')
      .replace(/:/g, '\\:')
      .replace(/\(/g, '\\(')
      .replace(/\)/g, '\\)')
      .replace(/\{/g, '\\{')
      .replace(/\}/g, '\\}')
      .replace(/\[/g, '\\[')
      .replace(/\]/g, '\\]')
      .replace(/"/g, '\\"')
      .replace(/\*/g, '\\*')
      .replace(/\?/g, '\\?')
      .replace(/</g, '\\<')
      .replace(/>/g, '\\>')
      .replace(/_/g, '\\_');
  }

  function buildClause(_prefix: string | null, value: string): string {
    const trimmed = value.trim();
    if (!trimmed) return '';
    const escaped = escapeKQL(trimmed);
    const shouldQuote = /[^a-zA-Z\s]/.test(trimmed) || /\s/.test(trimmed);
    return shouldQuote ? `"${escaped}"` : escaped;
  }
}

function hasFullTextIntent(node: QueryNode | undefined | null): boolean {
  if (!node || typeof node !== 'object') return false;

  if (node.kqlQuery) return true;
  if (node.exactPhrase) return true;
  if ('subject' in node && node.subject) return true;
  if ('text' in node && node.text) return true;
  if ('body' in node && node.body) return true;

  if ('$and' in node && node.$and) {
    return node.$and.some((child) => hasFullTextIntent(child));
  }
  if ('$or' in node && node.$or) {
    return node.$or.some((child) => hasFullTextIntent(child));
  }
  if ('$not' in node && node.$not) {
    return hasFullTextIntent(node.$not);
  }

  return false;
}

export function buildQueryFilter(query: QueryNode): string {
  function L(s: string) {
    return s;
  }

  function p(s: string) {
    return `(${s})`;
  }

  function recip(coll: string, V: string) {
    const left = `r/emailAddress/address eq ${V}`;
    const right = `r/emailAddress/name eq ${V}`;
    return `${coll}/any(r: ${left} or ${right})`;
  }

  function fv(field: string, raw: string | number | boolean | null | undefined): string {
    if (raw == null) return '';
    const rawStr = String(raw);

    if (field === 'text' || field === 'body') {
      if (rawStr.trim() === '') {
        throw new Error(`Invalid ${field} value: empty string`);
      }
      return '';
    }

    if (rawStr.trim() === '') {
      throw new Error(`Invalid ${field} value: empty string`);
    }

    const V = `'${rawStr.toLowerCase().replace(/'/g, "''")}'`;
    switch (field) {
      case 'from': {
        if (rawStr.includes('@')) {
          const addrEq = `${L('from/emailAddress/address')} eq ${V}`;
          const nameEq = `${L('from/emailAddress/name')} eq ${V}`;
          return p(`${addrEq} or ${nameEq}`);
        }
        return '';
      }
      case 'to':
        return recip('toRecipients', V);
      case 'cc':
        return recip('ccRecipients', V);
      case 'bcc':
        return recip('bccRecipients', V);
      case 'subject':
        if (rawStr.includes('-') || rawStr.length < 3) return '';
        return `startswith(subject, ${V})`;
      case 'categories': {
        const systemCategory = mapOutlookCategory(rawStr);
        return `categories/any(c: c eq '${systemCategory}')`;
      }
      case 'label':
        return `categories/any(c: c eq '${rawStr.replace(/'/g, "''")}')`;
      default:
        return '';
    }
  }

  function chain(op: 'and' | 'or', arr: Array<string | undefined>) {
    if (!arr || arr.length === 0) {
      throw new Error(`chain: empty array for ${op} operation`);
    }
    if (arr.length === 1) {
      const first = arr[0] ?? '';
      return String(first);
    }
    const joined = arr.join(` ${op} `);
    return p(joined);
  }

  function fieldExpr(field: string, op: FieldOperator | string): string {
    if (typeof op === 'string') {
      return fv(field, op);
    }

    if (Array.isArray(op.$any)) {
      const results = op.$any.map((v) => fv(field, String(v ?? '')));
      const validResults = results.filter((s) => s && s.trim());
      if (validResults.length === 0) return '';
      return chain('or', validResults);
    }
    if (Array.isArray(op.$all)) {
      const results = op.$all.map((v) => fv(field, String(v ?? '')));
      const validResults = results.filter((s) => s && s.trim());
      if (validResults.length === 0) return '';
      return chain('and', validResults);
    }
    if (Array.isArray(op.$none)) {
      const results = op.$none.map((v) => fv(field, String(v ?? '')));
      const validResults = results.filter((s) => s && s.trim());
      if (validResults.length === 0) return '';
      return `not ${p(chain('or', validResults))}`;
    }
    throw new Error(`Unknown field operator ${JSON.stringify(op)}`);
  }

  function dateExpr(d: DateOperator): string {
    const parts: string[] = [];
    if (d.$gte) parts.push(`receivedDateTime ge ${d.$gte}T00:00:00Z`);
    if (d.$lt) parts.push(`receivedDateTime lt ${d.$lt}T00:00:00Z`);
    return parts.length > 1 ? p(parts.join(' and ')) : (parts[0] ?? '');
  }

  function emit(n: QueryNode): string {
    if (!n || typeof n !== 'object') return '';

    if ('$and' in n && n.$and) {
      const andParts = n.$and.map(emit).filter((part: string): part is string => Boolean(part));
      if (andParts.length === 0) return '';
      if (andParts.length === 1) {
        const firstPart = andParts[0];
        return firstPart ?? '';
      }
      return p(andParts.join(' and '));
    }
    if ('$or' in n && n.$or) {
      const orParts = n.$or.map(emit).filter((part: string): part is string => Boolean(part));
      if (orParts.length === 0) return '';
      if (orParts.length === 1) {
        const firstPart = orParts[0];
        return firstPart ?? '';
      }
      return p(orParts.join(' or '));
    }
    if ('$not' in n && n.$not) {
      const notExpr = emit(n.$not);
      return notExpr ? `not (${notExpr})` : '';
    }
    if ('hasAttachment' in n && n.hasAttachment) return 'hasAttachments eq true';
    if ('date' in n && n.date) return dateExpr(n.date);

    const keys = Object.keys(n);
    if (keys.length === 1) {
      const k0 = String(keys[0] ?? '');
      if (k0 && ['from', 'to', 'cc', 'bcc', 'subject', 'text', 'body', 'categories', 'label'].includes(k0)) {
        const fieldValue = (n as FieldQuery)[k0 as keyof FieldQuery];
        if (fieldValue) return fieldExpr(k0, fieldValue);
      }
    }
    return '';
  }

  return emit(query);
}
export function toOutlookFilter(parsed: QueryNode, _options?: { includeBodyContent?: boolean; useCaseInsensitiveWrap?: boolean }) {
  const graphResult = toGraphFilter(parsed);
  const filters = extractOutlookFilters(parsed);
  return {
    filter: graphResult.filter,
    search: graphResult.search,
    requireBodyClientFilter: graphResult.requireBodyClientFilter,
    filters,
  };
}

interface ExtractedFilters {
  subjectIncludes: string[];
  bodyIncludes: string[];
  textIncludes: string[];
  fromIncludes: string[];
  toIncludes: string[];
  ccIncludes: string[];
  bccIncludes: string[];
  categoriesIncludes: string[];
  labelIncludes: string[];
  hasAttachment?: boolean;
  since?: string;
  before?: string;
}

export function extractOutlookFilters(parsed: QueryNode): ExtractedFilters {
  const filters: ExtractedFilters = {
    subjectIncludes: [],
    bodyIncludes: [],
    textIncludes: [],
    fromIncludes: [],
    toIncludes: [],
    ccIncludes: [],
    bccIncludes: [],
    categoriesIncludes: [],
    labelIncludes: [],
  };
  function walk(node: QueryNode): void {
    if (!node || typeof node !== 'object') return;

    if ('$and' in node && node.$and) {
      node.$and.forEach(walk);
      return;
    }
    if ('$or' in node && node.$or) {
      node.$or.forEach(walk);
      return;
    }
    if ('$not' in node && node.$not) {
      walk(node.$not);
      return;
    }

    if ('hasAttachment' in node && node.hasAttachment !== undefined) {
      filters.hasAttachment = node.hasAttachment === true;
      return;
    }
    if ('date' in node && node.date) {
      if (node.date.$gte) filters.since = node.date.$gte;
      if (node.date.$lt) filters.before = node.date.$lt;
      return;
    }

    if ('subject' in node && node.subject) {
      if (typeof node.subject === 'string') {
        filters.subjectIncludes.push(node.subject);
      } else {
        if (node.subject.$any) filters.subjectIncludes.push(...node.subject.$any);
        if (node.subject.$all) filters.subjectIncludes.push(...node.subject.$all);
      }
    }
    if ('body' in node && node.body) {
      if (typeof node.body === 'string') {
        filters.bodyIncludes.push(node.body);
      } else {
        if (node.body.$any) filters.bodyIncludes.push(...node.body.$any);
        if (node.body.$all) filters.bodyIncludes.push(...node.body.$all);
      }
    }
    if ('text' in node && node.text) {
      if (typeof node.text === 'string') {
        filters.textIncludes.push(node.text);
        filters.bodyIncludes.push(node.text);
      } else {
        if (node.text.$any) {
          filters.textIncludes.push(...node.text.$any);
          filters.bodyIncludes.push(...node.text.$any);
        }
        if (node.text.$all) {
          filters.textIncludes.push(...node.text.$all);
          filters.bodyIncludes.push(...node.text.$all);
        }
      }
    }
    if ('from' in node && node.from) {
      if (typeof node.from === 'string') {
        filters.fromIncludes.push(node.from);
      } else {
        if (node.from.$any) filters.fromIncludes.push(...node.from.$any);
        if (node.from.$all) filters.fromIncludes.push(...node.from.$all);
      }
    }
    if ('to' in node && node.to) {
      if (typeof node.to === 'string') {
        filters.toIncludes.push(node.to);
      } else {
        if (node.to.$any) filters.toIncludes.push(...node.to.$any);
        if (node.to.$all) filters.toIncludes.push(...node.to.$all);
      }
    }
    if ('cc' in node && node.cc) {
      if (typeof node.cc === 'string') {
        filters.ccIncludes.push(node.cc);
      } else {
        if (node.cc.$any) filters.ccIncludes.push(...node.cc.$any);
        if (node.cc.$all) filters.ccIncludes.push(...node.cc.$all);
      }
    }
    if ('bcc' in node && node.bcc) {
      if (typeof node.bcc === 'string') {
        filters.bccIncludes.push(node.bcc);
      } else {
        if (node.bcc.$any) filters.bccIncludes.push(...node.bcc.$any);
        if (node.bcc.$all) filters.bccIncludes.push(...node.bcc.$all);
      }
    }
    if ('categories' in node && node.categories) {
      if (typeof node.categories === 'string') {
        if (mapOutlookCategory(node.categories) !== null) {
          filters.categoriesIncludes.push(node.categories);
        }
      } else {
        // Filter out invalid categories to prevent memory leaks and maintain clean state
        if (node.categories.$any) {
          const validCategories = node.categories.$any.filter((cat: string) => mapOutlookCategory(cat) !== null);
          if (validCategories.length > 0) filters.categoriesIncludes.push(...validCategories);
        }
        if (node.categories.$all) {
          const validCategories = node.categories.$all.filter((cat: string) => mapOutlookCategory(cat) !== null);
          if (validCategories.length > 0) filters.categoriesIncludes.push(...validCategories);
        }
        if (node.categories.$none) {
          const validCategories = node.categories.$none.filter((cat: string) => mapOutlookCategory(cat) !== null);
          if (validCategories.length > 0) filters.categoriesIncludes.push(...validCategories);
        }
      }
    }
    if ('label' in node && node.label) {
      if (typeof node.label === 'string') {
        filters.labelIncludes.push(node.label);
      } else {
        if (node.label.$any) filters.labelIncludes.push(...node.label.$any);
        if (node.label.$all) filters.labelIncludes.push(...node.label.$all);
        if (node.label.$none) filters.labelIncludes.push(...node.label.$none);
      }
    }
  }
  walk(parsed);
  return filters;
}
