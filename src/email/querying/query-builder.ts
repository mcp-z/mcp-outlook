import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.js';

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
const OUTLOOK_CATEGORIES = {
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
function mapOutlookCategory(category: string): string {
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
}

export function toGraphFilter(query: QueryNode): ODataFilterResult {
  let requireBodyClientFilter = false;

  function L(s: string) {
    // Microsoft Graph does NOT support toLower() function in OData filters
    // Case-insensitive matching must be done client-side or via $search
    return s;
  }
  function p(s: string) {
    return `(${s})`;
  }

  function recip(coll: string, V: string) {
    // Graph may reject functions on nested properties; use equality on nested address/name fields.
    const left = `r/emailAddress/address eq ${V}`;
    const right = `r/emailAddress/name eq ${V}`;
    return `${coll}/any(r: ${left} or ${right})`;
  }

  function fv(field: string, raw: string | number | boolean | null | undefined): string {
    if (raw == null) return '';
    const rawStr = String(raw);

    // For text and body fields, skip OData filtering entirely (uses KQL search instead)
    // Do this BEFORE validation so we don't throw on valid text/body queries
    if (field === 'text' || field === 'body') {
      if (rawStr.trim() === '') {
        throw new Error(`Invalid ${field} value: empty string`);
      }
      requireBodyClientFilter = true;
      return '';
    }

    // For all other fields, validate non-empty AFTER checking for null
    if (rawStr.trim() === '') {
      throw new Error(`Invalid ${field} value: empty string`);
    }

    const V = `'${rawStr.toLowerCase().replace(/'/g, "''")}'`;
    switch (field) {
      case 'from': {
        // For email addresses emit equality checks on nested address/name;
        // for name-only tokens, skip server-side filter (handled client-side or via $search).
        if (rawStr.includes('@')) {
          const addrEq = `${L('from/emailAddress/address')} eq ${V}`;
          const nameEq = `${L('from/emailAddress/name')} eq ${V}`;
          return p(`${addrEq} or ${nameEq}`);
        }
        // Don't emit a server-side filter for name-only 'from' tokens.
        return '';
      }
      case 'to':
        return recip('toRecipients', V);
      case 'cc':
        return recip('ccRecipients', V);
      case 'bcc':
        return recip('bccRecipients', V);
      case 'subject':
        // Microsoft Graph does NOT support contains() function in OData filters
        // Use startswith as a partial alternative or rely on $search
        if (rawStr.includes('-') || rawStr.length < 3) return '';
        return `startswith(subject, ${V})`;
      case 'categories': {
        // mapOutlookCategory will throw on invalid input (fail fast)
        const systemCategory = mapOutlookCategory(rawStr);
        // Use the categories OData property with exact case-sensitive match
        return `categories/any(c: c eq '${systemCategory}')`;
      }
      case 'label': {
        // Direct passthrough to categories for Outlook (case-sensitive)
        return `categories/any(c: c eq '${rawStr.replace(/'/g, "''")}')`;
      }
      default:
        // Microsoft Graph does NOT support contains() function
        // For unknown fields, skip server-side filtering
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
      // Filter out empty strings (from fields like text/body that don't generate OData filters)
      const validResults = results.filter((s) => s && s.trim());
      if (validResults.length === 0) return '';
      return chain('or', validResults);
    }
    if (Array.isArray(op.$all)) {
      const results = op.$all.map((v) => fv(field, String(v ?? '')));
      // Filter out empty strings (from fields like text/body that don't generate OData filters)
      const validResults = results.filter((s) => s && s.trim());
      if (validResults.length === 0) return '';
      return chain('and', validResults);
    }
    if (Array.isArray(op.$none)) {
      const results = op.$none.map((v) => fv(field, String(v ?? '')));
      // Filter out empty strings (from fields like text/body that don't generate OData filters)
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

  const filterStr = emit(query);

  const terms: string[] = [];
  function pushTerms(arr?: string[]) {
    if (Array.isArray(arr)) for (const t of arr) if ((t ?? '').toString().trim()) terms.push(String(t).trim());
  }

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

    if ('exactPhrase' in node && node.exactPhrase) {
      // KQL uses double quotes for exact phrase matching
      pushTerms([`"${String(node.exactPhrase).replace(/"/g, '\\"')}"`]);
      return;
    }

    if ('subject' in node && node.subject) {
      if (typeof node.subject === 'string') {
        pushTerms([node.subject]);
      } else {
        pushTerms(node.subject.$any);
        pushTerms(node.subject.$all);
        pushTerms(node.subject.$none);
      }
    }
    if ('body' in node && node.body) {
      if (typeof node.body === 'string') {
        pushTerms([node.body]);
      } else {
        pushTerms(node.body.$any);
        pushTerms(node.body.$all);
        pushTerms(node.body.$none);
      }
    }
    if ('text' in node && node.text) {
      if (typeof node.text === 'string') {
        pushTerms([node.text]);
      } else {
        pushTerms(node.text.$any);
        pushTerms(node.text.$all);
        pushTerms(node.text.$none);
      }
    }
  }
  walk(query);

  /**
   * Escape KQL (Keyword Query Language) special characters in search terms.
   * KQL special characters that need escaping: \ : ( ) { } [ ] " * ? < > - _
   * According to Microsoft Graph KQL syntax, these must be escaped with backslash.
   */
  function escapeKQL(term: string): string {
    if (!term) return term;
    // Escape backslash first, then other special characters
    return term
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
      .replace(/-/g, '\\-')
      .replace(/_/g, '\\_');
  }

  const mapped = terms.map((t) => {
    const escaped = escapeKQL(String(t));
    return escaped;
  });
  const search = mapped.length ? mapped.join(' OR ') : null;

  return { search, filter: filterStr && typeof filterStr === 'string' && filterStr.length ? filterStr : null, requireBodyClientFilter };
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
