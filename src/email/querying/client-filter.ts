import type { FieldOperator } from '@mcp-z/email';
import type { Message, Recipient } from '@microsoft/microsoft-graph-types';
import type { OutlookQuery } from '../../schemas/outlook-query-schema.ts';
import { mapOutlookCategory } from './query-builder.ts';

type Predicate = (message: Message) => boolean;

export function buildClientPredicate(query: OutlookQuery): Predicate {
  const nodePredicate = buildNodePredicate(query);
  return (message) => nodePredicate(message);
}

export function needsBody(query: OutlookQuery): boolean {
  return hasBodyTextOrExact(query);
}

function buildNodePredicate(node: OutlookQuery | undefined | null): Predicate {
  if (!node || typeof node !== 'object') return () => true;

  if (node.$and) {
    const children = node.$and.map(buildNodePredicate);
    return (message) => children.every((child) => child(message));
  }
  if (node.$or) {
    const children = node.$or.map(buildNodePredicate);
    return (message) => children.some((child) => child(message));
  }
  if (node.$not) {
    const child = buildNodePredicate(node.$not);
    return (message) => !child(message);
  }

  const predicates: Predicate[] = [];

  if ('hasAttachment' in node && typeof node.hasAttachment === 'boolean') {
    predicates.push((message) => message.hasAttachments === node.hasAttachment);
  }
  if ('isRead' in node && typeof node.isRead === 'boolean') {
    predicates.push((message) => message.isRead === node.isRead);
  }
  if ('importance' in node && node.importance) {
    const expected = node.importance.toLowerCase();
    predicates.push((message) => typeof message.importance === 'string' && message.importance.toLowerCase() === expected);
  }
  if ('date' in node && node.date) {
    const { $gte, $lt } = node.date;
    predicates.push((message) => matchesDate(message.receivedDateTime, $gte, $lt));
  }

  if ('from' in node && node.from) {
    predicates.push((message) => matchesAddressField(normalizeAddress(message.from), node.from));
  }
  if ('to' in node && node.to) {
    predicates.push((message) => matchesRecipientField(message.toRecipients, node.to));
  }
  if ('cc' in node && node.cc) {
    predicates.push((message) => matchesRecipientField(message.ccRecipients, node.cc));
  }
  if ('bcc' in node && node.bcc) {
    predicates.push((message) => matchesRecipientField(message.bccRecipients, node.bcc));
  }
  if ('subject' in node && node.subject) {
    predicates.push((message) => matchesSubject(message, node.subject));
  }
  if ('text' in node && node.text) {
    predicates.push((message) => matchesText(message, node.text));
  }
  if ('body' in node && node.body) {
    predicates.push((message) => matchesBody(message, node.body));
  }
  if ('exactPhrase' in node && node.exactPhrase) {
    predicates.push((message) => matchesExactPhrase(message, node.exactPhrase));
  }
  if ('categories' in node && node.categories) {
    predicates.push((message) => matchesCategoryField(message.categories ?? [], node.categories));
  }
  if ('label' in node && node.label) {
    predicates.push((message) => matchesLabelField(message.categories ?? [], node.label));
  }

  if (predicates.length === 0) return () => true;
  return (message) => predicates.every((predicate) => predicate(message));
}

function hasBodyTextOrExact(node: OutlookQuery | undefined | null): boolean {
  if (!node || typeof node !== 'object') return false;
  if ('body' in node && node.body) return true;
  if ('text' in node && node.text) return true;
  if ('exactPhrase' in node && node.exactPhrase) return true;
  if ('$and' in node && node.$and) {
    return node.$and.some((child) => hasBodyTextOrExact(child));
  }
  if ('$or' in node && node.$or) {
    return node.$or.some((child) => hasBodyTextOrExact(child));
  }
  if ('$not' in node && node.$not) {
    return hasBodyTextOrExact(node.$not);
  }
  return false;
}

function matchesSubject(message: Message, value: string | FieldOperator): boolean {
  const normalizedSubject = normalizeForMatch(message.subject);
  return evaluateOperator(value, (term) => normalizedSubject.includes(term));
}

function matchesText(message: Message, value: string | FieldOperator): boolean {
  const textSources = getTextSources(message);
  return evaluateOperator(value, (term) => textSources.some((source) => source.includes(term)));
}

function matchesBody(message: Message, value: string | FieldOperator): boolean {
  const bodyContent = normalizeForMatch(message.body?.content);
  return evaluateOperator(value, (term) => bodyContent.includes(term));
}

function matchesExactPhrase(message: Message, phrase: string): boolean {
  const normalizedPhrase = normalizeTerm(phrase);
  if (!normalizedPhrase) return false;
  const textSources = getTextSources(message);
  return textSources.some((source) => source.includes(normalizedPhrase));
}

function matchesAddressField(addressData: NormalizedAddress, value: string | FieldOperator): boolean {
  const matchTerm = (term: string) => {
    const addressMatch = addressData.address.includes(term);
    const nameMatch = addressData.name.includes(term);
    if (term.includes('@')) {
      return addressMatch || nameMatch;
    }
    return addressMatch || nameMatch;
  };
  return evaluateOperator(value, matchTerm);
}

type NormalizedAddress = { address: string; name: string };

function normalizeAddress(recipient: Recipient | undefined): NormalizedAddress {
  return {
    address: normalizeForMatch(recipient?.emailAddress?.address),
    name: normalizeForMatch(recipient?.emailAddress?.name),
  };
}

function matchesRecipientField(recipients: Recipient[] | undefined, value: string | FieldOperator): boolean {
  const normalizedRecipients = (recipients ?? []).map((recipient) => ({
    address: normalizeForMatch(recipient?.emailAddress?.address),
    name: normalizeForMatch(recipient?.emailAddress?.name),
  }));
  const matchTerm = (term: string) => normalizedRecipients.some((recipient) => recipient.address.includes(term) || recipient.name.includes(term));
  return evaluateOperator(value, matchTerm);
}

function matchesCategoryField(categories: string[], value: string | FieldOperator): boolean {
  const mappedCategories = categories.map((category) => category ?? '');
  const matchTerm = (term: string) => {
    try {
      const mappedTerm = mapOutlookCategory(term);
      return mappedCategories.includes(mappedTerm);
    } catch {
      return false;
    }
  };
  return evaluateOperator(value, matchTerm);
}

function matchesLabelField(categories: string[], value: string | FieldOperator): boolean {
  const normalizeLabelTerm = (input: string): string | null => {
    const trimmed = input.trim();
    return trimmed === '' ? null : trimmed;
  };
  const matchTerm = (term: string) => categories.includes(term);
  return evaluateOperator(value, matchTerm, normalizeLabelTerm);
}

function matchesDate(received: string | undefined, gte?: string, lt?: string): boolean {
  if (!received) return false;
  const receivedTs = Date.parse(received);
  if (Number.isNaN(receivedTs)) return false;
  if (gte) {
    const start = Date.parse(`${gte}T00:00:00Z`);
    if (Number.isNaN(start) || receivedTs < start) return false;
  }
  if (lt) {
    const end = Date.parse(`${lt}T00:00:00Z`);
    if (Number.isNaN(end) || receivedTs >= end) return false;
  }
  return true;
}

function evaluateOperator(value: string | FieldOperator, matcher: (normalizedTerm: string) => boolean, normalizeFn: (input: string) => string = normalizeTerm): boolean {
  if (typeof value === 'string') {
    const normalized = normalizeFn(value);
    return normalized !== null && matcher(normalized);
  }
  let hasConstraint = false;
  if (Array.isArray(value.$any) && value.$any.length > 0) {
    hasConstraint = true;
    const terms = value.$any.map(normalizeFn).filter((term): term is string => Boolean(term));
    if (terms.length === 0 || !terms.some(matcher)) return false;
  }
  if (Array.isArray(value.$all) && value.$all.length > 0) {
    hasConstraint = true;
    const terms = value.$all.map(normalizeFn).filter((term): term is string => Boolean(term));
    if (terms.length === 0 || !terms.every(matcher)) return false;
  }
  if (Array.isArray(value.$none) && value.$none.length > 0) {
    hasConstraint = true;
    const terms = value.$none.map(normalizeFn).filter((term): term is string => Boolean(term));
    if (terms.some(matcher)) return false;
  }
  return !!hasConstraint;
}

function normalizeTerm(input: string): string | null {
  const collapsed = input.replace(/\s+/g, ' ').trim().toLowerCase();
  return collapsed === '' ? null : collapsed;
}

function normalizeForMatch(value?: string): string {
  return normalizeTerm(value ?? '') ?? '';
}

function getTextSources(message: Message): string[] {
  return [normalizeForMatch(message.subject), normalizeForMatch(message.body?.content), normalizeForMatch(message.bodyPreview)];
}
