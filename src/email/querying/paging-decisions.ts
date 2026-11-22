import { encodeNextPageToken } from './pagination-token.ts';

/** Token that resumes inside the page fetched by `link`, one message past the last one taken. */
export function midPageToken(params: { link?: string; pageStart: number; index: number; mode: 'search' | 'filter' }): string {
  return encodeNextPageToken({ ...(params.link && { link: params.link }), offset: params.pageStart + params.index + 1, mode: params.mode });
}

/** Which collection a page token resumes: a fallback-filter token resumes the filter collection, anything else the search. */
export function resumeCollection(mode: string | undefined): 'search' | 'filter' {
  return mode === 'filter' ? 'filter' : 'search';
}

/** Cut a batch to what still fits under maxItems; truncated when rows were cut or a link remains at the limit. */
export function cutToMaxItems<T>(rows: T[], taken: number, maxItems: number, hasMore: boolean): { rows: T[]; truncated: boolean } {
  const kept = rows.slice(0, Math.max(0, maxItems - taken));
  const truncated = kept.length < rows.length || (taken + kept.length >= maxItems && hasMore);
  return { rows: kept, truncated };
}
