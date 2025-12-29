import { type BaseEmailQueryFields, baseEmailQueryFields, type FieldOperator, FieldOperatorSchema } from '@mcp-z/email';
import { z } from 'zod';

/**
 * Outlook system category values.
 * These are the built-in category types that Outlook recognizes.
 * Note: This is distinct from OutlookCategory (user-created category objects with id/displayName/color).
 */
export type OutlookSystemCategory = 'work' | 'personal' | 'family' | 'travel' | 'important' | 'urgent';

/** Zod schema for validating OutlookSystemCategory values */
export const OutlookSystemCategorySchema = z.enum(['work', 'personal', 'family', 'travel', 'important', 'urgent']);

/**
 * Outlook-specific query schema with recursive operators and Outlook features.
 *
 * Includes Microsoft Graph/Outlook-specific features:
 * - exactPhrase: Strict exact phrase matching via KQL $search parameter
 * - categories: Outlook system categories (work, personal, family, travel, important, urgent)
 * - label: User-created Outlook categories (case-sensitive, discovered via outlook-categories-list)
 * - importance: Message importance level (high, normal, low)
 * - kqlQuery: Escape hatch for advanced KQL (Keyword Query Language) syntax
 *
 * Plus all base fields from baseEmailQueryFields:
 * - Email addresses: from, to, cc, bcc (support string or field operators)
 * - Content: subject, body, text
 * - Flags: hasAttachment, isRead
 * - Date range: date { $gte, $lt }
 * - Logical operators: $and, $or, $not (recursive)
 *
 * Note: Cast through unknown to work around Zod's lazy schema type inference issue
 * with exactOptionalPropertyTypes. The runtime schema is correct; this cast ensures
 * TypeScript sees the strict OutlookQuery type everywhere the schema is used.
 */
export const OutlookQuerySchema = z.lazy(() =>
  z
    .object({
      // Logical operators for combining conditions (recursive)
      $and: z.array(OutlookQuerySchema).optional().describe('Array of conditions that must ALL match'),
      $or: z.array(OutlookQuerySchema).optional().describe('Array of conditions where ANY must match'),
      $not: OutlookQuerySchema.optional().describe('Nested condition that must NOT match'),

      // Spread base email query fields (from, to, subject, body, etc.)
      ...baseEmailQueryFields,

      // Outlook-specific features

      // Exact phrase matching - KQL strict search with double quotes
      exactPhrase: z.string().min(1).optional().describe('Exact phrase matching - words must appear together in exact order (strict matching). Outlook uses KQL.'),

      // Outlook system categories with field operator support
      categories: z
        .union([
          OutlookSystemCategorySchema,
          z
            .object({
              $any: z.array(OutlookSystemCategorySchema).optional(),
              $all: z.array(OutlookSystemCategorySchema).optional(),
              $none: z.array(OutlookSystemCategorySchema).optional(),
            })
            .strict(),
        ])
        .optional()
        .describe('Filter by Outlook system categories (work, personal, family, travel, important, urgent)'),

      // User-created categories
      label: z
        .union([z.string().min(1), FieldOperatorSchema])
        .optional()
        .describe('Filter by user-created categories (case-sensitive). Use outlook-categories-list to see available categories'),

      // Message importance level - Outlook-specific property
      importance: z.enum(['high', 'normal', 'low']).optional().describe('Filter by message importance level (high, normal, low)'),

      // Raw KQL query string - escape hatch for advanced syntax
      kqlQuery: z.string().min(1).optional().describe('Raw KQL (Keyword Query Language) syntax for advanced use cases. Bypasses schema validation - use sparingly.'),
    })
    .strict()
) as unknown as z.ZodType<OutlookQuery>;

export type OutlookQuery = BaseEmailQueryFields & {
  $and?: OutlookQuery[];
  $or?: OutlookQuery[];
  $not?: OutlookQuery;
  exactPhrase?: string;
  categories?:
    | OutlookSystemCategory
    | {
        $any?: OutlookSystemCategory[];
        $all?: OutlookSystemCategory[];
        $none?: OutlookSystemCategory[];
      };
  label?: string | FieldOperator;
  importance?: 'high' | 'normal' | 'low';
  kqlQuery?: string;
};
