export function extractEmailsFromRecipients(recipients: Array<{ emailAddress?: { address?: string; name?: string } }>, mode: 'email' | 'name' | 'raw' = 'email'): string[] {
  if (!Array.isArray(recipients)) return [];
  return recipients
    .map((r) => {
      const name = r?.emailAddress?.name ?? '';
      const email = r?.emailAddress?.address ?? '';
      return mode === 'email' ? email : mode === 'name' ? name || email : `${name} <${email}>`;
    })
    .filter(Boolean);
}

export function formatAddressList(recipients: Array<{ emailAddress?: { address?: string; name?: string } }>, mode: 'email' | 'name' | 'raw' = 'email'): string {
  return extractEmailsFromRecipients(recipients, mode).join(', ');
}

export function extractFrom(from: { emailAddress?: { address?: string; name?: string } } | undefined): string | undefined {
  if (!from) return undefined;
  return from.emailAddress?.address || from.emailAddress?.name || undefined;
}
