/**
 * Outlook-specific message builder for composing emails using Microsoft Graph API format
 * This is the Outlook equivalent of Gmail's RFC822 builder
 */

export interface OutlookRecipient {
  emailAddress: {
    address: string;
    name?: string;
  };
}

export interface OutlookMessageArgs {
  to?: string | string[] | undefined;
  cc?: string | string[] | undefined;
  bcc?: string | string[] | undefined;
  subject?: string;
  body?: string;
  from?: string;
  importance?: 'Low' | 'Normal' | 'High';
  contentType?: 'Text' | 'HTML';
}

export interface OutlookMessage {
  subject?: string;
  body?: {
    contentType: string;
    content: string;
  };
  toRecipients?: OutlookRecipient[];
  ccRecipients?: OutlookRecipient[];
  bccRecipients?: OutlookRecipient[];
  importance?: string;
}

/**
 * Build an Outlook message object from email arguments for Microsoft Graph API
 * This is the Outlook equivalent of buildRfc822FromArgs for Gmail
 */
export function buildOutlookMessage(args: OutlookMessageArgs): OutlookMessage {
  const { to = [], cc = [], bcc = [], subject = '', body = '', importance = 'Normal', contentType = 'Text' } = args;

  const message: OutlookMessage = {};

  if (subject) {
    message.subject = subject;
  }

  if (body) {
    message.body = {
      contentType: contentType,
      content: body,
    };
  }

  function parseEmailAddresses(emails: string | string[]): OutlookRecipient[] {
    const emailArray = Array.isArray(emails) ? emails : [emails];
    return emailArray
      .filter((email) => email && email.trim())
      .map((email) => {
        const trimmed = email.trim();
        // Simple email parsing - could be enhanced for name parsing
        return {
          emailAddress: {
            address: trimmed,
          },
        };
      });
  }

  if (to && (Array.isArray(to) ? to.length > 0 : to.trim())) {
    message.toRecipients = parseEmailAddresses(to);
  }

  if (cc && (Array.isArray(cc) ? cc.length > 0 : cc.trim())) {
    message.ccRecipients = parseEmailAddresses(cc);
  }

  if (bcc && (Array.isArray(bcc) ? bcc.length > 0 : bcc.trim())) {
    message.bccRecipients = parseEmailAddresses(bcc);
  }

  if (importance && importance !== 'Normal') {
    message.importance = importance;
  }

  return message;
}
