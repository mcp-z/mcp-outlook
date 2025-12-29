// Strict type for Outlook message structure used in mapping and row conversion
export interface OutlookRecipient {
  emailAddress?: {
    address?: string;
    name?: string;
  };
}

export interface OutlookMessage {
  id?: string;
  conversationId?: string;
  subject?: string;
  from?: OutlookRecipient;
  toRecipients?: OutlookRecipient[];
  ccRecipients?: OutlookRecipient[];
  bccRecipients?: OutlookRecipient[];
  receivedDateTime?: string;
  categories?: string[];
  bodyPreview?: string;
  body?: {
    content?: string;
    contentType?: string;
  };
}
