import { z } from 'zod';

// Base schemas for common Microsoft types
export const EmailAddressSchema = z.object({
  name: z.string().optional().describe('Display name of the recipient'),
  address: z.string().email().describe('Email address'),
});

export const RecipientSchema = z.object({
  emailAddress: EmailAddressSchema,
});

// Outlook Message schema based on Microsoft Graph API
export const OutlookMessageSchema = z.object({
  id: z.string().optional().describe('Unique identifier for the message'),
  subject: z.string().optional().describe('Subject of the email message'),
  body: z
    .object({
      contentType: z.enum(['text', 'html']).optional().describe('Content type of the body'),
      content: z.string().optional().describe('Body content of the message'),
    })
    .optional(),
  bodyPreview: z.string().optional().describe('First 255 characters of the message body'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('Importance level of the message'),
  hasAttachments: z.boolean().optional().describe('Indicates whether the message has attachments'),
  parentFolderId: z.string().optional().describe('ID of the parent folder'),
  sender: RecipientSchema.optional().describe('Account used to send the message'),
  from: RecipientSchema.optional().describe('Sender of the message'),
  toRecipients: z.array(RecipientSchema).optional().describe('To recipients of the message'),
  ccRecipients: z.array(RecipientSchema).optional().describe('CC recipients of the message'),
  bccRecipients: z.array(RecipientSchema).optional().describe('BCC recipients of the message'),
  replyTo: z.array(RecipientSchema).optional().describe('Reply-to recipients'),
  conversationId: z.string().optional().describe('ID of the conversation the email belongs to'),
  conversationIndex: z.string().optional().describe('Index for conversation threading'),
  receivedDateTime: z.string().optional().describe('ISO datetime when message was received'),
  sentDateTime: z.string().optional().describe('ISO datetime when message was sent'),
  createdDateTime: z.string().optional().describe('ISO datetime when message was created'),
  lastModifiedDateTime: z.string().optional().describe('ISO datetime when message was last modified'),
  changeKey: z.string().optional().describe('Version of the message'),
  categories: z.array(z.string()).optional().describe('Categories associated with the message'),
  isDeliveryReceiptRequested: z.boolean().optional().describe('Whether delivery receipt is requested'),
  isReadReceiptRequested: z.boolean().optional().describe('Whether read receipt is requested'),
  isRead: z.boolean().optional().describe('Whether the message has been read'),
  isDraft: z.boolean().optional().describe('Whether the message is a draft'),
  webLink: z.string().optional().describe('URL to open the message in Outlook on the web'),
  inferenceClassification: z.enum(['focused', 'other']).optional().describe('Classification for focused inbox'),
  flag: z
    .object({
      flagStatus: z.enum(['notFlagged', 'complete', 'flagged']).describe('Status of the flag'),
      startDateTime: z.string().optional().describe('Start date for the flag'),
      dueDateTime: z.string().optional().describe('Due date for the flag'),
    })
    .optional()
    .describe('Flag information for the message'),
});

// Outlook Folder schema
export const OutlookFolderSchema = z.object({
  id: z.string().optional().describe('Unique identifier for the folder'),
  displayName: z.string().optional().describe('Display name of the folder'),
  parentFolderId: z.string().optional().describe('ID of the parent folder'),
  childFolderCount: z.number().optional().describe('Number of immediate child folders'),
  unreadItemCount: z.number().optional().describe('Number of unread items in the folder'),
  totalItemCount: z.number().optional().describe('Total number of items in the folder'),
  sizeInBytes: z.number().optional().describe('Size of the folder in bytes'),
  isHidden: z.boolean().optional().describe('Whether the folder is hidden'),
  wellKnownName: z
    .enum(['archive', 'clutter', 'conflicts', 'conversationhistory', 'deleteditems', 'drafts', 'inbox', 'junkemail', 'localfailures', 'msgfolderroot', 'outbox', 'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'sentitems', 'serverfailures', 'syncissues'])
    .optional()
    .describe('Well-known folder name if applicable'),
});

// Outlook Attachment schema
export const OutlookAttachmentSchema = z.object({
  id: z.string().optional().describe('Unique identifier for the attachment'),
  name: z.string().optional().describe('Name of the attachment'),
  contentType: z.string().optional().describe('MIME type of the attachment'),
  size: z.number().optional().describe('Size of the attachment in bytes'),
  isInline: z.boolean().optional().describe('Whether the attachment is inline'),
  lastModifiedDateTime: z.string().optional().describe('ISO datetime when attachment was last modified'),
  contentId: z.string().optional().describe('Content ID for inline attachments'),
  contentLocation: z.string().optional().describe('Content location for inline attachments'),
  contentBytes: z.string().optional().describe('Base64-encoded content of the attachment'),
});

// OneDrive file schema (subset of Microsoft Graph DriveItem)
export const OneDriveFileSchema = z.object({
  id: z.string().optional().describe('Unique identifier for the file'),
  name: z.string().optional().describe('Name of the file'),
  size: z.number().optional().describe('Size of the file in bytes'),
  createdDateTime: z.string().optional().describe('ISO datetime when file was created'),
  lastModifiedDateTime: z.string().optional().describe('ISO datetime when file was last modified'),
  webUrl: z.string().optional().describe('URL to view the file in OneDrive'),
  downloadUrl: z.string().optional().describe('Direct download URL for the file'),
  parentReference: z
    .object({
      driveId: z.string().optional().describe('ID of the drive containing the file'),
      driveType: z.enum(['personal', 'business', 'documentLibrary']).optional().describe('Type of drive'),
      id: z.string().optional().describe('ID of the parent folder'),
      name: z.string().optional().describe('Name of the parent folder'),
      path: z.string().optional().describe('Path to the parent folder'),
    })
    .optional()
    .describe('Reference to the parent folder'),
  file: z
    .object({
      mimeType: z.string().optional().describe('MIME type of the file'),
      hashes: z
        .object({
          quickXorHash: z.string().optional().describe('QuickXorHash of the file'),
          sha1Hash: z.string().optional().describe('SHA1 hash of the file'),
          sha256Hash: z.string().optional().describe('SHA256 hash of the file'),
        })
        .optional()
        .describe('Hash values for the file'),
    })
    .optional()
    .describe('File-specific properties'),
  folder: z
    .object({
      childCount: z.number().optional().describe('Number of children in the folder'),
    })
    .optional()
    .describe('Folder-specific properties (present if item is a folder)'),
  shared: z
    .object({
      scope: z.enum(['anonymous', 'organization', 'users']).optional().describe('Scope of sharing'),
      owner: z
        .object({
          user: z
            .object({
              displayName: z.string().optional(),
              id: z.string().optional(),
            })
            .optional(),
        })
        .optional(),
    })
    .optional()
    .describe('Sharing information'),
  createdBy: z
    .object({
      user: z
        .object({
          displayName: z.string().optional(),
          id: z.string().optional(),
        })
        .optional(),
    })
    .optional()
    .describe('User who created the file'),
  lastModifiedBy: z
    .object({
      user: z
        .object({
          displayName: z.string().optional(),
          id: z.string().optional(),
        })
        .optional(),
    })
    .optional()
    .describe('User who last modified the file'),
});

// Microsoft Calendar Event schema
export const OutlookCalendarEventSchema = z.object({
  id: z.string().optional().describe('Unique identifier for the event'),
  subject: z.string().optional().describe('Subject of the event'),
  body: z
    .object({
      contentType: z.enum(['text', 'html']).optional().describe('Content type of the body'),
      content: z.string().optional().describe('Body content of the event'),
    })
    .optional(),
  start: z
    .object({
      dateTime: z.string().describe('Start date and time in ISO format'),
      timeZone: z.string().optional().describe('Time zone for the start time'),
    })
    .optional()
    .describe('Start time of the event'),
  end: z
    .object({
      dateTime: z.string().describe('End date and time in ISO format'),
      timeZone: z.string().optional().describe('Time zone for the end time'),
    })
    .optional()
    .describe('End time of the event'),
  location: z
    .object({
      displayName: z.string().optional().describe('Display name of the location'),
      address: z
        .object({
          street: z.string().optional(),
          city: z.string().optional(),
          state: z.string().optional(),
          countryOrRegion: z.string().optional(),
          postalCode: z.string().optional(),
        })
        .optional(),
      coordinates: z
        .object({
          latitude: z.number().optional(),
          longitude: z.number().optional(),
        })
        .optional(),
    })
    .optional()
    .describe('Location of the event'),
  attendees: z
    .array(
      z.object({
        type: z.enum(['required', 'optional', 'resource']).optional().describe('Type of attendee'),
        status: z
          .object({
            response: z.enum(['none', 'organizer', 'tentativelyAccepted', 'accepted', 'declined', 'notResponded']).optional(),
            time: z.string().optional().describe('Time when response was given'),
          })
          .optional(),
        emailAddress: EmailAddressSchema.optional(),
      })
    )
    .optional()
    .describe('Attendees of the event'),
  organizer: z
    .object({
      emailAddress: EmailAddressSchema.optional(),
    })
    .optional()
    .describe('Organizer of the event'),
  recurrence: z
    .object({
      pattern: z
        .object({
          type: z.enum(['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly']).describe('Recurrence pattern type'),
          interval: z.number().describe('Interval between occurrences'),
          month: z.number().optional().describe('Month for yearly recurrence'),
          dayOfMonth: z.number().optional().describe('Day of month for monthly recurrence'),
          daysOfWeek: z.array(z.enum(['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'])).optional(),
          firstDayOfWeek: z.enum(['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday']).optional(),
          index: z.enum(['first', 'second', 'third', 'fourth', 'last']).optional(),
        })
        .describe('Recurrence pattern'),
      range: z
        .object({
          type: z.enum(['endDate', 'noEnd', 'numbered']).describe('Type of recurrence range'),
          startDate: z.string().describe('Start date for recurrence'),
          endDate: z.string().optional().describe('End date for recurrence'),
          numberOfOccurrences: z.number().optional().describe('Number of occurrences'),
        })
        .describe('Recurrence range'),
    })
    .optional()
    .describe('Recurrence pattern for the event'),
  isAllDay: z.boolean().optional().describe('Whether the event is an all-day event'),
  showAs: z.enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown']).optional().describe('Show as status'),
  sensitivity: z.enum(['normal', 'personal', 'private', 'confidential']).optional().describe('Sensitivity level'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('Importance level'),
  reminderMinutesBeforeStart: z.number().optional().describe('Minutes before start to show reminder'),
  isReminderOn: z.boolean().optional().describe('Whether reminder is enabled'),
  hasAttachments: z.boolean().optional().describe('Whether the event has attachments'),
  webLink: z.string().optional().describe('URL to open the event in Outlook on the web'),
  onlineMeetingUrl: z.string().optional().describe('URL for online meeting'),
  isOnlineMeeting: z.boolean().optional().describe('Whether the event is an online meeting'),
  onlineMeetingProvider: z.enum(['unknown', 'skypeForBusiness', 'skypeForConsumer', 'teamsForBusiness']).optional(),
  allowNewTimeProposals: z.boolean().optional().describe('Whether attendees can propose new times'),
  hideAttendees: z.boolean().optional().describe('Whether to hide attendee list'),
  responseRequested: z.boolean().optional().describe('Whether response is requested'),
  seriesMasterId: z.string().optional().describe('ID of recurring series master'),
  isCancelled: z.boolean().optional().describe('Whether the event is cancelled'),
  isDraft: z.boolean().optional().describe('Whether the event is a draft'),
  isOrganizer: z.boolean().optional().describe('Whether current user is the organizer'),
  createdDateTime: z.string().optional().describe('ISO datetime when event was created'),
  lastModifiedDateTime: z.string().optional().describe('ISO datetime when event was last modified'),
});

// Outlook Category schema based on Microsoft Graph API
export const OutlookCategorySchema = z.object({
  id: z.string().describe('Unique identifier for the category'),
  displayName: z.string().describe('Display name of the category'),
  color: z
    .enum(['preset0', 'preset1', 'preset2', 'preset3', 'preset4', 'preset5', 'preset6', 'preset7', 'preset8', 'preset9', 'preset10', 'preset11', 'preset12', 'preset13', 'preset14', 'preset15', 'preset16', 'preset17', 'preset18', 'preset19', 'preset20', 'preset21', 'preset22', 'preset23', 'preset24'])
    .describe('Color preset for the category'),
});

// Microsoft Contact schema
export const OutlookContactSchema = z.object({
  id: z.string().optional().describe('Unique identifier for the contact'),
  displayName: z.string().optional().describe('Display name of the contact'),
  givenName: z.string().optional().describe('First name'),
  surname: z.string().optional().describe('Last name'),
  middleName: z.string().optional().describe('Middle name'),
  nickName: z.string().optional().describe('Nickname'),
  title: z.string().optional().describe('Job title'),
  companyName: z.string().optional().describe('Company name'),
  department: z.string().optional().describe('Department'),
  officeLocation: z.string().optional().describe('Office location'),
  profession: z.string().optional().describe('Profession'),
  businessPhones: z.array(z.string()).optional().describe('Business phone numbers'),
  homePhones: z.array(z.string()).optional().describe('Home phone numbers'),
  mobilePhone: z.string().optional().describe('Mobile phone number'),
  emailAddresses: z
    .array(
      z.object({
        name: z.string().optional().describe('Name for the email address'),
        address: z.string().email().describe('Email address'),
      })
    )
    .optional()
    .describe('Email addresses'),
  homeAddress: z
    .object({
      street: z.string().optional(),
      city: z.string().optional(),
      state: z.string().optional(),
      countryOrRegion: z.string().optional(),
      postalCode: z.string().optional(),
    })
    .optional()
    .describe('Home address'),
  businessAddress: z
    .object({
      street: z.string().optional(),
      city: z.string().optional(),
      state: z.string().optional(),
      countryOrRegion: z.string().optional(),
      postalCode: z.string().optional(),
    })
    .optional()
    .describe('Business address'),
  otherAddress: z
    .object({
      street: z.string().optional(),
      city: z.string().optional(),
      state: z.string().optional(),
      countryOrRegion: z.string().optional(),
      postalCode: z.string().optional(),
    })
    .optional()
    .describe('Other address'),
  birthday: z.string().optional().describe('Birthday in ISO date format'),
  personalNotes: z.string().optional().describe('Personal notes about the contact'),
  categories: z.array(z.string()).optional().describe('Categories assigned to the contact'),
  createdDateTime: z.string().optional().describe('ISO datetime when contact was created'),
  lastModifiedDateTime: z.string().optional().describe('ISO datetime when contact was last modified'),
  changeKey: z.string().optional().describe('Version of the contact'),
  parentFolderId: z.string().optional().describe('ID of the parent folder'),
});

// Export TypeScript types inferred from schemas
export type EmailAddress = z.infer<typeof EmailAddressSchema>;
export type Recipient = z.infer<typeof RecipientSchema>;
export type OutlookMessage = z.infer<typeof OutlookMessageSchema>;
export type OutlookFolder = z.infer<typeof OutlookFolderSchema>;
export type OutlookAttachment = z.infer<typeof OutlookAttachmentSchema>;
export type OneDriveFile = z.infer<typeof OneDriveFileSchema>;
export type OutlookCalendarEvent = z.infer<typeof OutlookCalendarEventSchema>;
export type OutlookContact = z.infer<typeof OutlookContactSchema>;
export type OutlookCategory = z.infer<typeof OutlookCategorySchema>;
