// SharePoint lists the app keeps its data in. The Provisioner creates and repairs them from these definitions;
// Power Automate and Power BI read the same lists (docs/INTEGRATIONS.md).

/** 2: Requests.PartialDays, AutoReply, GroupEventId; Employees.LeavingDate. */
export const SCHEMA_VERSION = 2;

export type FieldType = 'Text' | 'Note' | 'Number' | 'Boolean' | 'DateTime' | 'Choice';

export interface IFieldDef {
  name: string;
  type: FieldType;
  indexed?: boolean;
  choices?: string[];
}

export type ListKey = 'settings' | 'employees' | 'requests' | 'ledger' | 'delegations' | 'notifications' | 'emailQueue' | 'events' | 'audit';

export interface IListDef {
  key: ListKey;
  name: string;
  description: string;
  attachments?: boolean;
  fields: IFieldDef[];
}

export const REQUEST_STATUSES = ['Pending', 'Approved', 'Rejected', 'Cancelled', 'CancelRequested'];

export const LISTS: IListDef[] = [
  {
    key: 'settings',
    name: 'Settings',
    description: 'Time Off Manager settings (configuration, licence, branding).',
    fields: [{ name: 'SettingValue', type: 'Note' }]
  },
  {
    key: 'employees',
    name: 'Employees',
    description: 'People who take time off: manager, grade, start date and work calendar.',
    fields: [
      { name: 'Email', type: 'Text', indexed: true },
      { name: 'AadId', type: 'Text' },
      { name: 'JobTitle', type: 'Text' },
      { name: 'Department', type: 'Text' },
      { name: 'ManagerEmail', type: 'Text' },
      { name: 'Grade', type: 'Text' },
      { name: 'HireDate', type: 'Text' },
      { name: 'WorkCalendar', type: 'Text' },
      { name: 'Location', type: 'Text' },
      { name: 'Gender', type: 'Text' },
      { name: 'LeavingDate', type: 'Text' },
      { name: 'IsActive', type: 'Boolean' }
    ]
  },
  {
    key: 'requests',
    name: 'Requests',
    description: 'Leave and comp-off requests with their approval steps and history.',
    attachments: true,
    fields: [
      { name: 'RequestKind', type: 'Choice', choices: ['leave', 'compoff'] },
      { name: 'EmployeeEmail', type: 'Text', indexed: true },
      { name: 'EmployeeName', type: 'Text' },
      { name: 'Department', type: 'Text' },
      { name: 'LeaveType', type: 'Text' },
      { name: 'LeaveTypeName', type: 'Text' },
      { name: 'StartDate', type: 'Text' },
      { name: 'EndDate', type: 'Text' },
      { name: 'StartOn', type: 'DateTime' },
      { name: 'EndOn', type: 'DateTime' },
      { name: 'Unit', type: 'Choice', choices: ['day', 'half', 'hours'] },
      { name: 'HalfDayPart', type: 'Text' },
      { name: 'PartialDays', type: 'Text' },
      { name: 'AutoReply', type: 'Boolean' },
      { name: 'GroupEventId', type: 'Text' },
      { name: 'StartTime', type: 'Text' },
      { name: 'Hours', type: 'Number' },
      { name: 'Days', type: 'Number' },
      { name: 'LeaveYear', type: 'Number', indexed: true },
      { name: 'Status', type: 'Choice', choices: REQUEST_STATUSES, indexed: true },
      { name: 'CurrentStep', type: 'Number' },
      { name: 'CurrentApprovers', type: 'Note' },
      { name: 'Steps', type: 'Note' },
      { name: 'Reason', type: 'Note' },
      { name: 'SubmittedByEmail', type: 'Text' },
      { name: 'SubmittedByName', type: 'Text' },
      { name: 'SeriesId', type: 'Text' },
      { name: 'SeriesRule', type: 'Text' },
      { name: 'History', type: 'Note' },
      { name: 'CalendarEventId', type: 'Text' },
      { name: 'CancelReason', type: 'Note' }
    ]
  },
  {
    key: 'ledger',
    name: 'Ledger',
    description: 'Balance adjustments, carry-forward and comp-off credits.',
    fields: [
      { name: 'EmployeeEmail', type: 'Text', indexed: true },
      { name: 'LeaveType', type: 'Text' },
      { name: 'LeaveYear', type: 'Number', indexed: true },
      { name: 'Days', type: 'Number' },
      { name: 'EntryKind', type: 'Choice', choices: ['adjustment', 'carry', 'compoff', 'grant', 'expire'] },
      { name: 'Note', type: 'Note' },
      { name: 'Effective', type: 'Text' },
      { name: 'ExpiresOn', type: 'Text' },
      { name: 'UniqueKey', type: 'Text', indexed: true },
      { name: 'RequestId', type: 'Number' },
      { name: 'EnteredBy', type: 'Text' }
    ]
  },
  {
    key: 'delegations',
    name: 'Delegations',
    description: 'Approvers who hand their approvals to someone else for a period.',
    fields: [
      { name: 'DelegatorEmail', type: 'Text', indexed: true },
      { name: 'DelegatorName', type: 'Text' },
      { name: 'DelegateEmail', type: 'Text' },
      { name: 'DelegateName', type: 'Text' },
      { name: 'FromDate', type: 'Text' },
      { name: 'ToDate', type: 'Text' },
      { name: 'Note', type: 'Note' },
      { name: 'IsActive', type: 'Boolean' }
    ]
  },
  {
    key: 'notifications',
    name: 'Notifications',
    description: 'In-app notifications (the bell).',
    fields: [
      { name: 'Recipient', type: 'Text', indexed: true },
      { name: 'Body', type: 'Note' },
      { name: 'Link', type: 'Text' },
      { name: 'IsRead', type: 'Boolean' }
    ]
  },
  {
    key: 'emailQueue',
    name: 'EmailQueue',
    description: 'Emails waiting for the Power Automate flow that sends them.',
    fields: [
      { name: 'MailTo', type: 'Note' },
      { name: 'MailCc', type: 'Note' },
      { name: 'MailBody', type: 'Note' },
      { name: 'MailStatus', type: 'Choice', choices: ['Queued', 'Sent', 'Failed'], indexed: true },
      { name: 'MailError', type: 'Note' },
      { name: 'Template', type: 'Text' }
    ]
  },
  {
    key: 'events',
    name: 'Events',
    description: 'Outbox of leave events (submitted, approved, rejected, cancelled...) for Power Automate and other systems.',
    fields: [
      { name: 'EventType', type: 'Text', indexed: true },
      { name: 'RequestId', type: 'Number' },
      { name: 'EmployeeEmail', type: 'Text' },
      { name: 'Payload', type: 'Note' },
      { name: 'ProcessedStatus', type: 'Choice', choices: ['New', 'Processed', 'Failed'] }
    ]
  },
  {
    key: 'audit',
    name: 'Audit',
    description: 'Who changed what: settings, balances, decisions made on behalf of others.',
    fields: [
      { name: 'Action', type: 'Text' },
      { name: 'Actor', type: 'Text' },
      { name: 'Target', type: 'Text' },
      { name: 'Details', type: 'Note' }
    ]
  }
];

export function listTitle(prefix: string, key: ListKey): string {
  const def = LISTS.filter((l) => l.key === key)[0];
  return prefix + def.name;
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** CAML field definition for fields/createfieldasxml. */
export function fieldSchemaXml(f: IFieldDef): string {
  const common = `Name="${f.name}" StaticName="${f.name}" DisplayName="${f.name}"${f.indexed ? ' Indexed="TRUE"' : ''}`;
  switch (f.type) {
    case 'Note':
      return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />`;
    case 'Number':
      return `<Field Type="Number" ${common} Decimals="2" />`;
    case 'Boolean':
      return `<Field Type="Boolean" ${common}><Default>0</Default></Field>`;
    case 'DateTime':
      return `<Field Type="DateTime" ${common} Format="DateOnly" />`;
    case 'Choice':
      return `<Field Type="Choice" ${common} FillInChoice="TRUE"><CHOICES>${(f.choices || []).map((c) => `<CHOICE>${esc(c)}</CHOICE>`).join('')}</CHOICES></Field>`;
    default:
      return `<Field Type="Text" ${common} MaxLength="255" />`;
  }
}
