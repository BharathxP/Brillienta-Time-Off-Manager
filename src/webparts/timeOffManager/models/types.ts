// Domain model of Brillienta Time Off Manager. Dates are ISO calendar days ("yyyy-mm-dd") unless named ...At (instants).

export type Role = 'admin' | 'hr' | 'manager' | 'employee';

export interface IPerson {
  email: string;
  name: string;
}

export interface ICurrentUser {
  id: number;
  title: string;
  email: string;
  loginName: string;
  isSiteAdmin: boolean;
  canManageWeb: boolean;
}

/** A person who takes time off (TO_Employees). Imported from Microsoft Entra ID, a CSV file or entered by HR. */
export interface IEmployee {
  id?: number;
  email: string;
  name: string;
  aadId?: string;
  jobTitle?: string;
  department?: string;
  managerEmail?: string;
  /** Level / grade, used by grade-based entitlements and eligibility. */
  grade?: string;
  /** Start date, used by tenure-based entitlements, accrual and proration. */
  hireDate?: string;
  /** Work calendar key (working week + holidays). Empty = the default calendar. */
  calendar?: string;
  location?: string;
  gender?: string;
  /** Last working day: entitlements are prorated to it and no leave can be requested after it. */
  leavingDate?: string;
  active: boolean;
}

// ------------------------------------------------------------------ leave types & policies

export type LeaveUnit = 'day' | 'half' | 'hours';
export type EntitlementMode = 'fixed' | 'accrual' | 'earned' | 'unlimited';
export type AccrualFrequency = 'monthly' | 'quarterly' | 'biweekly' | 'yearly';

export interface ITenureTier {
  /** Completed years of service at the start of the leave year (or today for accrual). */
  years: number;
  /** Extra days per year from this tier on. */
  extraDays: number;
}

export interface IEntitlement {
  mode: EntitlementMode;
  /** Days per leave year (fixed) or per year spread over the accrual periods (accrual). */
  days: number;
  /** Per grade override of `days` (grade → days per year). */
  byGrade: Record<string, number>;
  /** Tenure tiers: the highest tier reached adds its extra days. */
  tenure: ITenureTier[];
  accrual: { frequency: AccrualFrequency; timing: 'start' | 'end' };
  /** Prorate the first year by the start date. */
  prorate: boolean;
  /** Days carried into the next leave year at most (0 = none, -1 = all). */
  carryForwardMax: number;
  /** Balance never grows above this (0 = no cap). */
  maxBalance: number;
  /** Earned credits (comp off) lapse this many days after they are earned (0 = never). */
  expiryDays: number;
  /** Accrual: requests may use what will have accrued by the end of the leave year (not only what has accrued so far). */
  bookAhead?: boolean;
}

export interface ILeaveType {
  key: string;
  name: string;
  code: string;
  color: string;
  icon: string;
  active: boolean;
  paid: boolean;
  description: string;
  units: { day: boolean; half: boolean; hours: boolean };
  /** Count every calendar day (weekends and holidays included), e.g. maternity leave. */
  countCalendarDays: boolean;
  entitlement: IEntitlement;
  /** Days the balance may go below zero. */
  allowNegative: number;
  minNoticeDays: number;
  /** Longest single request in days (0 = no limit). */
  maxConsecutiveDays: number;
  /** Ask for a supporting document above this many days (0 = never). */
  attachmentOverDays: number;
  eligibility: { minServiceMonths: number; grades: string[]; genders: string[] };
  workflowKey: string;
  autoApprove: { enabled: boolean; maxDays: number };
  allowRecurrence: boolean;
  /** Comp off: what the credit is earned for. */
  isCompOff?: boolean;
  /**
   * Sandwich rule: weekends and holidays that fall between two leave days of the same request count as leave
   * (e.g. Friday + Monday = 4 days).
   */
  sandwich?: boolean;
}

// ------------------------------------------------------------------ workflows

export type ApproverKind = 'manager' | 'skipManager' | 'departmentHead' | 'hr' | 'person';

export interface IApprovalLevel {
  kind: ApproverKind;
  label: string;
  person?: IPerson;
  /** Only when the request is at least this many days (0 = always). */
  minDays: number;
}

export interface IWorkflow {
  key: string;
  name: string;
  levels: IApprovalLevel[];
}

export type StepStatus = 'waiting' | 'pending' | 'approved' | 'rejected' | 'skipped';

export interface IApprovalStep {
  label: string;
  kind: ApproverKind;
  /** Who may decide this step (delegates included). */
  approvers: Array<IPerson & { delegatedFrom?: string }>;
  status: StepStatus;
  actedBy?: IPerson;
  /** Decided by HR / an admin for the approver. */
  onBehalf?: boolean;
  comment?: string;
  actedAt?: string;
  startedAt?: string;
  remindedAt?: string;
  reminders?: number;
}

// ------------------------------------------------------------------ requests

export type RequestKind = 'leave' | 'compoff';
export type PartialDays = '' | 'start' | 'end' | 'both';
export type RequestStatus = 'Pending' | 'Approved' | 'Rejected' | 'Cancelled' | 'CancelRequested';

export interface IHistoryEntry {
  at: string;
  by: string;
  action: string;
  note?: string;
}

export interface ILeaveRequest {
  id: number;
  kind: RequestKind;
  employeeEmail: string;
  employeeName: string;
  leaveType: string;
  start: string;
  end: string;
  unit: LeaveUnit;
  halfDayPart?: 'am' | 'pm';
  /** Multi-day leave that starts in the afternoon ('start'), ends at midday ('end') or both. */
  partial?: PartialDays;
  startTime?: string;
  hours: number;
  /** Days consumed (or credited, comp off). */
  days: number;
  leaveYear: number;
  status: RequestStatus;
  currentStep: number;
  steps: IApprovalStep[];
  reason: string;
  submittedBy?: IPerson;
  seriesId?: string;
  seriesRule?: string;
  history: IHistoryEntry[];
  calendarEventId?: string;
  /** Event on the Microsoft 365 group (Teams channel) calendar. */
  groupEventId?: string;
  /** Set the employee's Outlook automatic reply for the leave. */
  autoReply?: boolean;
  cancelReason?: string;
  hasAttachment?: boolean;
  created: string;
  modified?: string;
}

export interface ILedgerEntry {
  id?: number;
  employeeEmail: string;
  leaveType: string;
  leaveYear: number;
  days: number;
  kind: 'adjustment' | 'carry' | 'compoff' | 'grant' | 'expire';
  note: string;
  effective: string;
  expiresOn?: string;
  uniqueKey?: string;
  requestId?: number;
  by?: string;
}

export interface IDelegation {
  id?: number;
  delegator: IPerson;
  delegate: IPerson;
  from: string;
  to: string;
  note?: string;
  active: boolean;
}

export interface INotification {
  id: number;
  recipient: string;
  title: string;
  body: string;
  link: string;
  read: boolean;
  created: string;
}

// ------------------------------------------------------------------ calendars & templates

export interface IHoliday {
  date: string;
  name: string;
  optional?: boolean;
}

export interface IWorkCalendar {
  key: string;
  name: string;
  /** Working weekdays, 0 = Sunday ... 6 = Saturday. */
  workDays: number[];
  hoursPerDay: number;
  dayStart: string;
  holidays: IHoliday[];
}

export interface ILeaveTemplate {
  key: string;
  name: string;
  leaveType: string;
  unit: LeaveUnit;
  halfDayPart?: 'am' | 'pm';
  startTime?: string;
  hours?: number;
  days?: number;
  reason: string;
  recurrence?: IRecurrence;
  owner?: string;
}

export type RecurrenceFreq = 'none' | 'weekly' | 'biweekly' | 'monthly';

export interface IRecurrence {
  freq: RecurrenceFreq;
  /** Weekdays for weekly / bi-weekly (0 = Sunday). Empty = the start date's weekday. */
  weekdays: number[];
  until?: string;
  count?: number;
}

// ------------------------------------------------------------------ configuration

/** Dates when leave is not allowed (or only with a warning), e.g. year-end close or a product launch. */
export interface IBlackout {
  key: string;
  name: string;
  from: string;
  to: string;
  /** Empty = every department. */
  departments: string[];
  /** Empty = every leave type. */
  leaveTypes: string[];
  mode: 'block' | 'warn';
}

export interface IConfig {
  appName: string;
  orgName: string;
  brandColor: string;
  accentColor: string;
  headline: string;
  /** Month the leave year starts (1 = January). */
  yearStartMonth: number;
  weekStartsOn: 0 | 1;
  leaveTypes: ILeaveType[];
  workflows: IWorkflow[];
  calendars: IWorkCalendar[];
  defaultCalendar: string;
  templates: ILeaveTemplate[];
  grades: string[];
  departmentHeads: Record<string, IPerson>;
  admins: IPerson[];
  hr: IPerson[];
  /** Hour-based leave: the selectable lengths. */
  hourOptions: number[];
  onBehalf: { managers: boolean; hr: boolean };
  cancellation: { pendingDirect: boolean; approvedNeedsApproval: boolean; allowPast: boolean };
  reminders: { enabled: boolean; afterHours: number; repeatHours: number; maxReminders: number; escalateAfterHours: number };
  autoApproveAfterHours: number;
  notifications: { email: boolean; inApp: boolean; notifyHr: boolean; ccEmails: string };
  integrations: { webhookUrl: string; outbox: boolean };
  calendarSync: { outlook: boolean; showAs: 'oof' | 'free' | 'busy' };
  teamCalendar: { scope: 'team' | 'department' | 'everyone'; showLeaveType: boolean };
  compOff: { enabled: boolean; leaveType: string; requireNonWorkingDay: boolean; maxDaysPerClaim: number };
  blackouts: IBlackout[];
  /** Minimum staffing: how many of a department may be away on the same day. */
  staffing: { enabled: boolean; maxAwayPercent: number; maxAwayCount: number; mode: 'warn' | 'block'; perDepartment: Record<string, number> };
  /** Publish approved leave to a Microsoft 365 group calendar (the Teams channel calendar). */
  groupCalendar: { enabled: boolean; groupId: string; groupName: string };
  /** Offer employees an Outlook automatic reply during their leave. */
  autoReply: { enabled: boolean; message: string };
}

export interface ISetupState {
  installed: boolean;
  schemaVersion: number;
}
