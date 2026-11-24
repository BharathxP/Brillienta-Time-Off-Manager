import { IEntitlement, ILeaveType } from '../models/types';

// The 30 leave types a new installation starts with. Admins rename, recolour, change or switch off any of them
// (Admin › Leave types); inactive types are hidden from employees but keep their history.

function ent(partial: Partial<IEntitlement>): IEntitlement {
  return {
    mode: 'fixed',
    days: 0,
    byGrade: {},
    tenure: [],
    accrual: { frequency: 'monthly', timing: 'end' },
    prorate: true,
    carryForwardMax: 0,
    maxBalance: 0,
    expiryDays: 0,
    bookAhead: true,
    ...partial
  };
}

interface ISeed {
  key: string;
  name: string;
  code: string;
  color: string;
  icon: string;
  e: Partial<IEntitlement>;
  o?: Partial<ILeaveType>;
}

const SEEDS: ISeed[] = [
  {
    key: 'annual',
    name: 'Annual leave',
    code: 'AL',
    color: '#2e7d5b',
    icon: 'Airplane',
    e: { mode: 'accrual', days: 20, byGrade: { L1: 18, L2: 20, L3: 22, L4: 24, L5: 26 }, tenure: [{ years: 3, extraDays: 2 }, { years: 5, extraDays: 4 }, { years: 10, extraDays: 6 }], carryForwardMax: 5 },
    o: { units: { day: true, half: true, hours: true }, minNoticeDays: 3, allowRecurrence: true, autoApprove: { enabled: false, maxDays: 0 } }
  },
  { key: 'casual', name: 'Casual leave', code: 'CL', color: '#3d7fd9', icon: 'Sunny', e: { days: 8 }, o: { units: { day: true, half: true, hours: true }, autoApprove: { enabled: true, maxDays: 1 }, allowRecurrence: true } },
  { key: 'sick', name: 'Sick leave', code: 'SL', color: '#d9534f', icon: 'Health', e: { days: 10, carryForwardMax: 10, maxBalance: 30 }, o: { units: { day: true, half: true, hours: true }, attachmentOverDays: 2 } },
  { key: 'earned', name: 'Earned / privilege leave', code: 'PL', color: '#5b8c2a', icon: 'Trophy2', e: { mode: 'accrual', days: 15, carryForwardMax: 30, maxBalance: 45, tenure: [{ years: 5, extraDays: 3 }] }, o: { minNoticeDays: 7 } },
  { key: 'compoff', name: 'Comp off (in lieu)', code: 'CO', color: '#8a5cd1', icon: 'Switch', e: { mode: 'earned', expiryDays: 90, prorate: false }, o: { isCompOff: true, units: { day: true, half: true, hours: true }, autoApprove: { enabled: false, maxDays: 0 } } },
  { key: 'wfh', name: 'Work from home', code: 'WFH', color: '#17a2b8', icon: 'Home', e: { mode: 'unlimited' }, o: { paid: true, units: { day: true, half: true, hours: false }, autoApprove: { enabled: true, maxDays: 2 }, allowRecurrence: true } },
  { key: 'maternity', name: 'Maternity leave', code: 'ML', color: '#e06fa8', icon: 'Family', e: { days: 182, prorate: false }, o: { countCalendarDays: true, units: { day: true, half: false, hours: false }, minNoticeDays: 30, attachmentOverDays: 1, eligibility: { minServiceMonths: 3, grades: [], genders: ['Female'] } } },
  { key: 'paternity', name: 'Paternity leave', code: 'PTL', color: '#4a90c2', icon: 'Family', e: { days: 10, prorate: false }, o: { eligibility: { minServiceMonths: 3, grades: [], genders: ['Male'] } } },
  { key: 'parental', name: 'Parental leave', code: 'PRL', color: '#c77dba', icon: 'Childof', e: { days: 15, prorate: false } },
  { key: 'adoption', name: 'Adoption leave', code: 'ADL', color: '#b07cc6', icon: 'Heart', e: { days: 84, prorate: false }, o: { countCalendarDays: true, units: { day: true, half: false, hours: false } } },
  { key: 'bereavement', name: 'Bereavement leave', code: 'BL', color: '#5f6b73', icon: 'Flower', e: { days: 5, prorate: false }, o: { autoApprove: { enabled: true, maxDays: 5 } } },
  { key: 'marriage', name: 'Marriage leave', code: 'MAR', color: '#e6a23c', icon: 'Ringer', e: { days: 5, prorate: false }, o: { minNoticeDays: 14 } },
  { key: 'birthday', name: 'Birthday leave', code: 'BDL', color: '#f39c12', icon: 'BirthdayCake', e: { days: 1, prorate: false }, o: { units: { day: true, half: true, hours: false }, autoApprove: { enabled: true, maxDays: 1 } } },
  { key: 'optional', name: 'Optional / floating holiday', code: 'OH', color: '#16a085', icon: 'Calendar', e: { days: 2, prorate: false }, o: { autoApprove: { enabled: true, maxDays: 1 } } },
  { key: 'volunteer', name: 'Volunteer leave', code: 'VOL', color: '#27ae60', icon: 'Group', e: { days: 2 } },
  { key: 'study', name: 'Study / exam leave', code: 'STL', color: '#2c6fbb', icon: 'Education', e: { days: 5 }, o: { attachmentOverDays: 1 } },
  { key: 'training', name: 'Training leave', code: 'TRL', color: '#386fa4', icon: 'D365TalentLearn', e: { mode: 'unlimited' } },
  { key: 'jury', name: 'Jury duty', code: 'JD', color: '#7f8c8d', icon: 'Bank', e: { mode: 'unlimited' }, o: { attachmentOverDays: 1 } },
  { key: 'military', name: 'Military leave', code: 'MIL', color: '#556b2f', icon: 'Shield', e: { days: 15, prorate: false }, o: { attachmentOverDays: 1 } },
  { key: 'religious', name: 'Religious observance', code: 'REL', color: '#a0522d', icon: 'Lightbulb', e: { days: 2, prorate: false } },
  { key: 'medical', name: 'Medical appointment', code: 'MED', color: '#e74c3c', icon: 'Hospital', e: { days: 3 }, o: { units: { day: false, half: true, hours: true }, autoApprove: { enabled: true, maxDays: 0.5 }, allowRecurrence: true } },
  { key: 'care', name: 'Family care leave', code: 'FCL', color: '#d35400', icon: 'HeartFill', e: { days: 5 } },
  { key: 'relocation', name: 'Relocation leave', code: 'RLC', color: '#8e6e53', icon: 'MapPin', e: { days: 3, prorate: false } },
  { key: 'sabbatical', name: 'Sabbatical', code: 'SAB', color: '#34495e', icon: 'Vacation', e: { days: 30, prorate: false }, o: { paid: false, minNoticeDays: 60, eligibility: { minServiceMonths: 60, grades: ['L4', 'L5'], genders: [] } } },
  { key: 'unpaid', name: 'Unpaid leave', code: 'LWP', color: '#95a5a6', icon: 'Money', e: { mode: 'unlimited' }, o: { paid: false } },
  { key: 'wellness', name: 'Wellness day', code: 'WD', color: '#1abc9c', icon: 'Emoji2', e: { days: 4 }, o: { units: { day: true, half: true, hours: false } } },
  { key: 'voting', name: 'Voting leave', code: 'VL', color: '#2980b9', icon: 'CheckboxComposite', e: { days: 1, prorate: false }, o: { units: { day: true, half: true, hours: true }, autoApprove: { enabled: true, maxDays: 1 } } },
  { key: 'disaster', name: 'Emergency / disaster leave', code: 'EML', color: '#c0392b', icon: 'Warning', e: { days: 3, prorate: false }, o: { autoApprove: { enabled: true, maxDays: 3 } } },
  { key: 'menstrual', name: 'Menstrual wellness leave', code: 'MWL', color: '#cf6a87', icon: 'Precipitation', e: { days: 12 }, o: { units: { day: true, half: true, hours: false }, eligibility: { minServiceMonths: 0, grades: [], genders: ['Female'] } } },
  { key: 'onduty', name: 'On duty / business travel', code: 'OD', color: '#607d8b', icon: 'Car', e: { mode: 'unlimited' }, o: { units: { day: true, half: true, hours: true } } }
];

export function defaultLeaveTypes(): ILeaveType[] {
  return SEEDS.map((s) => ({
    key: s.key,
    name: s.name,
    code: s.code,
    color: s.color,
    icon: s.icon,
    active: true,
    paid: true,
    description: '',
    units: { day: true, half: true, hours: false },
    countCalendarDays: false,
    entitlement: ent(s.e),
    allowNegative: 0,
    minNoticeDays: 0,
    maxConsecutiveDays: 0,
    attachmentOverDays: 0,
    eligibility: { minServiceMonths: 0, grades: [], genders: [] },
    workflowKey: s.key === 'sabbatical' || s.key === 'maternity' ? 'extended' : 'standard',
    autoApprove: { enabled: false, maxDays: 0 },
    allowRecurrence: false,
    ...(s.o || {})
  }));
}

/** A blank type for "New leave type". */
export function blankLeaveType(key: string): ILeaveType {
  return {
    ...defaultLeaveTypes()[1],
    key,
    name: 'New leave type',
    code: key.substring(0, 4).toUpperCase(),
    color: '#3d7fd9',
    icon: 'Calendar',
    entitlement: ent({ days: 5 }),
    autoApprove: { enabled: false, maxDays: 0 },
    allowRecurrence: false,
    description: ''
  };
}
