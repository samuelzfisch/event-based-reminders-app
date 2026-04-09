export type PlanType = "earnings" | "conference" | "press_release";

export type WeekendRule = "prior_business_day" | "none";

export type PlanItemStatus = "not_started" | "in_progress" | "complete";
export type PlanRowType = "reminder" | "email" | "calendar_event";
export type PlanDateBasis = "event" | "today";

export type ISODateString = string;

export interface PlanItem {
  id: string;
  title: string;
  customTitle?: string;
  body?: string;
  rowType?: PlanRowType;
  offsetDays: number;
  dateBasis?: PlanDateBasis;
  reminderTime?: string;
  rawDueDate: ISODateString;
  dueDate: ISODateString;
  customDueDate?: ISODateString;
  wasAdjusted: boolean;
  emailDraft?: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    body?: string;
  };
  durationDraft?: {
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
  };
  meetingDraft?: {
    attendees?: string[];
    location?: string;
    durationMinutes?: number;
    useCustomEnd?: boolean;
    endDate?: string;
    endTime?: string;
    isAllDay?: boolean;
    teamsMeeting?: boolean;
    addGoogleMeet?: boolean;
  };
  status: PlanItemStatus;
}

export interface Plan {
  id: string;
  name: string;
  type: PlanType;
  anchorDate: ISODateString;
  weekendRule: WeekendRule;
  version: number;
  items: PlanItem[];
  createdAt: string;
  updatedAt: string;
}
