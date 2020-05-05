import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface TasksResponse {
  '@odata.context': string;
  '@odata.nextLink': string;
  value: Task[];
}

export interface Task {
  '@odata.etag': string;
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  changeKey: string;
  categories: any[];
  assignedTo: string;
  hasAttachments: boolean;
  importance: string;
  isReminderOn: boolean;
  owner: string;
  parentFolderId: string;
  sensitivity: string;
  status: string;
  subject: string;
  completedDateTime?: CompletedDateTime;
  dueDateTime?: CompletedDateTime;
  recurrence?: any;
  reminderDateTime?: any;
  startDateTime?: any;
  body: Body;
  'singleValueExtendedProperties@odata.context'?: string;
  singleValueExtendedProperties?: SingleValueExtendedProperty[];
}

export interface SingleValueExtendedProperty {
  id: string;
  value: string;
}

export interface Body {
  contentType: string;
  content: string;
}

export interface CompletedDateTime {
  dateTime: string;
  timeZone: string;
}

export interface IToDoProps {
  itemCount: number;
  context: WebPartContext;


}

export interface IToDoState {
  loading: boolean;
  tasks: Task[];
  errorMessage?: string;
  newTaskName?: string;
  teamId?: string;
  theme?: string;
}
