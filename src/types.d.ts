import { Buffer, Workbook } from "exceljs";
export interface Person {
  id: string;
  name: string;
  address?: string;
  city: string;
  zip: string;
  yearOfBirth: number;
  gender: Gender;
  emailAddress?: string;
  phoneNumber?: string;
}

export type PersonRow = [
  name: string,
  address: string,
  zip: string,
  city: string,
  emailAddress: string,
  phoneNumber: string,
  gender: Gender,
  yearOfBirth: number
];

export type TimeslotColumn = [
  date: string, // DD-MM-YYYY
  time: string, // HH:MM
  type: "Fysisk" | "Digital",
  hoursWithoutTeacher: number,
  hours: number
];
export type Gender = "K" | "M";

export interface Timeslot {
  date: string; // DD-MM-YYYY
  time: string; // HH:MM
  type: "Fysisk" | "Digital";
  hoursWithoutTeacher: number;
  hours: number;
  attendingUids: string[];
  name: string;
  description: string;
}

export interface ReportInput {
  persons: Person[];
  timeSlots: Timeslot[];
  clearNotes?: boolean;
  labels?: {
    name: string;
    address: string;
    zipCode: string;
    place: string;
    emailAddress: string;
    phone: string;
    gender: string;
    yearOfBirth: string;
    date: string;
    startTime: string;
    rehearsalFormat: string;
    hoursWithTeacher: string;
    hoursSansTeacher: string;
    reportName: string;
  };
}
export interface ReportOutput {
  result: Buffer;
  workbook: Workbook;
}
