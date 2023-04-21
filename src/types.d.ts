import { Buffer, Workbook } from "exceljs";
interface Person {
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
  hours: number,
  hoursWithoutTeacher: number
];
type Gender = "F" | "M";

interface Timeslot {
  date: string; // DD-MM-YYYY
  time: string; // HH:MM
  type: "Fysisk" | "Digital";
  hoursWithoutTeacher: number;
  hours: number;
  attendingUids: string[];
}

interface ReportInput {
  persons: Person[];
  timeSlots: Timeslot[];
}
interface ReportOutput {
  result: Buffer;
  workbook: Workbook;
}
