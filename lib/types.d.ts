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

type Gender = "F" | "M";

interface TimeSlot {
  date: string; // DD-MM-YYYY
  time: string; // HH:MM
  type: "Fysisk" | "Digital";
  hoursWithoutTeacher: number;
  hours: number;
  attendingUids: string[];
}

interface ReportInput {
  persons: Person[];
  timeSlots: TimeSlot[];
}
interface ReportOutput {
  result: Buffer;
  workbook: Workbook;
}
