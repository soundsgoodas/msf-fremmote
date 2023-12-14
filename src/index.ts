import Excel from "exceljs";
import { TimeslotColumn, PersonRow, ReportInput, ReportOutput } from "./types";
import { join } from "path";

export async function getReport({
  persons,
  timeSlots,
  labels,
  clearNotes = false,
}: ReportInput): Promise<ReportOutput> {
  const file = join(
    __dirname,
    "..",
    "template",
    "Mal-for-import-av-fremmote.v21.04.2023.xlsx"
  );
  const workbook = new Excel.Workbook();

  await workbook.xlsx.readFile(file);

  const worksheet = workbook.worksheets[0];
  worksheet.insertRows(
    7,
    persons.map((it) => [
      "",
      ...([
        it.name,
        it.address || "",
        it.zip,
        it.city,
        it.emailAddress || "",
        it.phoneNumber || "",
        it.gender,
        it.yearOfBirth,
      ] as PersonRow),
      "",
      ...timeSlots.map(({ attendingUids }) =>
        attendingUids.includes(it.id) ? "x" : ""
      ),
    ])
  );

  let colIndex = 11;
  for (const timeSlot of timeSlots) {
    const column = worksheet.getColumn(colIndex);
    column.values = [
      "",
      ...([
        timeSlot.date,
        timeSlot.time,
        timeSlot.type,
        timeSlot.hoursWithoutTeacher,
        timeSlot.hours,
      ] as TimeslotColumn),
    ];
    worksheet.getCell(
      `${column.letter}2`
    ).note = `${timeSlot.name}\n${timeSlot.description}`;
    colIndex++;
  }

  if (typeof labels !== "undefined") {
    worksheet.name = labels.reportName;
    worksheet.getCell("B6").value = labels.name;
    worksheet.getCell("C6").value = labels.address;
    worksheet.getCell("D6").value = labels.zipCode;
    worksheet.getCell("E6").value = labels.place;
    worksheet.getCell("F6").value = labels.emailAddress;
    worksheet.getCell("G6").value = labels.phone;
    worksheet.getCell("H6").value = labels.gender;
    worksheet.getCell("I6").value = labels.yearOfBirth;
    worksheet.getCell("J2").value = labels.date;
    worksheet.getCell("J3").value = labels.startTime;
    worksheet.getCell("J4").value = labels.rehearsalFormat;
    worksheet.getCell("J5").value = labels.hoursSansTeacher;
    worksheet.getCell("J6").value = labels.hoursWithTeacher;
  }

  if (clearNotes) {
    worksheet.getCell("H6").note = "";
    worksheet.getCell("J2").note = "";
    worksheet.getCell("J3").note = "";
    worksheet.getCell("J4").note = "";
    worksheet.getCell("J5").note = "";
    worksheet.getCell("J6").note = "";
  }

  const buffer = await workbook.xlsx.writeBuffer();

  return {
    result: buffer,
    workbook: workbook,
  };
}
