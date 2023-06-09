import Excel from "exceljs";
import { TimeslotColumn, PersonRow, ReportInput, ReportOutput } from "./types";
import { join } from "path";

export async function getReport({
  persons,
  timeSlots,
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

  const buffer = await workbook.xlsx.writeBuffer();

  return {
    result: buffer,
    workbook: workbook,
  };
}
