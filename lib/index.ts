import Excel from "exceljs";
import { PersonRow, ReportInput, ReportOutput } from "./types";

export async function getReport({
  persons,
  timeSlots,
}: ReportInput): Promise<ReportOutput> {
  const file = "./template/Mal-for-import-av-fremmote.v21.04.2023.xlsx";
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
    ])
  );

  const buffer = await workbook.xlsx.writeBuffer();

  return {
    result: buffer,
    workbook: workbook,
  };
}
