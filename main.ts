import { getReport } from "./lib";
import { Person, ReportOutput, TimeSlot } from "./lib/types";

async function main() {
  const persons: Person[] = [demoUser1, demoUser2];
  const timeSlots: TimeSlot[] = [
    {
      attendingUids: ["1", "2"],
      date: "01-01-2021",
      hours: 2,
      hoursWithoutTeacher: 0,
      time: "12:00",
      type: "Fysisk",
    },
    {
      attendingUids: ["1"],
      date: "02-02-2022",
      hours: 1.4,
      hoursWithoutTeacher: 2,
      time: "14:00",
      type: "Fysisk",
    },
  ];

  const report = await getReport({ persons, timeSlots });

  await writeReportToFile(report);
}

async function writeReportToFile(report: ReportOutput) {
  // const base64String = (report.result as Buffer).toString("base64");
  await report.workbook.xlsx.writeFile("./output/result.xlsx");
}
const demoUser1: Person = {
  id: "1",
  name: "Ola Nordmann",
  city: "Oslo",
  gender: "M",
  yearOfBirth: 1990,
  zip: "1234",
};

const demoUser2: Person = {
  id: "2",
  name: "Kari Nordmann",
  city: "Oslo",
  gender: "F",
  yearOfBirth: 1991,
  zip: "1222",
};

main();
function writeFileSync(arg0: string, arg1: ArrayBuffer) {
  throw new Error("Function not implemented.");
}
