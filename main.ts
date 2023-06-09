import { getReport } from "./src";
import { Person, ReportOutput, Timeslot } from "./src/types";

async function main() {
  const persons: Person[] = [demoUser1, demoUser2];
  const timeSlots: Timeslot[] = [
    {
      attendingUids: ["1", "2"],
      date: "01-01-2021",
      hours: 2,
      hoursWithoutTeacher: 0,
      time: "12:00",
      type: "Fysisk",
      description: "Rehearsal",
      name: "Ukentlig korøvelse",
    },
    {
      attendingUids: ["1"],
      date: "02-02-2022",
      hours: 1.4,
      hoursWithoutTeacher: 2,
      time: "14:00",
      type: "Fysisk",
      description: "Rehearsal",
      name: "Concert",
    },
  ];

  const report = await getReport({ persons, timeSlots });

  await writeReportToFile(report);
}

async function writeReportToFile(report: ReportOutput) {
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
  gender: "K",
  yearOfBirth: 1991,
  zip: "1222",
};

main();
