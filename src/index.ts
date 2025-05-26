import xlsx from "node-xlsx";
import * as fs from "fs";
import * as xl from "excel4node";

const calculateHours = (data: any) => {
  if (isNaN(+data)) return 0;

  return +data;
};

const calculateSemHours = (
  firstSem: number,
  secondSem: number,
  groupsHours: number
): number[] | string[] => {
  if (firstSem === 0 && secondSem === 0) return ["", ""];

  if (groupsHours === 0) return ["", ""];

  if (firstSem === 0) return [0, groupsHours];

  const difference = +(firstSem / secondSem).toFixed(1) + 1;

  const secondSemResult = Math.floor(groupsHours / difference);
  const firstSemResult = groupsHours - secondSemResult;

  return [firstSemResult, secondSemResult];
};

const app = async () => {
  const workSheetsFromFile = xlsx.parse(`./input.xlsx`);

  const data = [];

  for await (const sheet of workSheetsFromFile) {
    const groupsOnTheSheet = [];

    let key = 0;

    for await (const el of sheet.data) {
      if (el[0]?.toLowerCase().includes("группа:")) {
        groupsOnTheSheet.push({
          name: el[1].trim(),
          startPosition: key,
        });
      }

      key++;
    }

    for await (const group of groupsOnTheSheet) {
      const headerData = sheet.data.slice(group.startPosition + 2);

      const props = [
        ...headerData[0].slice(3, 7),
        ...headerData[0].slice(8, 9),
        ...headerData[0].slice(12, 14),
      ];
      // console.log(props);

      for await (const el of sheet.data.slice(group.startPosition + 3)) {
        if (!el[0]) continue;

        if (el[0].toLowerCase().includes("итого")) {
          break;
        }

        const teacherName = el
          .slice(1)
          .filter(
            (el) => typeof el === "string" && el !== "" && el.trim().length > 1
          );

        if (teacherName[0] === "Елисеева") console.log(2);

        if (teacherName?.[0]?.includes("/")) continue;

        const findKey = data.findIndex(
          (el) => el.teacherName === teacherName[0]
        );

        const details = [];

        const ids = [3, 4, 5, 6, 8, 12, 13];

        ids.forEach((id, key) => {
          details.push({
            title: props?.[key]?.toLowerCase(),
            hours: calculateHours(el?.[id]),
          });
        });

        if (findKey === -1) {
          data.push({
            lessons: [
              {
                name: el[0],
                totalHours: el[1],
                group: group.name,
                details,
              },
            ],
            hours: el[1],
            teacherName: teacherName[0],
          });
        } else {
          data[findKey].lessons.push({
            name: el[0],
            totalHours: el[1],
            group: group.name,
            details,
          });

          data[findKey].hours += el[1];
        }
      }
    }
  }

  fs.writeFileSync("./all.json", JSON.stringify(data));

  buildExcel();
};

const writeHeader = (
  ws,
  offset: number,
  teacherName: string,
  style,
  textStyle
) => {
  ws.cell(1 + offset, 1)
    .string(teacherName)
    .style(style);
  ws.cell(1 + offset, 2, 1 + offset, 14, true)
    .string("")
    .style(style);
  ws.cell(2 + offset, 1, 4 + offset, 1, true)
    .string("Дисциплины")
    .style(textStyle);
  ws.cell(2 + offset, 2, 4 + offset, 2, true)
    .string("Группы")
    .style(textStyle);
  ws.cell(2 + offset, 3, 2 + offset, 10, true)
    .string("Часы педнагрузки")
    .style(textStyle);
  ws.cell(3 + offset, 3, 3 + offset, 6, true)
    .string("1 семестр")
    .style(textStyle);
  ws.cell(4 + offset, 3)
    .string("часы")
    .style(textStyle);
  ws.cell(4 + offset, 4)
    .string("подгр.")
    .style(textStyle);
  ws.cell(4 + offset, 5)
    .string("экз.")
    .style(textStyle);
  ws.cell(4 + offset, 6)
    .string("конс.")
    .style(textStyle);
  ws.cell(3 + offset, 7, 3 + offset, 10, true)
    .string("2 семестр")
    .style(textStyle);
  ws.cell(4 + offset, 7)
    .string("часы")
    .style(textStyle);
  ws.cell(4 + offset, 8)
    .string("подгр.")
    .style(textStyle);
  ws.cell(4 + offset, 9)
    .string("экз.")
    .style(textStyle);
  ws.cell(4 + offset, 10)
    .string("конс.")
    .style(textStyle);
  ws.cell(2 + offset, 11, 4 + offset, 11, true)
    .string("Проверка к/р")
    .style(textStyle);
  ws.cell(2 + offset, 12, 2 + offset, 13, true)
    .string("ВКР")
    .style(textStyle);
  ws.cell(3 + offset, 12, 4 + offset, 12, true)
    .string("рук.")
    .style(textStyle);
  ws.cell(3 + offset, 13, 4 + offset, 13, true)
    .string("защита")
    .style(textStyle);
  ws.cell(2 + offset, 14, 4 + offset, 14, true)
    .string("Итого")
    .style(textStyle);
};

const buildExcel = async () => {
  const wb = new xl.Workbook();

  const ws = wb.addWorksheet("Педнагрузка");

  const style = wb.createStyle({
    font: {
      color: "#000000",
      size: 10,
      bold: true,
    },
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
    border: {
      left: {
        style: "thin",
        color: "#000000",
      },
      right: {
        style: "thin",
        color: "#000000",
      },
      top: {
        style: "thin",
        color: "#000000",
      },
      bottom: {
        style: "thin",
        color: "#000000",
      },
    },
  });

  const footerStyle = wb.createStyle({
    font: {
      color: "#000000",
      size: 10,
      bold: true,
    },
    alignment: {
      horizontal: "right",
      vertical: "center",
    },
    border: {
      left: {
        style: "thin",
        color: "#000000",
      },
      right: {
        style: "thin",
        color: "#000000",
      },
      top: {
        style: "thin",
        color: "#000000",
      },
      bottom: {
        style: "thin",
        color: "#000000",
      },
    },
  });

  const textStyle = wb.createStyle({
    font: {
      color: "#000000",
      size: 10,
    },
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
    border: {
      left: {
        style: "thin",
        color: "#000000",
      },
      right: {
        style: "thin",
        color: "#000000",
      },
      top: {
        style: "thin",
        color: "#000000",
      },
      bottom: {
        style: "thin",
        color: "#000000",
      },
    },
  });

  const allStyle = wb.createStyle({
    font: {
      color: "#000000",
      size: 10,
    },
    alignment: {
      vertical: "top",
    },
    border: {
      left: {
        style: "thin",
        color: "#000000",
      },
      right: {
        style: "thin",
        color: "#000000",
      },
      top: {
        style: "thin",
        color: "#000000",
      },
      bottom: {
        style: "thin",
        color: "#000000",
      },
    },
  });

  const data = JSON.parse(fs.readFileSync("./all.json", "utf8"));

  ws.column(1).setWidth(45);

  const writeCell = (row: number, col: number, data: string | number) => {
    if (typeof data === "number" && !isNaN(+data) && +data > 0) {
      ws.cell(row, col).number(+data).style(allStyle);
    } else {
      ws.cell(row, col)
        .string(String(data ? data : ""))
        .style(allStyle);
    }
  };

  let offset: number = 0;

  for await (const inputData of data) {
    writeHeader(ws, offset, inputData.teacherName, style, textStyle);

    let startPosition = 5;

    let lessonsWithUniqueName = [];

    inputData.lessons.forEach((lesson: { name: string }) => {
      if (
        lessonsWithUniqueName.findIndex((el) => el.name === lesson.name) === -1
      ) {
        lessonsWithUniqueName.push(lesson);
      }
    });

    const lessons = lessonsWithUniqueName.map((lesson) => lesson.name);

    const uniqueLessons = [...new Set(lessons)];

    for await (const el of uniqueLessons) {
      const groupsWithThatLesson = inputData.lessons.filter(
        (lesson) => lesson.name === el
      );

      const groups = groupsWithThatLesson.map((lesson) => lesson.group);

      ws.cell(
        startPosition + offset,
        1,
        startPosition + groups.length - 1 + offset,
        1,
        true
      )
        .string(el)
        .style(allStyle);

      let key = 0;

      for await (const group of groups) {
        const groupWithThatLesson = groupsWithThatLesson.find(
          (lesson) => lesson.group === group
        );

        writeCell(startPosition + key + offset, 2, group);

        const firstSemHours = groupWithThatLesson.details.filter(
          (el) => el.title === "1 сем"
        )[0].hours;

        let total = 0;

        total += +firstSemHours || 0;

        writeCell(startPosition + key + offset, 3, firstSemHours);

        const hoursBySem = calculateSemHours(
          groupWithThatLesson.details.filter((el) => el.title === "1 сем")[0]
            .hours || 0,
          groupWithThatLesson.details.filter((el) => el.title === "2 сем")[0]
            .hours || 0,
          groupWithThatLesson.details.filter((el) =>
            el.title.includes("деление")
          )[0].hours
        );

        if (inputData.teacherName === "Александрия") {
          console.log(
            hoursBySem,
            groupWithThatLesson.details.filter((el) => el.title === "1 сем")[0]
              .hours || 0,
            groupWithThatLesson.details.filter((el) => el.title === "2 сем")[0]
              .hours || 0,
            groupWithThatLesson.details.filter((el) =>
              el.title.includes("деление")
            )[0].hours
          );
        }

        total += +hoursBySem[0] || 0;
        total += +hoursBySem[1] || 0;

        writeCell(startPosition + key + offset, 4, +hoursBySem[0]);

        writeCell(startPosition + key + offset, 5, "");

        writeCell(startPosition + key + offset, 6, "");

        const hoursBySem2 = groupWithThatLesson.details.filter(
          (el) => el.title === "2 сем"
        )[0].hours;

        total += +hoursBySem2 || 0;

        writeCell(startPosition + key + offset, 7, hoursBySem2);

        writeCell(startPosition + key + offset, 8, hoursBySem[1]);

        const hoursByFinal = groupWithThatLesson.details.filter((el) =>
          el.title.includes("экзамен")
        )?.[0]?.hours;

        total += +hoursByFinal || 0;

        writeCell(startPosition + key + offset, 9, hoursByFinal);

        writeCell(startPosition + key + offset, 10, "");

        const checkHours = groupWithThatLesson.details.filter((el) =>
          el.title.includes("проверка к")
        )?.[0]?.hours;

        total += +checkHours || 0;

        writeCell(startPosition + key + offset, 11, checkHours);

        const vkrHours = groupWithThatLesson.details.filter((el) =>
          el.title.includes("вкр")
        )?.[0]?.hours;

        total += +vkrHours || 0;

        writeCell(startPosition + key + offset, 12, +vkrHours || "");

        const giaHours = groupWithThatLesson.details.filter((el) =>
          el.title.includes("гиа")
        )?.[0]?.hours;

        total += +giaHours || 0;

        writeCell(startPosition + key + offset, 13, +giaHours || "");

        writeCell(startPosition + key + offset, 14, total);

        key++;
      }

      startPosition += groups.length;
    }

    ws.cell(offset + 5 + inputData.lessons.length, 1)
      .string("Итого")
      .style(footerStyle);

    ws.cell(offset + 5 + inputData.lessons.length, 2)
      .string("")
      .style(footerStyle);

    for (let i = 0; i < 12; i++) {
      const letterInPosition = xl.getExcelAlpha(i + 3);

      ws.cell(offset + 5 + inputData.lessons.length, i + 3)
        .formula(
          `IF(SUM(${letterInPosition}${offset + 5}:${letterInPosition}${
            offset + 5 + inputData.lessons.length - 1
          }) = 0, "", SUM(${letterInPosition}${offset + 5}:${letterInPosition}${
            offset + 5 + inputData.lessons.length - 1
          }))`
        )
        .style(footerStyle);
    }

    offset += startPosition + 1;
  }

  wb.write("Excel.xlsx");
};

app();
