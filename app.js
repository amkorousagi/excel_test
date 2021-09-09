import xlsx from "xlsx";
import fs from "fs";

const getColName = (num) => {
  let res = "";
  while (num > 0) {
    const modulo = (num - 1) % 26;
    res = String.fromCharCode("A".charCodeAt(0) + modulo) + res;
    num = Math.floor((num - modulo) / 26);
  }
  return res;
};

const getColNumber = (name) => {
  let num = 0;
  let pow = 1;

  for (let i = name.length - 1; i >= 0; i--) {
    num += (name.charAt(i).charCodeAt(0) - "A".charCodeAt(0) + 1) * pow;
    pow *= 26;
  }
  return num;
};

const main = async () => {
  const workbook = xlsx.readFile("sample.xlsx");
  console.log(workbook.SheetNames);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  console.log(worksheet["OP4"]);
  console.log(getColNumber("PM"));
  const meta = {};
  let recent_category = "기본";
  for (let col = 3; col <= getColNumber("PM"); col++) {
    const category = worksheet[getColName(col) + 1]
      ? worksheet[getColName(col) + 1].v
      : recent_category;
    if (category != recent_category) recent_category = category;
    const name = worksheet[getColName(col) + 2].v;
    const unit = worksheet[getColName(col) + 3]
      ? worksheet[getColName(col) + 3].v
      : "none";
    const vari = [];
    const values = [];
    let empty = 0;
    for (let row = 4; row <= 379; row++) {
      const cell = worksheet[getColName(col) + row];
      if (cell) {
        if (vari.find((e) => e.value == cell.v) === undefined) {
          vari.push({ value: cell.v, count: 1 });
          values.push(cell.v);
        } else {
          vari[vari.findIndex((e) => e.value == cell.v)].count++;
        }
      } else {
        if (vari.find((e) => e.value == undefined) === undefined) {
          vari.push({ value: undefined, count: 1 });
          values.push(undefined);
        } else {
          vari[vari.findIndex((e) => e.value == undefined)].count++;
        }
        empty++;
      }
    }
    meta[getColName(col)] = {
      category,
      name,
      unit,
      vari: vari.length > 10 ? values : vari,
      empty,
      empty_ratio: (empty / 376) * 100 + "%",
    };
  }

  const resultbook = xlsx.utils.book_new();
  const resultsheet = JSON.parse(JSON.stringify(worksheet));

  // init
  for (let col = 1; col <= getColNumber("PM"); col++) {
    for (let row = 1; row <= 379; row++) {
      resultsheet[getColName(col) + row] = undefined;
    }
  }
  for (let col = 1; col <= 2; col++) {
    for (let row = 1; row <= 379; row++) {
      resultsheet[getColName(col) + row] = worksheet[getColName(col) + row];
    }
  }

  // fill
  let res_col = 3;
  for (let col = 3; col <= getColNumber("PM"); col++) {
    let proper = true;

    // if (
    //   Number(
    //     meta[getColName(col)].empty_ratio.substr(
    //       0,
    //       meta[getColName(col)].empty_ratio.length - 1
    //     )
    //   ) > 99
    // ) {
    //   proper = false;
    //   console.log(
    //     meta[getColName(col)].category,
    //     meta[getColName(col)].name,
    //     "는 빈값이 ",
    //     meta[getColName(col)].empty_ratio,
    //     "라서 제거"
    //   );
    // }
    if (meta[getColName(col)].vari.length < 2) {
      proper = false;
      console.log(
        meta[getColName(col)].category,
        meta[getColName(col)].name,
        "는 값의 다양성(빈값 포함)이 ",
        meta[getColName(col)].vari.length,
        "라서 제거"
      );
    }

    if (proper) {
      console.log(
        meta[getColName(col)].category,
        meta[getColName(col)].name,
        "는 적절함"
      );
      resultsheet[getColName(res_col) + 1] = worksheet[getColName(col) + 1]
        ? worksheet[getColName(col) + 1]
        : { v: meta[getColName(col)].category };
      for (let row = 2; row <= 379; row++) {
        resultsheet[getColName(res_col) + row] =
          worksheet[getColName(col) + row];
      }
      res_col++;
    }
  }
  console.log(res_col - 1);
  console.log();
  // append
  xlsx.utils.book_append_sheet(
    resultbook,
    resultsheet,
    "1차재염-데이터분석후처리본"
  );
//   xlsx.utils.book_append_sheet(
//     resultbook,
//     xlsx.utils.json_to_sheet(meta),
//     "메타데이터"
//   );

  xlsx.writeFileSync(resultbook, "res.xlsx");
  //   console.log(JSON.stringify(meta));
  fs.writeFileSync("meta.json", JSON.stringify(meta));
};

main();
