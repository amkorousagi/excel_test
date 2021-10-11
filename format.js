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
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const meta = {};
  const dyes = [];
  const dyes_cnt = {};
  const recipe = {};
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
      if (cell == undefined) {
        continue;
      }
      if (name == "내광성등급") {
        // 표현 통일 (범위)->숫자
        if (cell.v == "1-2") {
          cell.v = 1.5;
        } else if (cell.v == "2-3") {
          worksheet[getColName(col) + row].v = 2.5;
        } else if (cell.v == "3-4") {
          cell.v = 3.5;
        } else if (cell.v == "4-5") {
          cell.v = 4.5;
        }
      }
      if (String(name).match(/염료.* 명/) != null) {
        // 염료 명 관련, 제조사명 은 안됨.
        if (!dyes.find((k) => k == cell.v)) {
          dyes.push(cell.v);
          dyes_cnt[cell.v] = 1;
        } else {
          dyes_cnt[cell.v]++;
        }
        const num = String(name).match(/[0-9]/)[0];
        //그냥 있는 것만 넣고 undefined면 나중에 0을 대신 채워넣는 식으로, sparse matrix?
        if (!recipe[row]) {
          recipe[row] = { name: [], amount: [] };
        }
        recipe[row].name.push(cell.v);
      }
      if (String(name).match(/염료.* 투입량/) != null) {
        // 염료 투입량 관련
        const num = String(name).match(/[0-9]/)[0];
        //그냥 있는 것만 넣고 undefined면 나중에 0을 대신 채워넣는 식으로, sparse matrix?
        recipe[row].amount.push(cell.v);
      }
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
  console.log("meta analy done");
  const list = [];

  const element = [];
  for (let col = 1; col <= getColNumber("AS") + dyes.length; col++) {
    element[col] = "";
  }
  for (let row = 1; row <= 379; row++) {
    list.push(element);
  }
  const resultbook = xlsx.utils.book_new();

  const resultsheet = xlsx.utils.aoa_to_sheet(list);
  // init
  for (let row = 1; row <= 379; row++) {
    resultsheet[getColName(1) + row] = { t: "n", v: row, w: String(row) };
  }

  // fill
  let res_col = 2; // 그냥 아예 1번 col에 못씀...
  for (let i = 0; i < dyes.length; i++) {
    resultsheet[getColName(res_col + i) + 1] = { v: "레시피:"+dyes[i] };
    resultsheet[getColName(res_col + i) + 2] = { v: dyes[i] };
    resultsheet[getColName(res_col + i) + 3] = { v: "g" };
    for (let row = 4; row <= 379; row++) {
      const index = recipe[row].name.findIndex((e) => e == dyes[i]);
      if (index != -1)
        resultsheet[getColName(res_col + i) + row] = {
          v: recipe[row].amount[index],
        };
      else resultsheet[getColName(res_col + i) + row] = { v: 0 };
    }
  }
  res_col += dyes.length;
  for (let col = 5; col <= getColNumber("PM"); col++) {
    let proper = true;

    if (meta[getColName(col)].vari.length < 2) {
      proper = false;
      /*  console.log(
        meta[getColName(col)].category,
        meta[getColName(col)].name,
        "는 값의 다양성(빈값 포함)이 ",
        meta[getColName(col)].vari.length,
        "라서 제거",
        meta[getColName(col)].empty_ratio
      );
      */
    }
    if (proper) {
      if (
        (meta[getColName(col)].name.match(/염료.* 투입량/) != null) |
        (meta[getColName(col)].name.match(/염료.* 명/) != null) |
        (meta[getColName(col)].name.match(/염료.* 제조사명/) != null) |
        (meta[getColName(col)].name.match(/조제.* 제조사명/) != null) 
      ) {
        console.log(meta[getColName(col)].name + " 삭제");
        continue;
      }
      console.log(
        meta[getColName(col)].category,
        meta[getColName(col)].name,
        "는 적절함"
      );
      resultsheet[getColName(res_col) + 1] = {
        v: meta[getColName(col)].category + ":" + meta[getColName(col)].name,
      };
      for (let row = 2; row <= 379; row++) {
        resultsheet[getColName(res_col) + row] =
          worksheet[getColName(col) + row];
      }
      res_col++;
    }
  }
  console.log(getColNumber("PM"), "->", res_col-1);

  // console.log(dyes.length);
  // console.log(dyes_cnt);
  // console.log(recipe);
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

  // xlsx.writeFileSync(resultbook, "format.xlsx");
  // fs.writeFileSync("meta.json", JSON.stringify(meta));
};

main();
