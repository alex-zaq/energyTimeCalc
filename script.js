'use strict'

let xlsx = require("xlsx");
let wb = xlsx.readFile('Input.xlsx')
let sh1 = wb.Sheets['data'];
let loadData = xlsx.utils.sheet_to_json(sh1);
let startDate = new Date(2021, 0, 1, 0, 0, 0);
let year = 2021;

loadData.forEach((value, i) => {
  value['date'] = new Date(new Date(2021, 0, 1, 0, 0, 0).setHours(i));
});


function GetIndexByDate(loadData, date) {

  let res = -1;
  for (let i = 0; i < loadData.length; i++) {
    if (date.getTime() == loadData[i]['date'].getTime()) {
      return i;
    }
  }

}


class Day {
  constructor(name, weekDayName, dayParts, weekDayNumber) {
    this.name = name;
    this.weekDayName = weekDayName;
    this.dayParts = dayParts;
    this.weekDayNumber = weekDayNumber
  }
  checkDayParts(dayParts) {
    return true
  }
  checkDayParts(weekDayNUmber) {
    return true
  }
}

class Season {
  constructor(name, days, intervals) {
    this.name = name;
    this.days = days;
    this.intervals = intervals;
  }
}

const dayPartTemplate = Array.from(Array(24).keys()).map(e => {
  return [e, e + 1]

})


let winterDaysArray = [new Day('Зимний_день', 'День_недели', dayPartTemplate, [0, 1, 2, 3, 4, 5, 6])];
let summerDaysArray = [new Day('Летний_день', 'День_недели', dayPartTemplate, [0, 1, 2, 3, 4, 5, 6])];

let winterIntervals = [
  [new Date("1 January 2021"), new Date("1 May 2021")],
  [new Date("1 October 2021"), new Date("1 January 2022")]
];

let summerIntervals = [
  [new Date("1 May 2021"), new Date("1 October 2021")]
];

let winter = new Season('Зима', winterDaysArray, winterIntervals);
let summer = new Season('Лето', summerDaysArray, summerIntervals);

const allYear = [winter, summer];

let result = new Map();
let currHour;

for (let season of allYear) {

  for (let [start, end] of season.intervals) {

    let curr = new Date(start.getTime());
    let index = GetIndexByDate(loadData, start)
    let currHourIndex = 1;

    nextHour:
    while (curr < end) {

      let currDayType = curr.getDay();
      let currHour = curr.getHours();

      for (const day of season.days) {

        if (day.weekDayNumber.includes(currDayType)) {

          for (const [partStart, partEnd] of day.dayParts) {

            if (currHour >= partStart && currHour < partEnd) {

              const key = `${season.name}_${day.name}_${day.weekDayName}_${partStart}_${partEnd}`;

              if (result.has(key)) {
                const res = result.get(key)
                result.set(key, [res[0] + 1, res[1] + loadData[index]['Мощность, МВт']])
              } else {
                result.set(key, [1, loadData[index]['Мощность, МВт']])
              } // конец обновления сегмента в словаре - название, время, мощность

              curr.setHours(curr.getHours() + 1);
              index++;
              currHourIndex++;
              continue nextHour;
            }

          }

        }

      }

      curr.setHours(curr.getHours() + 1);
      index++;
      currHourIndex++;
    }

  }

}


let sumHours = 0; let sumPowers = 0;
result.forEach((value, key) => {
  sumHours += value[0];
  sumPowers += value[1];
})


let postResults = [];
result.forEach((value, key) => {
  postResults.push([key, value[0] / sumHours, value[1] / sumPowers])
})

let toExcel = [];
postResults.forEach((value, key) => {
  toExcel.push({
    'Segment': value[0],
    'Time': value[1],
    'Energy': value[2],
  })
})

let newWB = xlsx.utils.book_new();
let newWS = xlsx.utils.json_to_sheet(toExcel)
xlsx.utils.book_append_sheet(newWB, newWS, 'Results');
xlsx.writeFile(newWB, 'Output.xlsx')


// сфотать полезное

