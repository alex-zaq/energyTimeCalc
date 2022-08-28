'use strict'

let xlsx = require("xlsx");
let wb = xlsx.readFile('Input.xlsx')
let sh1 = wb.Sheets['data'];
let loadData = xlsx.utils.sheet_to_json(sh1);
let year = 2021;
// let startDate = new Date(2021, 0, 1, 0, 0, 0);

loadData.forEach((value, i) => {
  value['date'] = new Date(new Date(year, 0, 1, 0, 0, 0).setHours(i));
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
  constructor( weekDayName, dayParts, weekDayNumber) {
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

let seasons = ['winter', 'summer']
let Daytypes = ['WorkDay']

let mouthPrioritet = {};
let DayTypePrioritet = {};

seasons.forEach((value, index) => {
 mouthPrioritet[value]=index
});

Daytypes.forEach((value, index) => {
  DayTypePrioritet[value]=index
});



const listOfDayArrays = [];
listOfDayArrays.push (new Day( 'WorkDay', dayPartTemplate, [0 ,1, 2, 3, 4, 5, 6]));


const winterIntervals = [];
winterIntervals.push([
  new Date(year, 1, 1), new Date(year, 6, 1),
  new Date(year, 9, 1), new Date(year, 12, 1)
]);

const summerIntervals = [];
summerIntervals.push([
  new Date(year, 6, 1), new Date(year, 9, 1)
]);

const allYear = []
allYear.push(new Season(seasons[0], listOfDayArrays, winterIntervals));
allYear.push(new Season(seasons[1], listOfDayArrays, summerIntervals));


console.log(allYear);

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

              const key = `${season.name}_${day.weekDayName}_${partStart}_${partEnd}`;

              if (result.has(key)) {
                const res = result.get(key)
                result.set(key, [res[0] + 1, res[1] + loadData[index]['Мощность, МВт']])
              } else {
                result.set(key, [1, loadData[index]['Мощность, МВт']])
              } 

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


postResults.sort( (value1, value2) =>
{
   const [season1,daytype1,start1] = value1[0].split('_');
   const [season2,daytype2,start2] = value2[0].split('_');

   if (mouthPrioritet[season1] > mouthPrioritet[season2]) {
     return 1
    } else
    if (mouthPrioritet[season1] < mouthPrioritet[season2]) {
     return -1   
    } else {
      if (daytype1 > daytype2){
        return -1
      } 
      else if (daytype1 < daytype2){
        return 1
      } else
      {
         if ((+start1)<(+start2)){ 
            return -1
         } else
          if ((+start1)>(+start2)){ 
            return 1
          }
          else
          return 0;
      } 
    
}})




let toExcel = [];
postResults.forEach((value, key) => {
  if (key.toString().length == 1) {key ='00'+key}
  if (key.toString().length == 2) {key ='0'+key}
  toExcel.push({
    'Segment':`S${key}`,
    'Segment_name': value[0],
    'Time': value[1],
    'Energy': value[2],
  })
})

let newWB = xlsx.utils.book_new();
let newWS = xlsx.utils.json_to_sheet(toExcel)
xlsx.utils.book_append_sheet(newWB, newWS, 'Results');
xlsx.writeFile(newWB, 'Output.xlsx')


// let excApp = new ActiveXObject("Excel.Application");

// excApp.visible = true;

// var excBook = excApp.Workbooks.open("Output.xlsx");

// idTmr = window.setInterval("Cleanup();",1000);



// сфотать полезное

