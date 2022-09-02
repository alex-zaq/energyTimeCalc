'use strict'

let xlsx = require("xlsx");
let wb = xlsx.readFile('Input.xlsx')
let sh1 = wb.Sheets['data'];
let loadData = xlsx.utils.sheet_to_json(sh1);
let year = 2021;

const energyTypes = Object.keys(loadData[0]);

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

let seasons = ['winter', 'winter_Bel_NPP_off', 'summer', 'summer_Bel_NPP_off', 'summer_New_NPP_off']
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
  new Date(year, 0, 1), new Date(year, 5, 1),  //  1 января - 1 июня
  new Date(year, 8, 1), new Date(year, 11, 1)  //  1 сентября - 1 декабря
]);

const winterBellNppOfIntervals = [];
winterBellNppOfIntervals.push([
  new Date(year, 11, 1), new Date(year + 1, 0, 1), // 1 декабря - 1 января след. года
]);

const summerIntervals = [];
summerIntervals.push([
  new Date(year, 7, 1), new Date(year, 8, 1), //  1 августа - 1 сентября 
]);


const summerBelNppOffIntervals = [];
summerBelNppOffIntervals.push([
  new Date(year, 6, 1), new Date(year, 7, 1) // 1 июля - 1 августа
]);

const summerNewNppOffIntervals = [];
summerNewNppOffIntervals.push([
  new Date(year, 5, 1), new Date(year, 6, 1) // 1 июня - 1 июля
]);





const allYear = []
allYear.push(new Season(seasons[0], listOfDayArrays, winterIntervals));
allYear.push(new Season(seasons[1], listOfDayArrays, winterBellNppOfIntervals));
allYear.push(new Season(seasons[2], listOfDayArrays, summerIntervals));
allYear.push(new Season(seasons[3], listOfDayArrays, summerBelNppOffIntervals));
allYear.push(new Season(seasons[4], listOfDayArrays, summerNewNppOffIntervals));


// console.log(allYear);



function UpdateSelectedResults(result, key, index, energyNames) {
  
  if (result.has(key)) {
    const res = result.get(key)

      const dataToAdd = [ res[0] + 1];

      
      energyNames.forEach((energyName, i) => {

        dataToAdd.push( res[i+1] +  loadData[index][energyName]   )

      });

    result.set(key, dataToAdd ) 

  } else {

    const dataToAdd = [1];

    energyNames.forEach((energyName, i) => {
      dataToAdd.push(loadData[index][energyName])
    });

    result.set(key, dataToAdd)
  } 

}



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

              
                UpdateSelectedResults(result, key, index, energyTypes)
              
              
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

let sumHours = 0; 
let sumpowers= new Array(energyTypes.length).fill(0);
result.forEach((value, key) => {

  sumHours += value[0];

  energyTypes.forEach((e,i)=>{
      sumpowers[i]+=value[i+1];
  })


})



let postResults = [];
result.forEach((value, key) => {

  const dataToAdd = [key]
 
  dataToAdd.push(value[0]/sumHours)
   
  energyTypes.forEach((e,i)=>{
    dataToAdd.push(value[i+1] / sumpowers[i] )

  })

  postResults.push(dataToAdd)
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
  
  const template = {
    'Сегмент':`S${key}`,
    'Имя сегмента': value[0],
    'Время': value[1],
  }

  energyTypes.forEach( (e, i) => {
    template[`${e}`]=value[i+2]
  })

  toExcel.push(template)
})

let newWB = xlsx.utils.book_new();
let newWS = xlsx.utils.json_to_sheet(toExcel)
xlsx.utils.book_append_sheet(newWB, newWS, 'Results');
xlsx.writeFile(newWB, 'Output.xlsx')





