const parse = require('csv-parse');
const transform = require('stream-transform');
const fs = require('fs');
const excel = require('node-excel-export');

const file = process.argv[2];
const fileDelimiter = ',';

var groupedPassings = {};
var teams           = [];
var gotStart        = false;
var finishTimeMs    = null;

var parser = parse({
  delimiter: fileDelimiter,
  columns: true,
  skip_lines_with_error: false,
});

var passingGrouper = transform(function(data){
  // Skip until green flag
  if(!gotStart){
    if(data["Naam"] == "Groene Vlag"){
      gotStart = true;
    }
    return;
  }

  // Save finish flag time
  if(data["Naam"] == "Finish Vlag"){
    finishTimeMs = timeToMs(data["Verstreken Tijd"]);
    return;
  }

  // Get team based on number column
  var team = data["Nr."].split("-")[0];

  // Save team and init group for passings
  if(!(team in groupedPassings)){
    groupedPassings[team] = [];
    teams.push(team);
  }

  // Add data to the group
  groupedPassings[team].push(data);
});

passingGrouper.on('finish', function(){
  teams.forEach(function(team){
    var previousStartTimeMs = 0;
    groupedPassings[team].forEach(function(passing, passingIndex){
      // Save start time and person
      if(passing["Verstreken Tijd"] == ""){
        return true;
      }
      currentTimeToMs = timeToMs(passing["Verstreken Tijd"]);
      if(passingIndex > 0){
        groupedPassings[team][passingIndex-1]["Rondetijd"]=laptime(previousStartTimeMs, currentTimeToMs);
      }
      previousStartTimeMs = currentTimeToMs;
    });
    groupedPassings[team][groupedPassings[team].length-1]["Rondetijd"] = laptime(previousStartTimeMs, finishTimeMs);

    const report = excel.buildExport(
      [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
        {
          name: 'Results', // <- Specify sheet name (optional)
          specification: {
            "Naam": {
              displayName: 'Naam',
              width: 190
            },
            "Rondetijd": {
              displayName: 'Rondetijd',
              width: 125
            },
          },
          data: groupedPassings[team] // <-- Report data
        }
      ]
    );

    fs.writeFileSync("Rondetijden 24KIKA 2018 team " + team + '.xlsx',  report, 'utf-8');
  });
});

const laptime = function(startTimeMs, finishTimeMs){
  return msToTime(laptimeMs(startTimeMs, finishTimeMs));
};

const laptimeMs = function(startTimeMs, finishTimeMs){
  return finishTimeMs - startTimeMs;
};

const msToTime = function(timeInMs, excludeHundreds) {
  var delim = ":";
  var minutes = Math.floor(timeInMs / (1000 * 60));
  var seconds = Math.floor(timeInMs / 1000 % 60);
  var hundreds = timeInMs % 1000;

  minutes = minutes < 10 ? '0' + minutes : minutes;
  seconds = seconds < 10 ? '0' + seconds : seconds;
  var ret = minutes + delim + seconds
  if(!excludeHundreds){
    ret += delim + hundreds;
  }
  return ret;
}

const timeToMs = function(time)
{
    let startTime = time;
    let startTimeParts = startTime.split(":");
    let secondParts = startTimeParts[startTimeParts.length-1].split(".");
    let hrs = startTimeParts[startTimeParts.length-3] || 0;
    let min = startTimeParts[startTimeParts.length-2] || 0;
    let sec = secondParts[0];
    let ms  = secondParts[1];
    return((parseInt(hrs)*60*60+parseInt(min)*60+parseInt(sec))*1000)+parseInt(ms);
}


fs.createReadStream(file).pipe(parser).pipe(passingGrouper);
