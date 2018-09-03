const parse = require('csv-parse');
const transform = require('stream-transform');
const fs = require('fs');
const excel = require('node-excel-export');
const sanitize = require("sanitize-filename");

const file = process.argv[2];
const fileDelimiter = ',';

var parser = parse({
  delimiter: fileDelimiter,
  columns: true,
  skip_lines_with_error: false,
});


var teams         = [];
var teamPassings  = {};
var greenFlag     = false;
var finishFlag    = false;
var transformer   = transform(function(currentPassing){
  // Skip until after green flag
  if(!greenFlag){
    if(currentPassing["Naam"] === "Groene Vlag"){
        greenFlag = true;
    }
    return;
  }

  // Skip extra flag
  if(currentPassing["Naam"] === "Extra Vlag"){
    return;
  }

  // Skip everything on finish flag
  if(currentPassing["Naam"] === "Finish Vlag" || finishFlag){
    finishFlag = true;
    return;
  }

  // Get team identifier
  var team = currentPassing["Nr."].split("-")[0];

  // Check for existing team passings
  if(!(team in teamPassings)){
      // No passings, init group for passings and add team to teams array
      teamPassings[team] = [currentPassing];
      teams.push(team);
      return;
  }

  // calculate laptime for previous team passing
  var previousTimeMs  = timeToMs(teamPassings[team].slice(-1)[0]["Huidige Tijd"]);
  var currentTimeMs   = timeToMs(currentPassing["Huidige Tijd"]);

  // Account for day jump if it occurs
  if(previousTimeMs > currentTimeMs){
      currentTimeMs += 86400000;
  }

  // Set calculated laptime for previous passing
  currentPassing["Calculated laptime"] = laptime(previousTimeMs, currentTimeMs);
  currentPassing["Lap started at"] = teamPassings[team].slice(-1)[0]["Huidige Tijd"];

  // Add the current passing to the list of team passings
  teamPassings[team].push(currentPassing);
});

transformer.on('finish', function(){
  // Build Excel sheet for each team
  teams.forEach(function(team){
    printPassings = teamPassings[team].slice(1);//Remove first passing since we store laptimes in the second passing
    const report = excel.buildExport(
      [
        {
          name: 'Rondetijden 24KIKA 2018 team ' + team,
          specification: {
              "Lap started at": {
              displayName: 'Starttijd',
                  headerStyle: {},
                  width:125
              },
            "Naam": {
              displayName: 'Naam',
                headerStyle: {},
              width: 190
            },
            'Calculated laptime': {
              displayName: 'Rondetijd',
                headerStyle: {},
              width: 125
            },
          },
          data: printPassings
        }
      ]
    );
    let fileName = sanitize("Rondetijden 24KIKA 2017 team " + team) + '.xlsx';
    console.log("Writing file \"" + fileName + "\"");
    fs.writeFileSync(fileName,  report, 'utf-8');
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


fs.createReadStream(file).pipe(parser).pipe(transformer);
