function onOpen() {
  createMenu();
}

// Creates menu buttons called "Update Team Hub" amd "Update Color Codes"
function createMenu() {
  SpreadsheetApp.getUi().createMenu("Update Team Hub")
    .addItem("Update Team Hub", "UpdateTeamHub")
    .addToUi();

  SpreadsheetApp.getUi().createMenu("Update Color Codes")
    .addItem("Update Color Codes", "UpdateColorCodes")
    .addToUi();
}

function UpdateTeamHub() {
  // Collects Spreadsheet Data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var STARTROW = 4
  var lastRow = 21;
  var playerNames = sheet.getRange(STARTROW, 3, lastRow - STARTROW + 1).getValues().flat(); // filters through player name column

  // Loops through each player's RealGM page
  for (var ticker = 0; ticker < playerNames.length; ticker++) {
    var playerUrl = sheet.getRange(ticker + STARTROW, 28).getValue().toString();

    if (!playerUrl) {
        continue;
    }

    // Fetches the HTML content of the player's RealGM page
    var response = UrlFetchApp.fetch(playerUrl);
    var htmlContent = response.getContentText();

    // Stores the First and Last name of the player
    var name = playerNames[ticker];

    // Stores the Position and Jersey Number of the player (A little buggy)
    var startIdx = htmlContent.indexOf("<h2");
    var endIdx = htmlContent.indexOf("</h2>", startIdx);
    var contentStart = startIdx + 3;
    var bio = htmlContent.substring(contentStart, endIdx);

    if (bio.length > 90){
      var jerseyNumber = bio.split(">")[4].replace("</span", "").replace("#", "");
      var position = bio.split(">")[2].replace("</span", "");
    }
    else {
      var jerseyNumber = "";
      var position = "";
    }

    var exp = extractData(htmlContent, /(\d+)\s*years\s*of\s*NBA\s*service/);

    // Stores the name of the player's team
    var team = extractData(htmlContent, /<strong>Current Team:<\/strong>\s*<a.*?>(.*?)<\/a>/);

    // Stores the birthday of the player
    var birthdate = extractData(htmlContent, /<strong>Born:<\/strong>\s*<a.*?>(.*?)<\/a>/);
    var birthdateObj = new Date(birthdate);
    var currentDate = new Date(); // today's date

    // converting age to ##.# years format 
    var ageInMilliseconds = currentDate - birthdateObj;
    var ageInYears = ageInMilliseconds / (365.25 * 24 * 60 * 60 * 1000);
    var age = ageInYears.toFixed(1);

    // Stores the height of the player
    var heightMatch = extractData(htmlContent, /<strong>Height:<\/strong>\s*([\d'-]+)/);
    var height = formatHeight(heightMatch);

    // Stores the weight of the player
    var weight = extractData(htmlContent, /<strong>Weight:<\/strong>\s*([\d]+)/);

    // Extracts the 2023-24 stats data
   var per36Stats = [];
   var statsPattern = extractData(htmlContent,/<tr[^>]*class\s*=\s*["']per_game["'][^>]*>\s*<td[^>]*>2023-24[\s\S]*?<\/td>\s*<td[^>]*id=["']teamLinenba_reg_.*?<\/td>([\s\S]*?)<\/tr>/);

   var floatValue = null;

  // going through the player stats table
    for (var j = 0; j < statsPattern.split('</td>\n<td>').length; j++) {
      var currentString = statsPattern.split('</td>\n<td>')[j];
      var parsedFloat = parseFloat(currentString);

      if (statsPattern.split('</td>\n<td>')[j].includes(".") && j > 0) {
        floatValue = parsedFloat;
        per36Stats[0] = statsPattern.split('</td>\n<td>')[j-2].slice(5); // gp
        per36Stats[1] = statsPattern.split('</td>\n<td>')[j]; // minutes
        per36Stats[2] = ((statsPattern.split('</td>\n<td>')[j+1]/per36Stats[1])*36).toFixed(1); // pts per 36
        per36Stats[3] = (((statsPattern.split('</td>\n<td>')[j+3]-statsPattern.split('</td>\n<td>')[j+6])/per36Stats[1])*36).toFixed(1); // 2PA per 36
        per36Stats[4] = (((statsPattern.split('</td>\n<td>')[j+2]-statsPattern.split('</td>\n<td>')[j+5])/(statsPattern.split('</td>\n<td>')[j+3]-statsPattern.split('</td>\n<td>')[j+6]))*100).toFixed(1); // 2P%
        per36Stats[5] = ((statsPattern.split('</td>\n<td>')[j+6]/per36Stats[1])*36).toFixed(1); // 3PA per 36
        per36Stats[6] = (statsPattern.split('</td>\n<td>')[j+7]*100).toFixed(1); // 3P%
        per36Stats[7] = ((statsPattern.split('</td>\n<td>')[j+9]/per36Stats[1])*36).toFixed(1); // FTA per 36
        per36Stats[8] = (statsPattern.split('</td>\n<td>')[j+10]*100).toFixed(1); // FT%
        per36Stats[9] = ((statsPattern.split('</td>\n<td>')[j+14]/per36Stats[1])*36).toFixed(1); // ast per 36
        per36Stats[10] = ((statsPattern.split('</td>\n<td>')[j+17]/per36Stats[1])*36).toFixed(1); // tov per 36
        per36Stats[11] = ((statsPattern.split('</td>\n<td>')[j+11]/per36Stats[1])*36).toFixed(1); // oreb per 36
        per36Stats[12] = ((statsPattern.split('</td>\n<td>')[j+12]/per36Stats[1])*36).toFixed(1); // dreb per 36
        per36Stats[13] = ((statsPattern.split('</td>\n<td>')[j+15]/per36Stats[1])*36).toFixed(1); // stl per 36
        per36Stats[14] = ((statsPattern.split('</td>\n<td>')[j+16]/per36Stats[1])*36).toFixed(1); // blk per 36
        per36Stats[15] = ((statsPattern.split('</td>\n<td>')[j+18].split('</td>')[0]/per36Stats[1])*36).toFixed(1); // fls per 36

        break;
      }
    }

    // Write data back into the sheet
    if (sheet.getRange(ticker + STARTROW, 4).getValue() == "") {
      sheet.getRange(ticker + STARTROW, 4).setValue(position);
    }
     sheet.getRange(ticker + STARTROW, 5).setValue(jerseyNumber);
    if (sheet.getRange(ticker + STARTROW, 6).getValue() == "") {
      sheet.getRange(ticker + STARTROW, 6).setValue(exp);
    }

    sheet.getRange(ticker + STARTROW, 7).setValue(age);
    sheet.getRange(ticker + STARTROW, 8).setValue(height);
    sheet.getRange(ticker + STARTROW, 10).setValue(weight);

    for (var i = 0; i < per36Stats.length; i++) {
      if (isNaN(per36Stats[i])){
        sheet.getRange(ticker + STARTROW, 12 + i).setValue(0);
      }
      else{
        sheet.getRange(ticker + STARTROW, 12 + i).setValue(per36Stats[i]);
      }
    }
  }
}

function UpdateColorCodes() {
  var sheets = ["ATL", "BOS", "BRK"]; // update for team sheets you have
  var bottomUpRanges = ["G4:G21", "V4:V21", "AA4:AA21"];
  var topDownRanges = ["J4:J21", "L4:L21", "M4:M21", "N4:N21", "O4:O21", "P4:P21", "Q4:Q21", "R4:R21", "S4:S21", "T4:T21", "U4:U21", "W4:W21", "X4:X21", "Y4:Y21", "Z4:Z21"];
  var heights = ["H4:H21", "I4:I21"];

  // Define color gradients
  var bottomUpColors = [];
  for (var x = 0; x < 10; x++) {
    bottomUpColors.push(gradientColor("#00a950", "#FFFFFF", x / 9)); // From green to white
  }
  for (var y = 0; y < 10; y++) {
    bottomUpColors.push(gradientColor("#FFFFFF", "#e55545", y / 9)); // From white to red
  }

  var topDownColors = [];
  for (var x = 0; x < 10; x++) {
    topDownColors.push(gradientColor("#e55545", "#FFFFFF", x / 9)); // From red to white
  }
  for (var y = 0; y < 10; y++) {
    topDownColors.push(gradientColor("#FFFFFF", "#00a950", y / 9)); // From white to green
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  heights.forEach(function(rangeString) {
    var allValues = [];
    
    // Collect all values from the specified range in all sheets
    sheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] !== "") { // Skip empty cells
          allValues.push(convertHeightToInches(values[i][0]));
        }
      }
    });
    
    // Determine quantile boundaries
    allValues.sort(function(a, b) { return a - b; });
    var quantiles = [];
    var numQuantiles = 20; // Change as needed
    for (var k = 0; k < numQuantiles; k++) {
      var index = Math.floor((k + 1) * allValues.length / numQuantiles) - 1;
      quantiles.push(allValues[index]);
    }
    
    // Apply formatting to each sheet and range
    sheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      
      for (var m = 0; m < values.length; m++) {
        var value = values[m][0];
        if (value === "") continue; // Skip empty cells
        value = convertHeightToInches(value);

        // Determine quantile
        var quantileIndex = 0;
        for (var n = 0; n < quantiles.length; n++) {
          if (value <= quantiles[n]) {
            quantileIndex = n;
            break;
          }
        }
        
        // Set background color based on quantile
        range.getCell(m + 1, 1).setBackground(topDownColors[quantileIndex]);
      }
    });
  });

  // Process each bottom up range
  bottomUpRanges.forEach(function(rangeString) {
    var allValues = [];
    
    // Collect all values from the specified range in all sheets
    sheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] !== "") { // Skip empty cells
          allValues.push(values[i][0]);
        }
      }
    });
    
    // Determine quantile boundaries
    allValues.sort(function(a, b) { return a - b; });
    var quantiles = [];
    var numQuantiles = 20; // Change as needed
    for (var k = 0; k < numQuantiles; k++) {
      var index = Math.floor((k + 1) * allValues.length / numQuantiles) - 1;
      quantiles.push(allValues[index]);
    }
    
    // Apply formatting to each sheet and range
    sheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      
      for (var m = 0; m < values.length; m++) {
        var value = values[m][0];
        if (value === "") continue; // Skip empty cells

        // Determine quantile
        var quantileIndex = 0;
        for (var n = 0; n < quantiles.length; n++) {
          if (value <= quantiles[n]) {
            quantileIndex = n;
            break;
          }
        }
        
        // Set background color based on quantile
        range.getCell(m + 1, 1).setBackground(bottomUpColors[quantileIndex]);
      }
    });
  });

  topDownRanges.forEach(function(rangeString) {
    var allValues = [];
    
    // Collect all values from the specified range in all sheets
    sheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] !== "") { // Skip empty cells
          allValues.push(values[i][0]);
        }
      }
    });
    
    // Determine quantile boundaries
    allValues.sort(function(a, b) { return a - b; });
    var quantiles = [];
    var numQuantiles = 20; // Change as needed
    for (var k = 0; k < numQuantiles; k++) {
      var index = Math.floor((k + 1) * allValues.length / numQuantiles) - 1;
      quantiles.push(allValues[index]);
    }
    
    // Apply formatting to each sheet and range
    sheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      
      for (var m = 0; m < values.length; m++) {
        var value = values[m][0];
        if (value === "") continue; // Skip empty cells

        // Determine quantile
        var quantileIndex = 0;
        for (var n = 0; n < quantiles.length; n++) {
          if (value <= quantiles[n]) {
            quantileIndex = n;
            break;
          }
        }
        
        // Set background color based on quantile
        range.getCell(m + 1, 1).setBackground(topDownColors[quantileIndex]);
      }
    });
  });
}

// Helper function to calculate intermediate colors
function gradientColor(startColor, endColor, ratio) {
  var start = parseInt(startColor.slice(1), 16);
  var end = parseInt(endColor.slice(1), 16);
  
  var r1 = (start >> 16) & 0xFF;
  var g1 = (start >> 8) & 0xFF;
  var b1 = start & 0xFF;
  
  var r2 = (end >> 16) & 0xFF;
  var g2 = (end >> 8) & 0xFF;
  var b2 = end & 0xFF;
  
  var r = Math.round(r1 + (r2 - r1) * ratio);
  var g = Math.round(g1 + (g2 - g1) * ratio);
  var b = Math.round(b1 + (b2 - b1) * ratio);
  
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

function extractData(content, regexPattern) {
  var regex = new RegExp(regexPattern);
  var match = regex.exec(content);
  return match ? match[1] : "";
}

// converting height data to #'##" format
function formatHeight(heightMatch) {
  if (heightMatch) {
    var parts = heightMatch.split('-');
    if (parts.length === 2) {
      var feet = parts[0];
      var inches = parts[1];
      return feet + "'" + inches + "\"";
    }
  }
  return "";
}

function convertHeightToInches(height) {
  var match = height.match(/^(\d+)'(\d+)"$/);
  if (match) {
    var feet = parseInt(match[1], 10);
    var inches = parseInt(match[2], 10);
    return feet * 12 + inches;
  }
}
