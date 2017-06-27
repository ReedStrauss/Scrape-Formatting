//Create second sheet with formated data for finding emails from current sheet
function getEmailFindingData() {
  //Sheets
  var rawDataSheet = ss.getActiveSheet();
  var lastRow_rawDataSheet = rawDataSheet.getLastRow();

  //header refers to a header and all the data in its columns
  var header = new columnHeaderData(rawDataSheet, [1,1], 0);

  //sort by swim time
  var startDataRow = header.startDataRow;
  var numDataRows = header.numDataRows;
  rawDataSheet.getRange(startDataRow, header.startCol, numDataRows, header.numCols).sort(14);
  
  //Insert formulas/data by overall rank
  //formulas/data
  var divRank = [['Div rank']];
  var profession = [['Profession']];
  var first = [['First']];
  var last = [['Last']];
  var city = [['City']];
  var state = [['State']];
  var email = [['Email']];
  var facebookSearch = [];
  var linkedin = [];
  var fisherpix = [];
  var map = [];
  var lastResortGoogleSearch = [];
  var overallRank = [['Overall Rank']];
  var middle = [['Middle']];
  //format
  //var timeFormat = [['[h]:mm:ss']];
  
  //for all ranks
  var lastRow = header.lastRow;
  for (r=startDataRow; r <= lastRow; r++) {
    var rStr = r.toString();
    
    //div rank
    divRank.push([header.getData('Div rank', r).replace('Div Rank: ', '')]);
    //profession
    profession.push([header.getData('Profession', r)]);
    //name
    var name = header.getData('Name', r);
    var nameArray = name.split(' ');
    var names = nameArray.length;
    if (names == 0) {
      var firstName = '';
      var lastName = '';
      var middleName = '';
    }
    else if (names == 1) {
      var firstName = nameArray[0];
      var lastName = '';
      var middleName = '';
    }
    else if (names == 2) {
      var firstName = nameArray[0];
      var lastName = nameArray[1];
      var middleName = '';
    }
    else {
      var firstName = nameArray[0];
      var lastName = nameArray[names-1];
      var middleName = '';
      for (mn=1; mn < names-1; mn++) {
        if (mn > 1) {
          middleName += ' ';
        }
        middleName += nameArray[mn];
      }
    }
    first.push([firstName]);
    last.push([lastName]);
    middle.push([middleName]);
    //city, state
    var city_state = header.getData('State', r);
    var city_stateArray = city_state.split(' ');
    var city_stateArrayLength = city_stateArray.length;
    var lastElement = city_stateArray[city_stateArrayLength-1];
    if (lastElement.length > 2) {
      var cityName = city_state;
      var stateName = '';
    }
    else {
      var cityName = '';
      var stateName = lastElement;
      for (cs=0; cs < city_stateArrayLength-1; cs++) {
        if (cs > 0) {
          cityName += ' ';
        }
        cityName += city_stateArray[cs];
      }
    }    
    city.push([cityName]);
    state.push([stateName]);
    //email
    email.push(['-']);
    //facebook search
    facebookSearch.push(['=JOIN("", {"https://www.facebook.com/search/str/"}, CONCATENATE(C' + rStr + ', "%20", D' + rStr + ', {"/keywords_users"}))']);
    //linkedin
    linkedin.push(['=JOIN("", {"https://www.linkedin.com/vsearch/f?type=all&keywords="}, CONCATENATE(C' + rStr + ', "+", D' + rStr + '))']);
    //fisherpix
    fisherpix.push(['=JOIN("", {"http://www.finisherpix.com/photos/my-photos/currency/USD/pctrl/Photos/paction/search/pevent/ironman-703-puerto-rico-2015/pbib/"}, O' + rStr + ', {".html"})']);
    //map
    map.push(['=JOIN("", {"https://google.com/maps/place/"}, CONCATENATE(SUBSTITUTE(E' + rStr + '," ","%20"), "%20", F' + rStr + '))']);
    //last resort google search
    lastResortGoogleSearch.push(['=JOIN("", {"https://www.google.com/search?q="}, CONCATENATE(C' + rStr + ', "+", D' + rStr + '))']);
    //overall rank
    overallRank.push([header.getData('Overall Rank', r).replace('Overall Rank: ', '')]);
    //time format
    //timeFormat.push([['[h]:mm:ss']]);
  }
  
  //insert new sheet
  var sheetDivision = header.getData('Division', 2);
  ss.insertSheet(sheetDivision);
  var formattedDataSheet = ss.getActiveSheet();
  
  //insert formulas/data
  //div rank
  formattedDataSheet.getRange(1,1,numDataRows+1,1).setValues(divRank);
  var divRankHeader = formattedDataSheet.getRange(1,1);
  divRankHeader.setFontWeight('bold');
  divRankHeader.setFontLine('underline');
  //profession
  formattedDataSheet.getRange(1,2,numDataRows+1,1).setValues(profession);
  var professionHeader = formattedDataSheet.getRange(1,2);
  professionHeader.setFontWeight('bold');
  professionHeader.setFontLine('underline');
  //first
  formattedDataSheet.getRange(1,3,numDataRows+1,1).setValues(first);
  var firstHeader = formattedDataSheet.getRange(1,3);
  firstHeader.setFontWeight('bold');
  firstHeader.setFontLine('underline');
  //last
  formattedDataSheet.getRange(1,4,numDataRows+1,1).setValues(last);
  var lastHeader = formattedDataSheet.getRange(1,4);
  lastHeader.setFontWeight('bold');
  lastHeader.setFontLine('underline');
  //city
  formattedDataSheet.getRange(1,5,numDataRows+1,1).setValues(city);
  var cityHeader = formattedDataSheet.getRange(1,5);
  cityHeader.setFontWeight('bold');
  cityHeader.setFontLine('underline');
  //state
  formattedDataSheet.getRange(1,6,numDataRows+1,1).setValues(state);
  var stateHeader = formattedDataSheet.getRange(1,6);
  stateHeader.setFontWeight('bold');
  stateHeader.setFontLine('underline');
  //email
  formattedDataSheet.getRange(1,7,numDataRows+1,1).setValues(email);
  var emailHeader = formattedDataSheet.getRange(1,7);
  emailHeader.setFontWeight('bold');
  emailHeader.setFontLine('underline');
  //country
  var countryDataRange = formattedDataSheet.getRange(1,8,numDataRows+1,1);
  var countryColumn = header.getColumnInbetweenRows('Country',1,lastRow);
  countryColumn.copyTo(countryDataRange);
  var countryHeader = formattedDataSheet.getRange(1,8);
  countryHeader.setFontWeight('bold');
  countryHeader.setFontLine('underline');
  //facebook search
  formattedDataSheet.getRange(1,9).setValue('Facebook Search');
  formattedDataSheet.getRange(2,9,numDataRows,1).setFormulas(facebookSearch);
  //linkedin
  formattedDataSheet.getRange(1,10).setValue('LinkedIn');
  formattedDataSheet.getRange(2,10,numDataRows,1).setFormulas(linkedin);
  //fisherpix
  formattedDataSheet.getRange(1,11).setValue('FisherPix');
  formattedDataSheet.getRange(2,11,numDataRows,1).setFormulas(fisherpix);
  //map
  formattedDataSheet.getRange(1,12).setValue('Map');
  formattedDataSheet.getRange(2,12,numDataRows,1).setFormulas(map);
  var mapHeader = formattedDataSheet.getRange(1,12);
  mapHeader.setFontWeight('bold');
  //last resort google search
  formattedDataSheet.getRange(1,13).setValue('Last Resort Google Search')
  formattedDataSheet.getRange(2,13,numDataRows,1).setFormulas(lastResortGoogleSearch);
  //overall rank
  formattedDataSheet.getRange(1,14,numDataRows+1,1).setValues(overallRank);
  var overallRankHeader = formattedDataSheet.getRange(1,14);
  overallRankHeader.setFontWeight('bold');
  overallRankHeader.setFontLine('underline');
  //bib
  var bibDataRange = formattedDataSheet.getRange(1,15,numDataRows+1,1);
  var bibColumn = header.getColumnInbetweenRows('BIB',1,lastRow);
  bibColumn.copyTo(bibDataRange);
  var bibHeader = formattedDataSheet.getRange(1,15);
  bibHeader.setFontWeight('bold');
  bibHeader.setFontLine('underline');
  //division
  var divisionDataRange = formattedDataSheet.getRange(1,16,numDataRows+1,1);
  var divisionColumn = header.getColumnInbetweenRows('Division',1,lastRow);
  divisionColumn.copyTo(divisionDataRange);
  var divisionHeader = formattedDataSheet.getRange(1,16);
  divisionHeader.setFontWeight('bold');
  divisionHeader.setFontLine('underline');
  //swim
  var swimDataRange = formattedDataSheet.getRange(1,17,numDataRows+1,1);
    //swimDataRange.setNumberFormats(timeFormat);
  var swimColumn = header.getColumnInbetweenRows('Swim',1,lastRow);
  swimColumn.copyFormatToRange(formattedDataSheet, 17,17,1,lastRow);
  swimColumn.copyTo(swimDataRange);
  //bike
  var bikeDataRange = formattedDataSheet.getRange(1,18,numDataRows+1,1);
    //bikeDataRange.setNumberFormats(timeFormat);
  var bikeColumn = header.getColumnInbetweenRows('Bike',1,lastRow);
  bikeColumn.copyFormatToRange(formattedDataSheet, 18,18,1,lastRow);
  bikeColumn.copyTo(bikeDataRange);
  //run
  var runDataRange = formattedDataSheet.getRange(1,19,numDataRows+1,1);
    //runDataRange.setNumberFormats(timeFormat);
  var runColumn = header.getColumnInbetweenRows('Run',1,lastRow);
  runColumn.copyFormatToRange(formattedDataSheet, 19,19,1,lastRow);
  runColumn.copyTo(runDataRange);
  //overall time
  var overallTimeDataRange = formattedDataSheet.getRange(1,20,lastRow,1);
    //overallTimeDataRange.setNumberFormats(timeFormat);
  var overallTimeColumn = header.getColumnInbetweenRows('Overall time',1,lastRow);
  overallTimeColumn.copyFormatToRange(formattedDataSheet, 20,20,1,lastRow);
  overallTimeColumn.copyTo(overallTimeDataRange);
  //middle
  formattedDataSheet.getRange(1,21,numDataRows+1,1).setValues(middle);
  formattedDataSheet.getRange(startDataRow, header.startCol, numDataRows, header.numCols).sort(1).setHorizontalAlignment("left");
  var lr = formattedDataSheet.getLastColumn();
  formattedDataSheet.getRange(1,1,1, lr).setFontWeight("Bold")
  formattedDataSheet.setFrozenRows(1);
  
  
  ss.deleteSheet(rawDataSheet);
}