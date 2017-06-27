Column Headers
var columnHeaderData = function(sheet, startCell, headerPadding) {
  //Returns column letters for a given column index (1 to infinity)
  //code taken from http://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
  this.columnIndexToLetters = function(col) {
    var letters = '';
    while (col > 0) {
      //alert(col);
      var lett = (col - 1) % 26;
      letters = String.fromCharCode(lett + 65) + letters;
      col = (col - lett - 1) / 26;
    }
    
    return letters;
  };
  
  //Returns the letter corresponding to a column header
  this.getColumnLetter = function(columnHeader) {
    var colLetter = this.columnLetters[columnHeader];
    if (colLetter == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    
    return colLetter;
  };
  
  //Returns data at a specific cell (under a certain column header, at a specific row)
  this.getData = function(columnHeader, rowIndex) {
    var colIndex = this.columnIndices[columnHeader];
    if (colIndex == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    
    var colData = this.columnData[columnHeader];
    var cellArrayIndex = rowIndex - this.startDataRow;
    var cellData = colData[cellArrayIndex][0];
    return cellData;
  };
  
  //Returns range corresponding to column header (to last row with data)
  this.getColumn = function(columnHeader) {
    var colIndex = this.columnIndices[columnHeader];
    if (colIndex == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    
    return this.sheet.getRange(this.startDataRow, colIndex, this.numDataRows, 1);
  };
  
  //Returns range corresponding to column header (to the inputed row)
  this.getColumnThroughRow = function(columnHeader, lastRowIndex) {
    var colIndex = this.columnIndices[columnHeader];
    if (colIndex == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    var numRows = lastRowIndex - this.startDataRow + 1;
    
    return this.sheet.getRange(this.startDataRow, colIndex, numRows, 1);
  };
  
  //Returns range corresponding to column header (from and to the inputed row)
  this.getColumnInbetweenRows = function(columnHeader, firstRowIndex, lastRowIndex) {
    var colIndex = this.columnIndices[columnHeader];
    if (colIndex == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    var numRows = lastRowIndex - firstRowIndex + 1;
    Logger.log(firstRowIndex + ', ' + colIndex + ', ' + numRows + ', ' + 1);
    return this.sheet.getRange(firstRowIndex, colIndex, numRows, 1);
  };
  
  //Returns range corresponding to column header (defined entirely by input)
  this.getColumnRange = function(firstColumnHeader, lastColumnHeader, firstRowIndex, lastRowIndex) {
    var firstColIndex = this.columnIndices[firstColumnHeader];
    if (firstColIndex == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    var lastColIndex = this.columnIndices[lastColumnHeader];
    if (lastColIndex == undefined) {
      throw new Error('Column Header "' + columnHeader + '" does not exist.');
    }
    var numCols = lastColIndex - firstColIndex + 1;
    var numRows = lastRowIndex - firstRowIndex + 1;
    
    return this.sheet.getRange(firstRowIndex, firstColumnHeader, numRows, numCols);
  };
  
  //Initialize
  this.sheet = sheet;
  var startRow = startCell[0];
  this.startDataRow = startCell[0] + headerPadding + 1;
  this.lastRow = this.sheet.getLastRow();
  this.numDataRows = this.lastRow - this.startDataRow + 1;
  this.startCol = startCell[1];
  this.lastCol = this.sheet.getLastColumn();
  this.numCols = this.lastCol - this.startCol + 1;
  
  if (this.lastCol > 0) {
    var header = this.sheet.getRange(startRow, this.startCol, 1, this.lastCol);
    
    //column letters
    this.columnLetters = {};
    this.columnIndices = {};
    this.columnData = {};
    for (c=this.startCol; c <= this.lastCol; c++) {
      var headerLabel = header.getCell(startRow,c).getValue();
      if (this.columnLetters[headerLabel] == undefined) {
        this.columnLetters[headerLabel] = this.columnIndexToLetters(c);
        this.columnIndices[headerLabel] = c;
        var dataRange = this.sheet.getRange(this.startDataRow, c, this.numDataRows, 1).getValues();
        this.columnData[headerLabel] = dataRange;
      }
      else if (this.columnLetters[headerLabel] != '') {
        throw new Error('Two or more column headers have the same name ' + headerLabel + '.');
      }
    }
  }
  else {
    throw new Error('No data in sheet.');
  }
};