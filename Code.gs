function Convert_to_RST_Table() {

  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Get Spreadsheet object
  var sheet = ss.getSheets()[0]; // Get the object for the first sheet
  var range = sheet.getDataRange(); // Get the full data range of the sheet (object)
  var values = range.getValues(); // Get all the values from the range as an array

  var numrows, numcols
  numrows = values.length; // Get the number of rows for the table
  numcols = values[0].length; // Get the number of coulmn for the table

  // Determine the max string length for each column
  var max_array = [];
  var cell_str
  for(let jj=0; jj<numcols; jj++) {
    let max = 0; // initialize max number of string in that column
    for(let ii=0; ii<numrows; ii++) {
      cell_str = values[ii][jj]; // get the cell value
      if(isNaN(cell_str)==false) {
        cell_str=cell_str.toString(); // convert the value if it is a number
      }
      if(cell_str.length>max) {
        max = cell_str.length; // Update new max if this value is the longest
      }
    }
    max_array.push(max);
  }

  // Construct the separators
　var horiz_separator1='' // For the normal horiz separator
  for(let jj=0; jj<numcols; jj++) {
    horiz_separator1=horiz_separator1+'+';
    horiz_separator1=horiz_separator1+'-'.repeat(max_array[jj]+2);
  }
  horiz_separator1=horiz_separator1+'+ ';    

　var horiz_separator2 // For the title horiz separator 
  horiz_separator2=horiz_separator1.replace(/-/g,'=') // Regular expression

  // Construct the table text
  var str='';

  for(let ii=0; ii<numrows; ii++) {
    if(ii==1){
      str = str + horiz_separator2 + '\n'; // only for the title row
    }
    else {
      str = str + horiz_separator1 + '\n'; // nominal
    }

    for(let jj=0; jj<numcols; jj++) {
      cell_str = values[ii][jj]; // get the cell value
      if(isNaN(cell_str)==false) {
        cell_str=cell_str.toString(); // convert the value if it is a number
      }

      // construct the contents line
      str = str + '| '
      str=str+cell_str+' '.repeat(max_array[jj]-cell_str.length+1)
      
    }
    str = str + '|\n' // carriage return
  }
  str = str + horiz_separator1 + '\n' // bottom separator

  // Add the result to the second sheet
  var result_sheet = ss.getSheets()[1];
  result_sheet.getRange(1,1).setValue(str)
  SpreadsheetApp.flush();
}
