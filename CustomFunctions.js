/**
getKegReturnString

@customfunction
*/
function getKegReturnString(sixtels, halfs) {
  return "(" +sixtels +" x 1/6, " +halfs +" x 1/2)";
} //getKegReturnString

/**
getKegDepositNumbersByType

@customfunction
*/
function getKegDepositNumbersByType(cells) {
  var sixtels = 0;
  var halfs = 0;
  var total = 0;
  var value;
  var row;
  
  if(Array.isArray(cells) && cells.length > 0) {
    console.log(cells.length +" : " +cells[0].length +" : " +cells[1].length);
    for(row = 0; row < cells.length; row++) {
      value = cells[row][1];
      if(value == sixtelProductString) {
        sixtels += cells[row][0];
      } else if(value == halfBblProductString) {
        halfs += cells[row][0];
      }
      console.log(value +" sixtels: " +sixtels +" halfs: " +halfs);
    }
  }
  
  return "(" +sixtels +" x 1/6, " +halfs +" x 1/2)";
}

/**
getKegDepositNumber

@customfunction
*/
function getKegDepositNumber(cells) {
  var sixtels = 0;
  var halfs = 0;
  var total = 0;
  var value;
  var row;
  
  if(Array.isArray(cells) && cells.length > 0) {
    console.log(cells.length +" : " +cells[0].length +" : " +cells[1].length);
    for(row = 0; row < cells.length; row++) {
      value = cells[row][1];
      if(value == sixtelProductString) {
        sixtels += cells[row][0];
      } else if(value == halfBblProductString) {
        halfs += cells[row][0];
      }
      console.log(value +" sixtels: " +sixtels +" halfs: " +halfs);
    }
  }
  
  return sixtels+halfs;
}



/**
@customfunction
*/
function copyCell(inValue) {
  var outValue = Object.assign({}, inValue);
  
  return outValue;
}

/**
yearlySum
@param dateCol - Single Column range containing dates
@param sumCol - Single Column range containing amounts to sum
@param year - 4 digit year
@return single value sum

@customfunction
*/
function yearlySum(dateCol, sumCol, year) {
  var dateString;
  var date;
  var value;
  var row;
  var total = 0;
  
  console.log("in yearlySum");
  
  if(Array.isArray(dateCol) && dateCol.length > 0 && Array.isArray(sumCol)) {
    console.log(dateCol.length +" : " +dateCol[0].length +" : " +dateCol[1].length);
    for(row = 0; row < dateCol.length; row++) {
      dateString = dateCol[row][0];
      date = new Date(dateString);
      value = sumCol[row][0];
      var type = typeof dateString;
      console.log("Cell Type: " +type  +" year = " +date.getFullYear() +" value = " +sumCol[row][0] +" total = " +total);
       
      if(year == date.getFullYear()) {
        total += value;
      
      }
      //console.log(value +" sixtels: " +sixtels +" halfs: " +halfs);
      
    }
  } else {
    console.log("columns are not array or zero length");
  }
  
  return total;
}

/**
quarterlySum
@param dateCol - Single Column range containing dates
@paramsumCol - Single Column range containing amounts to sum
@paramquarter - 1, 2, 4, or 4
@param year - 4 digit year
@return single value sum

@customfunction
*/
function quarterlySum(dateCol, sumCol, quarter, year) {
  var dateString;
  var date;
  var value;
  var row;
  var month;
  var total = 0;
  
  console.log("in quarterlySum");
  
  if(Array.isArray(dateCol) && dateCol.length > 0 && Array.isArray(sumCol)) {
    console.log(dateCol.length +" : " +dateCol[0].length +" : " +dateCol[1].length);
    for(row = 0; row < dateCol.length; row++) {
      dateString = dateCol[row][0];
      date = new Date(dateString);
      month = date.getMonth();
      value = sumCol[row][0];
      console.log(" month = " +date.getMonth() +" value = " +sumCol[row][0]);
       
      if(year == date.getFullYear()) {
        if(quarter == 1) {
          if(month == 0 || month == 1 || month == 2)
            total += value;        
        } else if(quarter == 2) {
          if(month == 3 || month == 4 || month == 5)
            total += value;  
        } else if(quarter == 3) {
          if(month == 6 || month == 7 || month == 8)
            total += value;  
        } else if(quarter == 4) {
          if(month == 9 || month == 10 || month == 11)
            total += value;  
        }
      }
    }
  } else {
    console.log("columns are not array or zero length");
  }
  
  return total;
}

/**
monthlySum
@param dateCol - Single Column range containing dates
@param sumCol - Single Column range containing amounts to sum
@param month - month number 0 - 11
@param year - 4 digit year
@return single value sum

@customfunction
*/
function monthlySum(dateCol, sumCol, inMonth, year) {
  var dateString;
  var date;
  var value;
  var row;
  var month;
  var total = 0;
  
  console.log("in monthlySum");
  
  if(Array.isArray(dateCol) && dateCol.length > 0 && Array.isArray(sumCol)) {
    console.log(dateCol.length +" : " +dateCol[0].length +" : " +dateCol[1].length);
    for(row = 0; row < dateCol.length; row++) {
      dateString = dateCol[row][0];
      date = new Date(dateString);
      month = date.getMonth();
      value = sumCol[row][0];
      console.log(" month = " +date.getMonth() +" value = " +sumCol[row][0]);
       
      if(year == date.getFullYear()) {
        if(month == inMonth)
          total += value;
      }
    }
  } else {
    console.log("columns are not array or zero length");
  }
  
  return total;
}