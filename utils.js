/**
  utils.gs - 
  
*/

/**
  key / value object for fake hashmap
*/
var KeyValue = {
  key:null,
  value:null
};


/** 
  Fake hashmap object with access methods. Based on array of keyValue objects.
*/
var FauxMap = {
  keyValues : [],
  
/**
    get value associated with key parameter
    
    @returns value associated with key or null if not found
  */
  get : function (key) {
  var foundKey = null;
  
    for(i = 0; i < this.keyValues.length; i++) {
      if(this.keyValues[i].key == key)
        return(this.keyValues[i].value);
    }
  
    // key not found
    return null; 
  },
  
  /**
    adds key/value pair to array or replaces value if key exists
    
    @return length of array
  */
  put : function(key, value) {
    var found = false;
  
    for(i = 0; i < this.keyValues.length; i++) {
      if(this.keyValues[i].key == key) {
        this.keyValues[i].value = value;
        //console.log("FauxMap.put: replacing key: " +key +" old value: " +this.keyValues[i].value +" with new value: " +value);
        return this.keyValues.length;
      }
    }
    
    var keyValue = Object.create(KeyValue);
    keyValue.key = key;
    keyValue.value = value;
    this.keyValues.push(keyValue);
    //console.log("FauxMap.put: adding new keyValue - key: " +keyValue.key, + " value: " +keyValue.value);
    return this.keyValues.length;  
  },
  
  /**
    looks for existance of key in array
    
    @return boolean true if found
  */
  containsKey : function (key) {
    
    for(i = 0; i < this.keyValues.length; i++) {
      if(this.keyValues[i].key == key) {
        hasKey = true;
        return true;
      }
    }
    return false;
  },
  
  /**
    return Array of non-null keys in object array
  */
  keys : function () {
    var keys = new Array();
    var key;
    
    for(i = 0; i < this.keyValues.length; i++) {
      key = this.keyValues[i].key;
      console.log("FauxMap.keys - key: " +key);
      if(key != null && key != "") {
        keys.push(key);
      }
    }
    return keys;
  },
  
  /** 
    Delete entry from array
    
    @return index of deleted element or -1 if not found
  */
  deleteKey : function (key) {
    var key = -1;
    
    for(i = 0; i < this.keyValues.length; i++) {
      if(this.keyValues[i].key == key)
        key = i;
    }
  
    if(i != -1)
      this.keyValues.splice(i, 1);

    return i;
  },
  
  /**
    Print array contents as String
  */
  toString : function() {
    var outString =  "FauxMap:" +this.keyValues.length +": "; 
    for(i = 0; i < this.keyValues.length; i++) {
      outString = outString +this.keyValues[i].key +":" +this.keyValues[i].value +"; ";
    }
    
    return outString;
  }
}; //FauxMap

/**
  get index of 1st column with matching name or -1 if not found
  
  @return matching column index or -1
*/
function getColumnNumberByName(sheet, name) {
  var maxColNum = sheet.getLastColumn();
  var range;
  var value;
  
  for(col = 1; col < maxColNum; col++) {
    value = sheet.getRange(1, col).getValue();
    
    //console.log("getColumnNumberByName: name: " +name +" value: " +value +" column: " +col +" max col num: " +maxColNum);
    
    if(value == name)
      return col;
  }
  
  return -1;
} //getColumnNumberByName

/**
  get index of 1st row with matching name or -1 if not found
  
  @return matching column index or -1
*/
function getRowNumberByName(sheet, column, name) {
  var range;
  var value;
  
  for(row = 1; row < sheet.getLastRow(); row++) {
    value = sheet.getRange(row, column).getValue();
    
    console.log("getRowNumberByName(): row: " +row +" col: " +column +" name: " +name +" product name: " +value);
    if(value == name)
      return row;
  }
  
  return -1;
} //getRowNumberByName

/**
   get the value from a single cell
   
   @return cell value
*/
function getCellValue(sheet, row, column) {
  console.log("getCellValue() - row: " +row +" col: " +column);
  var vals = sheet.getSheetValues(row, column, 1, 1);
  var val = vals[0][0];
  
  return val;
}

/**
  set value of cell in sheet
*/
function setCellValue(sheet, row, column, value) {
  var range = sheet.getRange(row, column, 1, 1);
  var cell = range.getCell(0,0);
  cell.setValue(value);
}

/**

*/
function deleteNamedFiles(directory, filename) {
  var file;
  var files = directory.getFilesByName(filename);
   
  while(files.hasNext()) {
    file = files.next();
    console.log("deleting file: " +filename +" in dir: " +directory);
    file.setTrashed(true);
  }
}
