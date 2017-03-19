function lookupTypeColumn(type) {  
  /*currType = "Not set";
  currTypeAttrColumn = -1;
  Logger.log('Looking up column for '+type);
  Logger.log('Previous type was: '+currType);
  
  if(currType == type) {
    Logger.log('Column is '+currTypeAttrColumn);
    return currTypeAttrColumn;
  }
  
  Logger.log('Changed currType to '+type);*/
  var types = typesRange.getValues()[0];
  if(types == null) return;
  Logger.log('Looking through '+types.length+' types');
  for(var i = 0;i<types.length;i++) {
    if(types[i] == type) {
      currTypeAttrColumn = i+1;
      Logger.log('Found column '+(i+1));
      return currTypeAttrColumn;
    }
  }
  Logger.log('Could not find column for '+type);
  currTypeAttrColumn = -1;
  return -1;
}

function getPropertyColumns() {
  if(attrPropRange == null) return;
  var properties = attrPropRange.getValues();
  Logger.log('Property columns: '+properties);
  return properties;
}

function getPropertyOptions() {
  if(attrPropRange == null) return;
  properties = attrPropRange.getValues();
  Logger.log('Property options: '+properties);
  return properties;
}

function getPropertyList(type) {
  var properties = [];
  if(typesRange == null) return;
  
  Logger.log('Searching for prop column');
  var lastColumn = typesRange.getLastColumn();
  for(var i = 1;i<=lastColumn;i++) {
    if(typesRange.getCell(1,i).getValue() == type) {
      properties = attributeSheet.getRange(2,i,totalColumns,1);
      break;
    }
  }
  Logger.log('Found prop column');
  if(properties != [])
    properties = properties.getValues();
  var propList = [];
  Logger.log('Forming the array');
  for(var i=0;i<properties.length;i++) {
    if(properties[i] == 1) propList.push(propertyList[i]);
  };
  Logger.log('Formed array, returning');
  return propList;
};

function updateProperty(type, prop, value) {
  Logger.log('Updating : '+type+'\'s '+prop+ ' value :'+value);
  typeColumn = lookupTypeColumn(type);
  
  if(typeColumn == -1) return;
  
  var attrRange = attributeSheet.getDataRange();
  var lastRow = attrRange.getLastRow();
  
  for(var j=2;j<=lastRow;j++) {
    var propCellVal = attrRange.getCell(j,1).getValue();
    if(propCellVal == prop) {
      Logger.log('Found prop: '+ prop+' at: '+j);
      if(value) {
        attrRange.getCell(j,typeColumn).setValue(1);
        Logger.log('Setting');
      }
      else {
        attrRange.getCell(j,typeColumn).clearContent();
        Logger.log('Clearing');
      }
      break;
    }
  }
}

function getTypeList() {  
  if(typesRange == null) return;
  var types = typesRange.getValues()[0];
  Logger.log("Got types from spreadsheet: "+types);
  return types.slice(2,types.indexOf(""));
};

function updatePropertyOptions() {
  Logger.log('Updating property options');
  var propColumns = getPropertyColumns();
  var attributes = attributeSheet.getRange(1,1,totalColumns);
  for(var i=0;i<propColumns.length;i++) {
    attributes.getCell(2+i,1).setValue(propColumns[i]);
  }
}