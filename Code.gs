var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var inventorySheet = spreadsheet.getSheetByName('Компоненты');
var attributeSheet = spreadsheet.getSheetByName('Атрибуты');
var equipmentSheet = spreadsheet.getSheetByName('Оборудование');

var compPropRange = spreadsheet.getRangeByName('Properties_Компоненты');
var attrPropRange = spreadsheet.getRangeByName('Properties_Атрибуты');
var typesRange = spreadsheet.getRangeByName("Type_Names");

var totalColumns = inventorySheet.getDataRange().getLastColumn();
var totalRows = inventorySheet.getDataRange().getLastRow();
var ui = SpreadsheetApp.getUi();
var uniqueTypes = [];
var propertyList = formPropertyList();

var FIRST_DATA_ROW = 2;
var FIRST_PROP_COLUMN = 5;

var cellTypes = inventorySheet.getRange(FIRST_DATA_ROW, 3, totalRows, 1).getValues();
var compTypes = inventorySheet.getRange(FIRST_DATA_ROW, 4, totalRows, 1).getValues();

//Logger = BetterLog.useSpreadsheet('1r_X2l6F8CtcMG3yxqWmhxfehUpdZiTyQJ-lBxPu2mB4');

function onOpen() {
  createMyMenus();
  //updatePropertyOptions();
  showPropertySidebar();
};

String.prototype.capitalizeFirstLetter = function() {
    return this.charAt(0).toUpperCase() + this.slice(1);
}

function Writelog(object) {
  Logger.log(object);
};

function filterByParentType(type) {
  //ui.alert("filtering by :"+type);
  hideAllRows();
  
  //ui.alert(cellTypes);
  //inventorySheet.hideRows(3, totalRows-1);
  var startShow = -1;
  var r;
  for(r=0;r<cellTypes.length;r++) {
    var currCellType = (String(cellTypes[r]).split(" "))[0];
    var currCompType = (String(compTypes[r]).split(" "))[0];
    
    if(currCellType == type || currCompType == type) {
      if(startShow == -1)
        startShow = r+FIRST_DATA_ROW;
    } else {
      if(startShow != -1) {
        inventorySheet.showRows(startShow, r-startShow+FIRST_DATA_ROW);
        startShow = -1;
      }
    }
  }
  if(startShow != -1) {
    inventorySheet.showRows(startShow, totalRows-startShow);
    startShow = -1;
  }
  //ui.alert("done");
};

function filterByProperties(props) {
  Logger.log('Filtering by props' + props)
  hideAllColumns();
  var columns = getPropertyColumns();
  Logger.log('Columns ' + columns+ 'length = '+columns.length)
  for(var c=FIRST_PROP_COLUMN-1;c<columns.length;c++) {
    Logger.log(c+': Checking '+columns[c]+' against props...');
    if(arrContains(props, columns[c])) {
      inventorySheet.showColumns(c+1);
      Logger.log(c+': Hiding '+columns[c]);
    }
  }
}

function arrContains(array, obj) {
  Logger.log('Checking if '+array);
  Logger.log('Contains '+obj);
  for(var i=0;i<array.length;i++)
    if(array[i]==obj) return true;
  return false;
}

function showAllRows() {
  inventorySheet.showRows(1,totalRows);  
};

function hideAllRows() {
  inventorySheet.hideRows(FIRST_DATA_ROW,totalRows-FIRST_DATA_ROW+1);  
};

function showAllColumns() {
  inventorySheet.showColumns(1,totalColumns);  
};

function hideAllColumns() {
  inventorySheet.hideColumns(FIRST_PROP_COLUMN,totalColumns-FIRST_PROP_COLUMN+1);  
};

function fixCellNumbers() {
  //var cellNumbers = inventorySheet.getRange(4, 1, inventorySheet.getDataRange().getLastRow()-3, 4);
  var cellNumbers = inventorySheet.getRange(4, 1, 30, 4);
  var values = cellNumbers.getValues();
  var savedNum = values[0][0];
  var savedType = values[0][2];

  for(var i = 1;i<=totalRows-3;i++) {
    if(values[i][0]=="") {
      if(values[i][3] !="") {
        cellNumbers.getCell(i+1,1).setValue(savedNum);
        cellNumbers.getCell(i+1,3).setValue(savedType);
      }
      else {
        continue;
      }
    }
    else
    {
      savedNum = values[i][0];
      savedType = values[i][2];
    }
  }
};

function formTypeList() {
  Logger.log('Forming type list');
  var parentTypes = inventorySheet.getRange(3, 3,inventorySheet.getDataRange().getLastRow()-2 , 1);
  var values = parentTypes.getValues();
  uniqueTypes = trunkateToOneWord(values);
  uniqueTypes = eliminateDuplicates(uniqueTypes);
};

function formPropertyList() {
  Logger.log('Forming property list');
  var propertySheet = spreadsheet.getSheetByName('Атрибуты');
  return propertySheet.getSheetValues(2, 1, propertySheet.getDataRange().getLastRow()-1, 1);
};

function eliminateDuplicates(arr) {
  var i,
      len=arr.length,
      out=[],
      obj={};

  for (i=0;i<len;i++) {
    if(arr[i]=="") continue;
    obj[arr[i]]=0;
  }
  for (i in obj) {
    out.push(i);
  }
  return out;
};

function formCompCatTree() {
  var data = inventorySheet.getRange(2,2,inventorySheet.getLastRow()-1,3).getValues();
  if(spreadsheet.getSheetByName('Дерево_Компоненты') === null)
    spreadsheet.insertSheet('Дерево_Компоненты');
  var outputSheet = spreadsheet.getSheetByName('Дерево_Компоненты');
  formCatTree(data, outputSheet);
};

function formEquipCatTree() {
  var data = equipmentSheet.getRange(3,2,equipmentSheet.getLastRow()-1,3).getValues();
  if(spreadsheet.getSheetByName('Дерево_Оборудование') === null)
    spreadsheet.insertSheet('Дерево_Оборудование');
  var outputSheet = spreadsheet.getSheetByName('Дерево_Оборудование');
  formCatTree(data, outputSheet);
};

function formCatTree(data, outputSheet) {
  Logger.log('Starting');
        
  var catCount = 0;
  var objCount = 0;
  var typeCount = 0;
  
  var catTree = {};
  for(var i in data) {
    var cat = data[i][0]||"Нет категории";
    var obj = data[i][1]||"Нет объекта";
    var type = data[i][2]||"Нет типа";
    
    if(typeof catTree[cat] == 'undefined') {
      catTree[cat] = {};
      catTree[cat][obj] = {};
      catTree[cat][obj][type] = 1;
      catCount++;
      objCount++;
      typeCount++;
    }
    else {
      if(typeof catTree[cat][obj] == 'undefined') {
        catTree[cat][obj] = {};
        catTree[cat][obj][type] = 1;
        objCount++;
        typeCount++;
      }
      else {
        if(typeof catTree[cat][obj][type] == 'undefined') {
          catTree[cat][obj][type] = 1;
          typeCount++;
        }
        else {
          catTree[cat][obj][type]++;
        }
      }
    }
  }
  
  Logger.log(catCount+' categories, '+objCount+' objects, '+typeCount+' types');
  Logger.log('Wrting to sheet...');
  writeTreeToSheet(catTree, catCount, objCount, typeCount, outputSheet);
  Logger.log('Done');
};

function writeTreeToSheet(catTree, catCount, objCount, typeCount, treeSheet) {
  var treeRange = treeSheet.getRange(1, 1, catCount+objCount+typeCount, 4);
  var row = 1;
  
  for(var cat in catTree) {
    var catRange = treeSheet.getRange(row,1,1,3).merge().setHorizontalAlignment("left").setBackgroundRGB(200,247,163);
    catRange.setValue(cat);
    row++;
    //treeRange.getCell(row++,1).setValue(cat);
    for(var obj in catTree[cat]) {
      var objRange = treeSheet.getRange(row,1,1,3).breakApart();
      objRange = treeSheet.getRange(row,2,1,2).mergeAcross().setHorizontalAlignment("center").setBackgroundRGB(200,247,230);
      objRange.setValue(obj);
      row++;
      //treeRange.getCell(row++, 2).setValue(obj);
      for(var type in catTree[cat][obj]) {
        treeSheet.getRange(row,1,1,3).merge().setValue(type).setHorizontalAlignment("right");
        treeRange.getCell(row, 4).setValue(catTree[cat][obj][type]);
        row++;
      }
    }
  }
};

function trunkateToOneWord(arr) {
  var i;
  var len=arr.length;
  var out=[];
  
  for(i=0;i<len;i++)
  {
    var perword = String(arr[i]).split(" ");
    out.push(perword[0]);
  }
  return out;
};


function createMyMenus() {
  ui.createMenu('Фильтр')
  .addItem('Показать сайдбар для фильрации','showPropertySidebar')
  .addSeparator()
  .addItem('Отсортировать столбцы по возрастанию','sortByColumns')
  .addItem('Отсортировать столбцы по убыванию','sortByColumns_descending')
  .addSeparator()
  .addSubMenu(ui.createMenu('Составить дерево категорий...')
              .addItem('Компонентов','formCompCatTree')
              .addItem('Оборудования','formEquipCatTree'))
  .addToUi();
};


function showFilterSidebar() {
  formTypeList();
  var html = HtmlService.createTemplateFromFile('FilterSidebar').evaluate()
      .setTitle('Filter')
      .setWidth(300);
      ui.showSidebar(html);
};


function showPropertySidebar() {
  formPropertyList();
  var html = HtmlService.createTemplateFromFile('AsyncProperties').evaluate()
      .setTitle('Фильтр по атрибутам')
      .setWidth(300);
      ui.showSidebar(html);
};

