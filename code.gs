function doGet(request) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .addMetaTag('viewport','width=device-width , initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function globalVariables(){ 
  var varArray = {
    spreadsheetId   : '1kqGc2_TW0M2YcbrU3wkDjU9AXyt0PKmgVP6ujSeHn1U', //** CHANGE !!!
    dataRage        : 'ข้อมูล!A2:AF',                                    //** CHANGE !!!
    idRange         : 'ข้อมูล!A2:A',                                    //** CHANGE !!!
    lastCol         : 'AF',                                            //** CHANGE !!!
    insertRange     : 'ข้อมูล!A1:AF1',                                   //** CHANGE !!!
    sheetID         : '0'                                             //** CHANGE !!! 
  };
  return varArray;
}

/* PROCESS FORM */
function processForm(formObject){  
  if(formObject.RecId && checkID(formObject.RecId)){//Execute if form passes an ID and if is an existing ID
    updateData(getFormValues(formObject),globalVariables().spreadsheetId,getRangeByID(formObject.RecId)); // Update Data
  }
  var result = "";
  if(formObject.searchtext && formObject.searchtext2){
      var seacrhData = formObject.searchtext + formObject.searchtext2;
      result = search(seacrhData);
  }
  return result;

}

function search(searchtext){
  var spreadsheetId = '1kqGc2_TW0M2YcbrU3wkDjU9AXyt0PKmgVP6ujSeHn1U';
  var dataRage  = 'ข้อมูล!A2:AF';
  var data = Sheets.Spreadsheets.Values.get(spreadsheetId, dataRage).values;
  var ar = [];
  
  data.forEach(function(f) {
    if (~f.indexOf(searchtext)) {
      ar.push(f);
    }
  });
  return ar;
}

/* GET FORM VALUES AS AN ARRAY */
function getFormValues(formObject){
/* ADD OR REMOVE VARIABLES ACCORDING TO YOUR FORM*/
  if(formObject.RecId && checkID(formObject.RecId)){
    var values = [[formObject.RecId.toString(),
                  formObject.department,
                  formObject.faculty,
                  formObject.prefix,
                  formObject.name,
                  formObject.lastname,
                  formObject.prefixen,
                  formObject.nameen,
                  formObject.lastnameen,
                  formObject.id,
                  formObject.startid,
                  formObject.endid,
                  formObject.country,
                  formObject.nationality,
                  formObject.ethnicity,
                  formObject.gender,
                  formObject.blood,
                  formObject.dateOfBirth,
                  formObject.email,
                  formObject.telephone,
                  formObject.addressatcard,
                  formObject.provinceatcard,
                  formObject.districtatcard,
                  formObject.subdistrictatcard,
                  formObject.postcodeatcard,
                  formObject.realaddress,
                  formObject.realprovince,
                  formObject.realdistrict,
                  formObject.realsubdistrict,
                  formObject.realpostcode]];
  }else{
    var values = [[new Date().getTime().toString(),//https://webapps.stackexchange.com/a/51012/244121
                  formObject.department,
                  formObject.faculty,
                  formObject.prefix,
                  formObject.name,
                  formObject.lastname,
                  formObject.prefixen,
                  formObject.nameen,
                  formObject.lastnameen,
                  formObject.id,
                  formObject.startid,
                  formObject.endid,
                  formObject.country,
                  formObject.nationality,
                  formObject.ethnicity,
                  formObject.gender,
                  formObject.blood,
                  formObject.dateOfBirth,
                  formObject.email,
                  formObject.telephone,
                  formObject.addressatcard,
                  formObject.provinceatcard,
                  formObject.districtatcard,
                  formObject.subdistrictatcard,
                  formObject.postcodeatcard,
                  formObject.realaddress,
                  formObject.realprovince,
                  formObject.realdistrict,
                  formObject.realsubdistrict,
                  formObject.realpostcode]];
  }
  return values;
}

/*
## CURD FUNCTIONS ----------------------------------------------------------------------------------------
*/

/* CREATE/ APPEND DATA */
function appendData(values, spreadsheetId,range){
  var valueRange = Sheets.newRowData();
  valueRange.values = values;
  var appendRequest = Sheets.newAppendCellsRequest();
  appendRequest.sheetID = spreadsheetId;
  appendRequest.rows = valueRange;
  var results = Sheets.Spreadsheets.Values.append(valueRange, spreadsheetId, range,{valueInputOption: "RAW"});
}

/* READ DATA */
function readData(spreadsheetId,range){
  var result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

/* UPDATE DATA */
function updateData(values,spreadsheetId,range){
  var valueRange = Sheets.newValueRange();
  valueRange.values = values;
  var result = Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, range, {
  valueInputOption: "RAW"});
}

/* CHECK FOR EXISTING ID, RETURN BOOLEAN */
function checkID(ID){
  var idList = readData(globalVariables().spreadsheetId,globalVariables().idRange).reduce(function(a,b){return a.concat(b);});
  return idList.includes(ID);
}

/* GET DATA RANGE A1 NOTATION FOR GIVEN ID */
function getRangeByID(id){
  if(id){
    var idList = readData(globalVariables().spreadsheetId,globalVariables().idRange);
    for(var i=0;i<idList.length;i++){
      if(id==idList[i][0]){
        return 'ข้อมูล!A'+(i+2)+':'+globalVariables().lastCol+(i+2);
      }
    }
  }
}

/* GET RECORD BY ID */
function getRecordById(id){
  if(id && checkID(id)){
    var result = readData(globalVariables().spreadsheetId,getRangeByID(id));
    return result;
  }
}

/* GET ROW NUMBER FOR GIVEN ID */
function getRowIndexByID(id){
  if(id){
    var idList = readData(globalVariables().spreadsheetId,globalVariables().idRange);
    for(var i=0;i<idList.length;i++){
      if(id==idList[i][0]){
        var rowIndex = parseInt(i+1);
        return rowIndex;
      }
    }
  }
}

/* INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
