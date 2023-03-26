/* globals apex,$ */
window.CETOIG = window.CETOIG || {};
var workbook;
var workSheet;

//Execute script
CETOIG.main = function(config) {

    //variables
    var fileBrowserItemName = config.fileBrowserItemName;
    var iGNameVal = config.igName;
    var fileNameVal = apex.item(config.fileName).getValue();
    var selectListName = config.selectListItem;
    var fileExtensionType = config.extensionType;

    //Import from file
    if(config.mode === "IMPORTPREPARE") {
       CETOIG.importFromFile(fileBrowserItemName,selectListName);
    }
    //Execute Import
    else if(config.mode === "IMPORTEXECUTE") {
        var selectedSheetName = apex.item(selectListName).getValue();
        workSheet = workbook.Sheets[selectedSheetName];
        workSheet = CETOIG.sheetTo2dArray(workSheet);
        CETOIG.insertNewValues(workSheet,iGNameVal);
    }
    //Paste from clipboard
    else if(config.mode === "PASTE") {
        CETOIG.pasteFromClipboardExecute(iGNameVal);
    }
    //export selected rows
    else {
        var fBrowserVal = apex.item(fileBrowserItemName).getValue()[0];
        var fParameterVal = (typeof fBrowserVal !== "undefined") ? fBrowserVal : fileNameVal + '.' + fileExtensionType;
        CETOIG.exportSelectedRows(fParameterVal,iGNameVal,selectListName,fileExtensionType);
    }
}

//Import from file
CETOIG.importFromFile = function(pFileItemName,pSelectList) {
    var file = document.getElementById(pFileItemName).files[0];
    var reader = new FileReader();

    reader.onload = (function (pFile) {
        return function (e) {
            if (pFile) {
                workbook = XLSX.read(e.target.result);
                CETOIG.removeWorksheetEntries(pSelectList);
                CETOIG.setWorksheetEntries(workbook.SheetNames,pSelectList);
            }
        }
    }) (file);
    reader.readAsArrayBuffer(file);
}

//Paste from clipboard - Prepare
CETOIG.pasteFromClipboardPrepare = function() {
    document.onpaste = function(e) {
        var str = e.clipboardData.getData("text/html");
        workbook = XLSX.read(str, {type: "string"});
    };
}

//Paste from clipboard - Execute
CETOIG.pasteFromClipboardExecute = function(pIGName) {
    workSheet = workbook.Sheets[workbook.SheetNames[0]];
    workSheet = CETOIG.sheetTo2dArray(workSheet);
    console.log(workbook);
    CETOIG.insertNewValues(workSheet,pIGName);
}

//Export selected rows
CETOIG.exportSelectedRows = function(pFileName,pIGName,pSelectList,pExtensionType) {
    //init variables
    var preparedArray = [];
    var newRow = {};
    var selectedSheetName = apex.item(pSelectList).getValue();
    var $widget = apex.region(pIGName).widget();
    var $grid = $widget.interactiveGrid('getViews').grid;
    var $model = $grid.model;
    var $modelCols;
    var $modelSelectedRecords = $model.getSelectedRecords();

    //if there is an imported file then use its headers otherwise the ones
    //that's provided by the IG if include header is selected
    if (workbook) {
        workSheet = workbook.Sheets[selectedSheetName];
        workSheet = CETOIG.sheetTo2dArray(workSheet);
        $modelCols = Object.keys(workSheet[0]);
    }
    else {
        $modelCols = CETOIG.removeIgCols(Object.keys($model._options.fields));
    }

    //prepare array
    for(i = 0;i < $modelSelectedRecords.length; i++) {
        for(j = 0;j < $modelSelectedRecords[i].length - 1; j++) {
            newRow[$modelCols[j]] = $modelSelectedRecords[i][j];
        }
        preparedArray.push(newRow);
        newRow = {};
    }

    //if worksheet was selected append data else create new table
    if(workSheet) {
        for(i = 0;i < $modelSelectedRecords.length; i++) { 
            workSheet.push(preparedArray[i]);
        }
    }
    else {
        workSheet = preparedArray;
    }

    //export to table
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.json_to_sheet(workSheet, {header:$modelCols});

    XLSX.utils.book_append_sheet(wb, ws, selectedSheetName ? selectedSheetName : 'sheet');
    XLSX.writeFile(wb, pFileName);

}

//Insert IG rows and paste values
CETOIG.insertNewValues = function(pVals,pIGName) {
    //init variables
    var $widget = apex.region(pIGName).widget();
    var $grid = $widget.interactiveGrid('getViews').grid;
    var $model = $grid.model;
    var $modelCols = CETOIG.removeIgCols(Object.keys($model._options.fields));
    var $modelNewRows = [];
    var $modelCurrentRecord;
    var currentFileRecord;

    //create and store new records
    for(i = 0;i < pVals.length; i++) {
        var newRecordId = $model.insertNewRecord();
        $modelNewRows.unshift(newRecordId);
    }
    
    //set new records
    for(i = 0;i < $modelNewRows.length; i++) {

        $modelCurrentRecord = $model.getRecord($modelNewRows[i]);
        excelRecordKeys = Object.keys(pVals[i]);
        currentFileRecord = pVals[i];

        for(j = 0;j < excelRecordKeys.length; j++) {
            $model.setValue($modelCurrentRecord, $modelCols[j], currentFileRecord[excelRecordKeys[j]]);
        }

    }
}

//Set worksheets to selectlist
CETOIG.setWorksheetEntries = function(pSheets,pItemName) {
    var selectList = document.getElementById(pItemName);
    var option = document.createElement('option');
    for(i = 0;i < pSheets.length; i++) {
        option.value = option.text = pSheets[i];
        selectList.add(option);
    }
}

//Clear worksheet entries in selectlist
CETOIG.removeWorksheetEntries = function(pItemName) {
    var selectList = document.getElementById(pItemName);
    for(i = 0;i < pItemName.length; i++) {
        selectList.remove(i);
    }
}

//Find sheet index
CETOIG.removeIgCols = function(pArray) {
    var removeableCols = ['APEX$ROW_ACTION','_meta'];
    var vArray = pArray;
    for(i = 0;i < removeableCols.length; i++) {
        vArray.splice(vArray.indexOf(removeableCols[i]),1);
    }
    return vArray;
}

//Sheet to 2D Array
CETOIG.sheetTo2dArray = function(pSheetName) {
    return XLSX.utils.sheet_to_json(pSheetName, {defval:""});
}