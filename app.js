function Main() {
    // source we're copying off of
    var source = SpreadsheetApp.getActiveSpreadsheet();
    // create output sheet
    var newSheet = source.getSheetByName('Output');
    if (newSheet != null) {
        source.deleteSheet(newSheet);
    }
    newSheet = source.insertSheet();
    newSheet.setName('Output');
    //copy from pivot table to output sheet
    var source_sheet = source.getSheetByName('Master Pivot Table');
    // get full data range
    var source_range = source_sheet.getDataRange();
    // get A1 notation Identifying range
    var A1Range = source_range.getA1Notation();
    // get data values in range
    var source_data = source_range.getValues();
    // copy values to new sheet
    newSheet.getRange(A1Range).setValues(source_data);

    deleteRows(newSheet);
    formattingRows(newSheet);
    formattingSheet(newSheet);
}

function deleteRows(newSheet) {
    // getting data from Output Sheet
    var rangeData = newSheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    console.log("delete rows last column: " + lastColumn);
    var lastRow = rangeData.getLastRow();
    console.log("delete rows last row: " + lastRow);
    var searchRange = newSheet.getRange(1, 2, lastRow, lastColumn - 1);
    var rangeValues = searchRange.getValues();
    console.log("delete rows rangeValues: " + rangeValues);

    //delete rows that are completely 0
    var rowDeletion = newSheet.getRange(1, 2, lastRow, lastColumn - 1).getValues();
    rowDeletion.splice(0, 1);
    var rowDeletionSize = rowDeletion.length;
    function greaterThanZero(gtz) {
        return gtz > 0;
    }
    for (i = rowDeletionSize - 1; i >= 0; i--) {
        if (rowDeletion[i].some(greaterThanZero) === false) {
            newSheet.deleteRow(i + 2);
        }
    }
}

function formattingRows(newSheet) {
    // getting data for formatting
    var rangeData = newSheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    console.log("formatting rows last column: " + lastColumn);
    var lastRow = rangeData.getLastRow();
    console.log("formatting rows last row: " + lastRow);
    var searchRange = newSheet.getRange(1, 2, lastRow, lastColumn - 1);
    var rangeValues = searchRange.getValues();
    console.log("formatting rows rangeValues: " + rangeValues);
    //scrubbing through data, changing 1, 0, CountA
    //looping through and checking values
    for (i = 0; i < lastColumn; i++) {
        for (j = 0; j < lastRow; j++) {
            if (rangeValues[j][i] >= 1) {
                newSheet.getRange(j + 1, i + 2).setValue('?');
                newSheet.getRange(j + 1, i + 2).setHorizontalAlignment("center");
                newSheet.getRange(j + 1, i + 2).setVerticalAlignment("center");
                newSheet.getRange(j + 1, i + 2).setFontColor('#d2515e');
            } else if (rangeValues[j][i] === 0) {
                newSheet.getRange(j + 1, i + 2).setValue('-');
                newSheet.getRange(j + 1, i + 2).setHorizontalAlignment("center");
                newSheet.getRange(j + 1, i + 2).setVerticalAlignment("center");
                newSheet.getRange(j + 1, i + 2).setFontColor('#d2515e');
            }
        }
    }
    // getting rid of countA
    var values = newSheet.getDataRange().getValues();
    for (u = 2; u < lastColumn; u++) {
        var replaced_Values = values[0][u].toString().slice(10);
        values[0][u] = replaced_Values;
    }
    newSheet.getDataRange().setValues(values);
}

function formattingSheet(newSheet) {
    var rangeData = newSheet.getDataRange();


    //format header
    var allCells = newSheet.getDataRange();
    var header = newSheet.getRange('1:1');
    allCells.setFontSize(11);
    header.setFontWeight("bold");
    header.setFontColor("white");
    allCells.setWrap(true);



    //deletes unnecessary rows and columns
    newSheet.deleteColumn(1);

    //formats
    rangeData.applyRowBanding()
        .setHeaderRowColor('#929497')
        .setFirstRowColor('#D3D3D3')
        .setSecondRowColor('#FFFFFF')

    // delete last column
    var newLastColumn = newSheet.getDataRange().getLastColumn();
    newSheet.deleteColumn(newLastColumn + 1);
}




