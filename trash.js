//try {
//    var body = JSON.parse(req.body);
//    var columns=body.columns;   console.log("columns=",columns);
//    var rows=body.rows;         console.log("rows=",rows);
//}catch(e){
//    res.sendStatus(500);                                                                                        console.log("Impossible to parse data! Reason:"+e);
//    return;
//}
//var uniqueFileName = util.getUIDNumber();
//
//var fname = path.join(__dirname, '../../XLSX_temp/' + uniqueFileName + '.xlsx');
//
//try {
//    fs.writeFileSync(fname);
//} catch (e) {
//    res.sendStatus(500);                                                                                        console.log("writeFileSync error=", e);
//    return;
//}
//
//var wb=
//
//var worksheetColumns = [];
//for(var i = 0; i < columns.length; i++){
//    var header=columns[i];
//    worksheetColumns.push({wpx: header.width, v:header.name} );
//}
//
//= {
//    SheetNames: ['Лист_1'],
//    Sheets:  {'Лист_1': {'!ref': 'A1:', '!cols': worksheetColumns}}
//};
//var lineNum = 1;
//
//fillTableData(0, rows);
//
function fillTableData(index, rows) {
    if(!rows[index]) return;
    var rowData=rows[index]; console.log("rowData=",rowData);
    for (var i = 0; i < columns.length; i++) {
        var columnDataID=columns[i].data;           console.log("columnDataID=",columnDataID);
        var displayValue = rowData[columnDataID];   console.log("displayValue=",displayValue);
        // var displayValue = fieldMap[selectedFields[i].displayName];

        console.log("lineNum 330=",lineNum);
        var currentCell = calculateCurrentCellReference(i, lineNum);   console.log("currentCell=",currentCell);
        wb.Sheets['Лист_1'][currentCell] = {
            t: "s",
            v: displayValue,
            s: {
                font: {sz: "11", bold: false},
                alignment: {wrapText: true, vertical: 'top'},
                fill: {fgColor: {rgb: 'ffffff'}},
                border: {
                    left: {style: 'thin', color: {auto: 1}},
                    right: {style: 'thin', color: {auto: 1}},
                    top: {style: 'thin', color: {auto: 1}},
                    bottom: {style: 'thin', color: {auto: 1}}
                }
            }
        }
    }
    lineNum++;
    fillTableData(index++, rows);
}
//
////function calculateCurrentCellReference(index, lineNumber){ console.log("calculateCurrentCellReference index lineNumber=",index,lineNumber);
////    var ALPHA = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P',
////        'Q','R','S','T','U','V','W','X','Y','Z'];
////   // return (index > 25) ? ALPHA[Math.floor((index/26)-1)] + ALPHA[index % 26] + lineNumber : ALPHA[index] + lineNumber;
////    console.log("ALPHA[index]=",ALPHA[index]);
////    return ALPHA[index] + lineNumber;
////}
//
//// var sheet={'Лист_1': {'!ref': 'A1:', '!cols': worksheetColumns}};
//
//// var ws= XLSX.utils.json_to_sheet( {'!ref': 'A1:', '!cols': worksheetColumns});
//// wb.SheetNames.push('Лист_1');
//// wb.Sheets['Лист_1'] = ws;
//
//
//XLSX.writeFile(wb, fname, {bookType:"xlsx"});
//var options = {headers: {'Content-Disposition': 'attachment; filename =out.xlsx'}};
//res.sendFile(fname, options, function (err) {
//    if (err) {
//        res.sendStatus(500);                                                                                    console.log("send xlsx file err=", err);
//        //  return;
//    }
//    // fs.unlinkSync(fname);
//});



// worksheetColumns.push( {wpx: 25} );
//var currentCell = self._calculateCurrentCellReference(i, lineNum);
//workbook.Sheets[spreadsheetName][currentCell] = { t: "s", v: selectedFields[i].displayName, s: { font: { bold: true } } };
