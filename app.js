var fs = require('fs');
var express = require('express');
var app = express();
var path = require('path');
var XLSX = require('xlsx');
var bodyParser = require('body-parser');
var uid = require('uniqid');
var port = 8183;

app.use('/', express.static('public'));
app.use(bodyParser.json({limit: '5mb'}));
app.use(bodyParser.urlencoded({limit: '5mb'}));
app.use(bodyParser.text());


app.get('/', function (req, res) {

    try {
        var body = JSON.parse(fs.readFileSync(path.join(__dirname, '/data.json')));
        var columns = body.columns;
        console.log("columns=", columns);
        var rows = body.rows;
        console.log("rows=", rows);
    } catch (e) {
        res.sendStatus(500);
        console.log("Impossible to parse data! Reason:" + e);
        return;
    }
    var uniqueFileName = getUIDNumber();
    var headers = [];
    for (var j in columns) {
        headers.push(columns[j].data);
    }
    var fname = path.join(__dirname, './XLSX_temp/' + uniqueFileName + '.xlsx');
    try {
        fs.writeFileSync(fname);
    } catch (e) {
        res.sendStatus(500);
        console.log("writeFileSync error=", e);
        return;
    }
    try {
        var wb = XLSX.readFileSync(fname);
        console.log("wb initial=", JSON.stringify(wb));
    } catch (e) {
        res.sendStatus(500);
        console.log("Impossible to create xlsx file! Reason:" + e);
        return;
    }
    var worksheetColumns = [];

    wb.SheetNames = [];
    wb.SheetNames.push('Sheet1');
    wb.Sheets['Sheet1'] = {
        //'!ref': 'A1:',
        '!cols': worksheetColumns
    };
    for (var j = 0; j < columns.length; j++) {
        worksheetColumns.push({wpx: columns[j].width});
        var currentHeader = XLSX.utils.encode_cell({c: j, r: 0});
        wb.Sheets['Sheet1'][currentHeader] = {t: "s", v: columns[j].name, s: {font: {bold: true}}};
    }

    var lineNum = 1;

    fillTableData(0, rows);

    function fillTableData(index, rows) {
        console.log("fillTableData index, rows=", index, rows);
        if (!rows[index]) return;
        var rowData = rows[index];
        console.log("rowData=", rowData);
        var lastCellInRaw;
        for (var i = 0; i < columns.length; i++) {

            var headerName=wb.Sheets['Sheet1'][XLSX.utils.encode_cell({c: i, r: 0})].v;
            var columnDataID;
            for(var k in columns){
                var columnUnit=columns[k];
                if(columnUnit.name==headerName){
                    columnDataID = columnUnit.data;  console.log("columnUnit.name, columnUnit.data=",columnUnit.name, columnUnit.data);
                    break;
                }
            }
            var displayValue = rowData[columnDataID];
            var currentCell = XLSX.utils.encode_cell({c: i, r: lineNum});
            lastCellInRaw=currentCell;
            wb.Sheets['Sheet1'][currentCell] = {
                t: "s",
                v: displayValue
                , s: {
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
        wb.Sheets['Sheet1']['!ref']='A1:'+lastCellInRaw;  console.log("!ref=",wb.Sheets['Sheet1']['!ref']);
        fillTableData(index + 1, rows);
    }

    XLSX.writeFileAsync(fname, wb, {bookType: "xlsx", cellStyles: true}, function (err, result) {
        if (err)console.log("err=", err);
        var options = {headers: {'Content-Disposition': 'attachment; filename =out.xlsx'}};
        res.sendFile(fname, options, function (err) {
            if (err) {
                res.sendStatus(500);
                console.log("send xlsx file err=", err);
                // return;
            }
            // fs.unlinkSync(fname);
        })

    });
});


app.listen(port, function (err) {
    console.log("app run on port ", port);
});


function getUIDNumber() {
    var str = uid.time();
    var len = str.length;
    var num = 0;
    for (var i = (len - 1); i >= 0; i--) {
        num += Math.pow(256, i) * str.charCodeAt(i);
    }
    return num;
}