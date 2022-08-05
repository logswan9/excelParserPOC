
const ExcelJS = require('exceljs');
const homeDir = require('os').homedir(); 
const desktopDir = `${homeDir}\\Desktop\\`;
const electron = require('electron');
const ipcRen = electron.ipcRenderer;
const fs = require('fs');
const { stringify } = require('querystring');


var filePath = "";
var fileName = "";
var submitBtn = document.getElementById('submitBtn').addEventListener("click", getFilePath);
var month = "";
var year = "";
var adjReason = "";
var filteredData = [];
var columnHeaders = [];
var fallonClaims = [];
var riClaims = [];
var orthoClaims = []

async function getFilePath() {

    filePath = document.getElementById('file_upload').files[0].path;
    fileName = document.getElementById('file_upload').files[0].name;
    
    ipcRen.send('get-input');

}


ipcRen.on('input-window-close', function(e, year, month, adjReason) {
    // FIRST NETWORK COPY
    // fs.copyFile(filePath, desktopDir + 'OUT\\Retroterm ' + month + " " + year + ".xlsx", (err) => {
    //     if (err) console.log(err);
    // });

    // fs.copyFile(filePath, "\\\\NAS05059PN\\Shared_MN020\\Dental\\SB Support\\SKYNET\\Node Testing\\Test.xlsx", (err) => {
    //     if (err) console.log(err);
    // });

    // SECOND NETWORK COPY
    // fs.copyFile(filePath, desktopDir + 'OUT\\' + fileName, (err) => {
    //     if (err) console.log(err);
    // });


    // START EXCEL LOGIC

    
    const wb = new ExcelJS.Workbook();
    //console.log(fileName1);
    //const fileName = file;

    wb.xlsx.readFile(filePath).then(() => {
        
        const ws = wb.getWorksheet(1);

        // const c1 = ws.getColumn(1);
        
        // c1.eachCell(c => {

        //     console.log(c.value);
        // });

        // const c2 = ws.getColumn(3);
        
        // c2.eachCell(c => {

        //     console.log(c.value);
        // });

        var rowCount = ws.rowCount;
        //var output = "";

        var count = 0;
        
        ws.getRow(1).eachCell(function(cell) {
            columnHeaders.push(cell.text);
        });
        
        // var yum = ws.getColumn('A');
        // yum.eachCell(function(cell, rowNumber) {
        //     console.log(cell.text);
        // });
        ws.eachRow({includeEmpty: true}, function(row, rowNumber) {

            // row.eachCell(function(cell) {
            //     console.log(cell.);
            // });
            
            
            if (rowNumber > 2) {
                if (row.getCell(42).text === 'N' && row.getCell(43).text === '') {
                    //var compactArray = Object.values(row.values);
                    // console.log(rowNumber);
                    // console.log(row.values);
                    var rowData = {};
                    
                    for (var i = 1; i <= row.values.length; i++) {
                        if ((String(row.values[i])) === 'undefined') {
                            rowData[columnHeaders[i - 1]] = '';
                        } else {
                            rowData[columnHeaders[i - 1]] = (String(row.values[i]));
                        }
                        
                        // if (String(row.values[i]).substring(0, 2).includes('D8')) {
                        //     console.log(row.values);
                        //     console.log(rowNumber);
                        // }
                    }
                    
                    //count++;
                    filteredData.push(rowData);
                }
                //console.log();
            }
            //console.log(JSON.stringify(rowData));
            
            
            // console.log(rowNumber);
            // var compactArray = Object.values(row.values);
            // console.log(compactArray);


            //console.log("Row: " + rowNumber + "  Values: " + row.values);
            
        });
        //console.log(filteredData.length);
        // filteredData.forEach(row => {
        //     console.log(JSON.stringify(row));
        // });
        // console.log(count);
        seperateData();

    }).catch(err => {
        console.log(err.message);
    });

    





    // const workbook = new ExcelJS.Workbook();
    // const worksheet = workbook.addWorksheet("My Sheet");

    // worksheet.columns = [
    // {header: 'Id', key: 'id', width: 10},
    // {header: 'Name', key: 'name', width: 32}, 
    // {header: 'D.O.B.', key: 'dob', width: 15,}
    // ];

    // worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
    // worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});


    // const fs = require('fs');

    // //const folderName = '/Users/joe/test';

    // try {
    // if (!fs.existsSync(desktopDir + 'OUT')) {
    //     fs.mkdirSync(desktopDir + 'OUT');
    // }
    // } catch (err) {
    // console.error(err);
    // }

    // // save under export.xlsx
    // await workbook.xlsx.writeFile(desktopDir + 'OUT\\export.xlsx');




    
}); 

function seperateData() {
    console.log(filteredData.length);
    var count = 0;
    filteredData.forEach((row, index) => {
        if (row['PAGR_NAME'] === 'Fallon Community Health Plan') {
            filteredData.splice(row, 1);
            
        }


        if (row['GRGR_STATE'] === 'RI') {
            riClaims.push(row);
            filteredData.splice(row, 1);
            
        }

        if (row['DPDP_ID'].substring(0, 2).includes('D8')) {
            //console.log('row to be added: ' + JSON.stringify(row));
            //console.log('index on push: ' + index);
            orthoClaims.push(row);
    
            //console.log('index after push: ' + index);
            //filteredData.splice(row, 1);
            count++;
            //console.log('index after delete: ' + index);
            
            //console.log('row to be deleted: ' + JSON.stringify(row));
        }
    });

    console.log(count);
    console.log('ORTHO: ' + orthoClaims.length);
    console.log(filteredData.length);
}



//console.log(fileName);
    
// if(files.length==0){
//     alert("Please first choose or drop any file(s)...");
//     return;
// }
//var filename1 = "";
// for(var i=0;i<files.length;i++){
    
// }
//filename1 = file;
//alert("Selected file(s) :\n____________________\n"+filename1);


