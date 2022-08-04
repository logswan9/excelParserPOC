
const ExcelJS = require('exceljs');
const homeDir = require('os').homedir(); 
const desktopDir = `${homeDir}\\Desktop\\`;
const electron = require('electron');
const ipcRen = electron.ipcRenderer;
const fs = require('fs');


var filePath = "";
var fileName = "";
var submitBtn = document.getElementById('submitBtn').addEventListener("click", getFilePath);
var month = "";
var year = "";
var adjReason = "";

async function getFilePath() {

    filePath = document.getElementById('file_upload').files[0].path;
    fileName = document.getElementById('file_upload').files[0].name;
    
    ipcRen.send('get-input');

}


ipcRen.on('input-window-close', function(e, year, month, adjReason) {
    // FIRST NETWORK COPY
    fs.copyFile(filePath, desktopDir + 'OUT\\Retroterm ' + month + " " + year + ".xlsx", (err) => {
        if (err) console.log(err);
    });

    // SECOND NETWORK COPY
    // fs.copyFile(filePath, desktopDir + 'OUT\\' + fileName, (err) => {
    //     if (err) console.log(err);
    // });
}); 



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


// const homeDir = require('os').homedir(); 
// const desktopDir = `${homeDir}\\Desktop\\`;
// console.log(desktopDir);

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








// const wb = new ExcelJS.Workbook();
// //console.log(fileName1);
// const fileName = file;

// wb.xlsx.readFile(fileName).then(() => {
    
//     const ws = wb.getWorksheet(2);

//     const c1 = ws.getColumn(1);
    
//     c1.eachCell(c => {

//         console.log(c.value);
//     });

//     const c2 = ws.getColumn(2);
    
//     c2.eachCell(c => {

//         console.log(c.value);
//     });
// }).catch(err => {
//     console.log(err.message);
// });
