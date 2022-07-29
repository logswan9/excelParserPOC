
const ExcelJS = require('exceljs');
var submitBtn = document.getElementById('submitBtn').addEventListener("click", uploadFiles);

async function uploadFiles() {
    var files = document.getElementById('file_upload').files[0].path;
    // if(files.length==0){
    //     alert("Please first choose or drop any file(s)...");
    //     return;
    // }
    var filename1="";
    // for(var i=0;i<files.length;i++){
        
    // }
    filename1+=files;
    //alert("Selected file(s) :\n____________________\n"+filename1);


    

    const wb = new ExcelJS.Workbook();
    //console.log(fileName1);
    const fileName = filename1;

    wb.xlsx.readFile(fileName).then(() => {
        
        const ws = wb.getWorksheet('Sheet1');

        const c1 = ws.getColumn(1);
        
        c1.eachCell(c => {

            console.log(c.value);
        });

        const c2 = ws.getColumn(2);
        
        c2.eachCell(c => {

            console.log(c.value);
        });
    }).catch(err => {
        console.log(err.message);
    });


}
