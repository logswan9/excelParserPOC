import { Base } from './Base.js';

export function loadEmployeeChanges() {
    const ExcelJS = require('exceljs');
    const homeDir = require('os').homedir(); 
    const desktopDir = `${homeDir}\\Desktop\\`;
    const electron = require('electron');
    const ipcRen = electron.ipcRenderer;
    const fs = require('fs');
    const { stringify } = require('querystring');
    module.exports = Base;
    const base = new Base;
    

    //const base = new Base();
    
    var filePath = "";
    var fileName = "";
    var comparisonFilePath = "";
    //var retrotermBtn = document.getElementById('retroterm');

    var started = false;
    
    document.getElementById("selectionDiv").style.display = "none";
    document.getElementById("wrapper2EmployeeChanges").style.display = "block";
    document.getElementById("title2").style.display = "block";
    document.getElementById("title2").innerHTML = "Employee Changes";

    var cancelButton = document.getElementById("cancelEmployeeChanges");
    var submitButton = document.getElementById("submitEmployeeChanges");

    
    var rawData = [];
    var events = false;
    var columnHeaders = [];
    var rawFirstFileData = [];
    var rawSecondFileData = [];
    var changedEntries = [];
    var workbook;
    var compareWS;


    if (!events) {
        cancelButton.addEventListener('click', function () {
            //document.getElementById("wrapper2").style.display = "none";
            location.reload();
        });
    
        submitButton.addEventListener('click', function () {
            if (!started) {
                document.getElementById("noFileMessageEmployeeChanges").textContent = "";

                try {
                    filePath = document.getElementById('file_upload').files[0].path;
                    fileName = document.getElementById('file_upload').files[0].name;
                    comparisonFilePath = document.getElementById('file_upload2Emp').files[0].path;
                } catch (error) {
                    setTimeout(function () {
                        document.getElementById("noFileMessageEmployeeChanges").textContent = "No File selected! Please select a file.";
                    }, 1000);
                    
                }
                


                if (filePath !== "") {
                    if (fileName.endsWith(".xlsx")) {
                        beginEmployeeChanges();
                        started = true;
                        submitButton.style.display = "none";
                        
                        
                    } else {
                        setTimeout(function () {
                            document.getElementById("noFileMessageCheck").textContent = "ERROR! You must select a xlsx file.";
                        }, 1000);
                    }
                    
                } 
                
            }
        });
    }

    events = true;
    var rowCount = 0;

    async function beginEmployeeChanges() {
        // tomorrowsDate = new Date(Date.now());
        // tomorrowsDate.setDate(tomorrowsDate.getDate() + 1)
        // tomorrowsDate = tomorrowsDate.toLocaleDateString();
        

        
        
        loadingIcon(true);
        //document.getElementById("wrapper2").style.display = "none";
        

        const options = {
            sharedStrings: 'cache',
            hyperlinks: 'emit',
            worksheets: 'emit',
            entries: 'emit'
        };
        const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, options);
        workbookReader.read();
        
        var wsCreationCount = 0;
        
        var rowNumber = 0;
        var atHeaderCheck = false;
        var pastHeaderCheck = false;
        var rowCellData = {};
        var rowWithKey = {};
        
        
        workbookReader.on('worksheet', worksheet => {
            
            worksheet.on('row', row => {
                
                rowCellData = {};
                rowWithKey = {};
                
                
                row.eachCell(function(cell) {
                    if (cell.text.trim() == "Employee Name") {
                        atHeaderCheck = true;
                    }
                    if (atHeaderCheck) {
                        columnHeaders.push(cell.text.trim());
                    }
                    
                });
                
                
                if (pastHeaderCheck) { //Main logic loop for parsing data
                    //TODO
                    row.eachCell({includeEmpty: true}, function(cell, cellIndex) { // Iterate over each cell, include empty cells
                        rowCellData[columnHeaders[(cellIndex - 1)]] = String(cell.text.trim());
                    });
                    //rowWithKey[rowCellData['Employee ID']] = rowCellData;
                    rawFirstFileData.push(rowCellData);
                    
                }

                if (atHeaderCheck) {
                    atHeaderCheck = false;
                    pastHeaderCheck = true; // Past all headers in row
                }
                
                //rowCount++;
            });
        });
        
        
        workbookReader.on('end', () => {
            console.log(rawFirstFileData);
            beginParseSecondFile();
            //createFile();
            //console.log(columnHeaders);
            //seperateData();
        });
        workbookReader.on('error', (err) => {
            console.log(err);
        });

        
        
    } 

    function beginParseSecondFile() {
        const options = {
            sharedStrings: 'cache',
            hyperlinks: 'emit',
            worksheets: 'emit',
            entries: 'emit'
        };
        const workbookReader2 = new ExcelJS.stream.xlsx.WorkbookReader(comparisonFilePath, options);
        workbookReader2.read();

        var atHeaderCheck = false;
        var pastHeaderCheck = false;
        var rowCellData = {};
        var rowWithKey = {};
        
        
        workbookReader2.on('worksheet', worksheet => {
            
            worksheet.on('row', row => {
                
                rowCellData = {};
                rowWithKey = {};
                
                
                row.eachCell(function(cell) {
                    if (cell.text.trim() == "Employee Name") {
                        atHeaderCheck = true;
                    }
                    // if (atHeaderCheck) {
                    //     columnHeaders.push(cell.text.trim());
                    // }
                    
                });
                
                
                if (pastHeaderCheck) { //Main logic loop for parsing data
                    //TODO
                    row.eachCell({includeEmpty: true}, function(cell, cellIndex) { // Iterate over each cell, include empty cells
                        rowCellData[columnHeaders[(cellIndex - 1)]] = String(cell.text.trim());
                    });
                    //rowWithKey[rowCellData['Employee ID']] = rowCellData;
                    rawSecondFileData.push(rowCellData);
                    
                }

                if (atHeaderCheck) {
                    atHeaderCheck = false;
                    pastHeaderCheck = true; // Past all headers in row
                }
                
                //rowCount++;
            });
        });
        
        
        workbookReader2.on('end', () => {
            console.log(rawSecondFileData);
            compareForChanges();
            //createFile();
            //console.log(columnHeaders);
            //seperateData();
        });
        workbookReader2.on('error', (err) => {
            console.log(err);
        });

    }

    function compareForChanges() {
        createFile();
        var newMemberCheck = false;
        var currentColumnChanges = [];
        for (var i = 0; i < rawSecondFileData.length; i++) {
            for (var t = 0; t < rawFirstFileData.length; t ++) {
                if (rawSecondFileData[i]['Employee ID'] === rawFirstFileData[t]['Employee ID']) {
                    newMemberCheck = false;
                    for (var z = 0; z < columnHeaders.length; z++) {
                        if (rawSecondFileData[i][columnHeaders[z]] !== rawFirstFileData[t][columnHeaders[z]]) {
                            currentColumnChanges.push(columnHeaders[z]);
                            console.log(rawFirstFileData[t][columnHeaders[z]] + '  CHANGED TO:  ' + rawSecondFileData[i][columnHeaders[z]]);
                            //let testRow = compareWS.addRow(rawSecondFileData[i]);

                            

                        } 
                        
                        
                    }
                    break;
                } else {
                    newMemberCheck = true;
                }
            }
            if (newMemberCheck) {
                //console.log('NEW MEMBER ADDED:');
                //console.table(rawSecondFileData[i]);
                newMemberCheck = false;
            }
            console.log(currentColumnChanges);
            let testRow = compareWS.addRow(rawSecondFileData[i]);

            testRow.eachCell(function(cell) {
                //console.log(cell);
                //console.log(cell._column._key);
                for (var p = 0; p < currentColumnChanges.length; p++) {
                    console.log(currentColumnChanges[p]);
                    if (cell._column._key === currentColumnChanges[p]) {
                        cell.font = {
                            color: { argb: "f00a0a" },
                            bold: true,
                            };
                    }
                }
                
                
            });

            testRow.commit();
            currentColumnChanges = [];
        }


        compareWS.commit();

        workbook.commit();

        console.log("DONE");
        //ACCOUNT FOR NEW MEMBER ADDITIONs
    }

    async function createFile() {
        const fs = require('fs');

        try {
            if (!fs.existsSync(desktopDir + 'OUT')) {
                fs.mkdirSync(desktopDir + 'OUT');
            }
            } catch (err) {
            console.error(err);
            }

        const options = {
            filename: desktopDir + 'OUT\\Employee Comparison.xlsx',
            useStyles: true,
            useSharedStrings: true
        };

        workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
        compareWS = workbook.addWorksheet("RawData");
        
        
        var excelColumns = [];




        columnHeaders.forEach(value => {
            excelColumns.push({header: value, key: value, width: 15});
        });

        
        
        //console.log(excelColumns);
        compareWS.columns = excelColumns;

        // while (checkNumbPro <= numberProcessors) {
        //     var 
        // }
        

        // console.log("STARTING RAW DATA INSERT!");
        // for (let i = 0; i < rawData.length; i++) {
        //     compareWS.addRow(rawData[i]).commit();
        // }

        // compareWS.commit();
        // //console.log(Math.trunc(rowCount / numberProcessors));

        // var dataIndex = 0;
        // var previousSpotInData = 0;
        // var atEnd = false;
        // while (checkNumbPro <= numberProcessors) {
        //     if (checkNumbPro == numberProcessors) {
        //         atEnd = true;
        //     }
        //     var processorWS = workbook.addWorksheet("Processor " + checkNumbPro);
        //     processorWS.columns = excelColumns;
        //     if (75 * numberProcessors > rowCount) {
        //         for (var i = previousSpotInData; i < rawData.length; i++) {
        //             processorWS.addRow(rawData[i]).commit();
        //             previousSpotInData ++;
        //             dataIndex++;
        //             if (dataIndex == Math.trunc(rowCount / numberProcessors)) {
        //                 processorWS.commit();
        //                 dataIndex = 0;
        //                 break;
        //             }
        //         }
        //     } else {
        //         for (var i = previousSpotInData; i < rawData.length; i++) {
        //             processorWS.addRow(rawData[i]).commit();
        //             previousSpotInData ++;
        //             dataIndex++;
        //             if (dataIndex == 75) {
        //                 processorWS.commit();
        //                 dataIndex = 0;
        //                 break;
        //             }
        //         }
        //     }
            
        //     checkNumbPro++;
        //     if (atEnd) {
        //         var extraWS = workbook.addWorksheet("Extra");
        //         extraWS.columns = excelColumns;
        //         for (var i = previousSpotInData; i < rawData.length; i++) {
        //             extraWS.addRow(rawData[i]).commit();
        //             previousSpotInData ++;
        //         }
        //         extraWS.commit();
        //     }
        // }
        
        
        
        // await workbook.commit();
        // console.log("COMPLETE!");

        
        // for (let i = 0; i < mcaFINAL.length; i++) {
        //     mcaFinalWS.addRow(mcaFINAL[i]).commit();
        // }

        // mcaFinalWS.commit();
        // await workbookMCA.commit();
        //console.log(mcaFINALString);


        rawData = [];
        
        //loadingIcon(false);

    }

    async function loadingIcon(loading) {
        if (loading) {
            document.getElementById("loader").style.display = "block";
            document.getElementById("loadingMessage").style.display = "block"
            document.getElementById("consoleMessage").style.display = "block"
        } else {
            document.getElementById("loader").style.display = "none";
            document.getElementById("consoleMessage").innerHTML = "";
            document.getElementById("loadingMessage").innerHTML = "COMPLETE. File Has been saved in the OUT folder on your Desktop.";
            setTimeout(function () {
                document.getElementById("loadingMessage").innerHTML = "";
                location.reload();
            }, 10000);
        }
    }


}
