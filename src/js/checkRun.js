import { Base } from './Base.js';

export function loadCheck() {
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
    var forcastedPaidDate;
    var numberProcessors;
    //var retrotermBtn = document.getElementById('retroterm');

    var started = false;
    //var submitBtn = document.getElementById('keanClean').addEventListener("click", startKeanClean); // Set listener on Keen Button
    document.getElementById("selectionDiv").style.display = "none";
    document.getElementById("wrapper2CheckRun").style.display = "block";
    document.getElementById("title2").style.display = "block";
    document.getElementById("title2").innerHTML = "Check Run";

    var cancelButton = document.getElementById("cancelCheck");
    var submitButton = document.getElementById("submitCheck");

    var monthFN = "";
    var yearFN = "";
    var adjReasonG = "";
    var projectNoteG = "";
    var rawData = [];
    var tomorrowsDate;
    var events = false;
    var columnHeaders = [];


    if (!events) {
        cancelButton.addEventListener('click', function () {
            //document.getElementById("wrapper2").style.display = "none";
            location.reload();
        });
    
        submitButton.addEventListener('click', function () {
            if (!started) {
                document.getElementById("noFileMessageCheck").textContent = "";

                try {
                    filePath = document.getElementById('file_upload').files[0].path;
                    fileName = document.getElementById('file_upload').files[0].name;
                } catch (error) {
                    setTimeout(function () {
                        document.getElementById("noFileMessageCheck").textContent = "No File selected! Please select a file.";
                    }, 1000);
                    
                }
                


                if (filePath !== "") {
                    if (fileName.endsWith(".xlsx")) {
                        beginCheckRun();
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

    async function beginCheckRun() {
        // tomorrowsDate = new Date(Date.now());
        // tomorrowsDate.setDate(tomorrowsDate.getDate() + 1)
        // tomorrowsDate = tomorrowsDate.toLocaleDateString();
        var monthFPD;
        var dayFPD;
        var currentYear = new Date;
        forcastedPaidDate = new Date(Date.parse(document.getElementById("checkRunDate").value));
        monthFPD = forcastedPaidDate.getMonth() + 1;
        dayFPD = forcastedPaidDate.getDate();
        forcastedPaidDate = new Date(+currentYear.getFullYear(), monthFPD - 1, +dayFPD + 1);
        forcastedPaidDate = String(Date.parse(forcastedPaidDate));
        numberProcessors = document.getElementById("checkRunPNum").value;
        
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
        const formatter = new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD',
          });
        
        workbookReader.on('worksheet', worksheet => {
            
            worksheet.on('row', row => {
                
                rowCellData = {};
                
                row.eachCell(function(cell) {
                    if (cell.text.trim() == "ForcastedPaidDate") {
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

                    if (rowCellData['ExplCode'] != 'PNDUC') {
                        if (rowCellData['Dscrption'] != 'CARRIER/LOCALITY UNDETERMINED ON DOS' && rowCellData['Dscrption'] != 'COMPLIANCE PAY DATE NOT CALCULATED - MATCHING COMPLIANCE RULE NOT FOUND.') {
                            //console.log(String(Date.parse(base.getDateFromExcel(rowCellData['ForcastedPaidDate']))));
                            if (String(Date.parse(base.getDateFromExcel(rowCellData['ForcastedPaidDate']))) == forcastedPaidDate) {
                                rowCellData['ForcastedPaidDate'] = base.getDateFromExcel(rowCellData['ForcastedPaidDate']);
                                rowCellData['Compliance Date'] = base.getDateFromExcel(rowCellData['Compliance Date']);
                                rowCellData['March Payable Date'] = base.getDateFromExcel(rowCellData['March Payable Date']);
                                rowCellData['DateRcvd'] = base.getDateFromExcel(rowCellData['DateRcvd']);
                                if (rowCellData['MOI'] == '0') {
                                    rowCellData['MOI'] = 'False';
                                } else {
                                    rowCellData['MOI'] = 'True';
                                }
                                rowCellData['Charge'] = parseFloat(rowCellData['Charge']);
                                rowCellData['Charge'] = String(formatter.format(rowCellData['Charge']));
                                rowCellData['FromDOS'] = base.getDateFromExcel(rowCellData['FromDOS']);
                                rowCellData['Daily  Aging'] = parseFloat(rowCellData['Daily  Aging']);
                                rowCellData['Daily  Aging'] = String(Math.round(rowCellData['Daily  Aging']));
                                rawData.push(rowCellData);
                                rowCount++;
                            }
                            
                        }
                        
                    }
                    //console.log(rowCellData);
                    
                }

                if (atHeaderCheck) {
                    atHeaderCheck = false;
                    pastHeaderCheck = true; // Past all headers in row
                }
                
                //rowCount++;
            });
        });
        
        
        workbookReader.on('end', () => {
            console.log(rawData);
            console.log(forcastedPaidDate);
            createFile();
            //console.log(columnHeaders);
            //seperateData();
        });
        workbookReader.on('error', (err) => {
            console.log(err);
        });

        
        
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
            filename: desktopDir + 'OUT\\Check Run.xlsx',
            useStyles: true,
            useSharedStrings: true
        };

        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
        const rawDataWS = workbook.addWorksheet("RawData");
        
        var checkNumbPro = 1;
        var excelColumns = [];




        columnHeaders.forEach(value => {
            excelColumns.push({header: value, key: value, width: 15});
        });

        
        
        //console.log(excelColumns);
        rawDataWS.columns = excelColumns;

        // while (checkNumbPro <= numberProcessors) {
        //     var 
        // }
        

        console.log("STARTING RAW DATA INSERT!");
        for (let i = 0; i < rawData.length; i++) {
            rawDataWS.addRow(rawData[i]).commit();
        }

        rawDataWS.commit();
        //console.log(Math.trunc(rowCount / numberProcessors));

        var dataIndex = 0;
        var previousSpotInData = 0;
        var atEnd = false;
        while (checkNumbPro <= numberProcessors) {
            if (checkNumbPro == numberProcessors) {
                atEnd = true;
            }
            var processorWS = workbook.addWorksheet("Processor " + checkNumbPro);
            processorWS.columns = excelColumns;
            if (75 * numberProcessors > rowCount) {
                for (var i = previousSpotInData; i < rawData.length; i++) {
                    processorWS.addRow(rawData[i]).commit();
                    previousSpotInData ++;
                    dataIndex++;
                    if (dataIndex == Math.trunc(rowCount / numberProcessors)) {
                        processorWS.commit();
                        dataIndex = 0;
                        break;
                    }
                }
            } else {
                for (var i = previousSpotInData; i < rawData.length; i++) {
                    processorWS.addRow(rawData[i]).commit();
                    previousSpotInData ++;
                    dataIndex++;
                    if (dataIndex == 75) {
                        processorWS.commit();
                        dataIndex = 0;
                        break;
                    }
                }
            }
            
            checkNumbPro++;
            if (atEnd) {
                var extraWS = workbook.addWorksheet("Extra");
                extraWS.columns = excelColumns;
                for (var i = previousSpotInData; i < rawData.length; i++) {
                    extraWS.addRow(rawData[i]).commit();
                    previousSpotInData ++;
                }
                extraWS.commit();
            }
        }
        
        
        
        await workbook.commit();
        console.log("COMPLETE!");

        
        // for (let i = 0; i < mcaFINAL.length; i++) {
        //     mcaFinalWS.addRow(mcaFINAL[i]).commit();
        // }

        // mcaFinalWS.commit();
        // await workbookMCA.commit();
        //console.log(mcaFINALString);


        rawData = [];
        
        loadingIcon(false);

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
