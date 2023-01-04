export function loadRetro() {
    const ExcelJS = require('exceljs');
    const homeDir = require('os').homedir(); 
    const desktopDir = `${homeDir}\\Desktop\\`;
    const electron = require('electron');
    const ipcRen = electron.ipcRenderer;
    const fs = require('fs');
    const { stringify } = require('querystring');


    var filePath = "";
    var fileName = "";
    //var retrotermBtn = document.getElementById('retroterm');

    var started = false;
    //var submitBtn = document.getElementById('keanClean').addEventListener("click", startKeanClean); // Set listener on Keen Button
    document.getElementById("selectionDiv").style.display = "none";
    document.getElementById("wrapper2").style.display = "block";
    document.getElementById("title2").style.display = "block";
    document.getElementById("title2").innerHTML = "Retroterm/MCA";

    var cancelButton = document.getElementById("cancelRetro");
    var submitButton = document.getElementById("submitRetro");

    var monthFN = "";
    var yearFN = "";
    var adjReasonG = "";
    var projectNoteG = "";
    var rawData = [];
    var filteredData = [];
    var columnHeaders = [];
    //var fallonClaims = [];
    var riClaims = [];
    var orthoClaims = []
    //var adjustedOrtho = [];
    var nonAdjusted = [];
    var mca = [];
    var mcaFINALString = "";
    var tomorrowsDate;
    var events = false;
    var stateList = {'AK':'12','AL':'12','AR':'18','AZ':'12','CA':'12','CA':'12','CO':'12','CT':'12','DC':'6','DE':'12','FL':'12',
                'GA':'18','HI':'12','IA':'12','ID':'12','IL':'12','IN':'24','KS':'12','KY':'24','LA':'12','MA':'12','MD':'6',
                'ME':'12','MI':'12','MN':'12','MO':'12','MS':'12','MT':'24','NC':'12','ND':'12','NE':'12','NH':'18','NJ':'18',
                'NM':'12','NV':'12','NY':'24','OH':'24','OK':'24','OR':'24','PA':'12','PR':'12','RI':'18','SC':'12','SD':'12',
                'TN':'18','TX':'6','UT':'12','VA':'12','VI':'12','VT':'12','WA':'24','WI':'12','WV':'12','WY':'12'};


    if (!events) {
        cancelButton.addEventListener('click', function () {
            //document.getElementById("wrapper2").style.display = "none";
            location.reload();
        });
    
        submitButton.addEventListener('click', function () {
            if (!started) {
                document.getElementById("noFileMessageRetro").textContent = "";

                try {
                    filePath = document.getElementById('file_upload').files[0].path;
                    fileName = document.getElementById('file_upload').files[0].name;
                } catch (error) {
                    setTimeout(function () {
                        document.getElementById("noFileMessageRetro").textContent = "No File selected! Please select a file.";
                    }, 1000);
                    
                }
                


                if (filePath !== "") {
                    if (fileName.endsWith(".xlsx")) {
                        filterInitialDataset(document.getElementById('form1').value, document.getElementById('form2').value, document.getElementById('form3').value, document.getElementById('form4').value);
                        started = true;
                        submitButton.style.display = "none";
                        
                        
                    } else {
                        setTimeout(function () {
                            document.getElementById("noFileMessageRetro").textContent = "ERROR! You must select a xlsx file.";
                        }, 1000);
                    }
                    
                } 
                
            }
        });
    }

    events = true;

    // retrotermBtn.addEventListener("click", function () {
    //     tomorrowsDate = new Date(Date.now());
    //     tomorrowsDate.setDate(tomorrowsDate.getDate() + 1)
    //     tomorrowsDate = tomorrowsDate.toLocaleDateString();
    //     console.log(tomorrowsDate);

    //     filePath = document.getElementById('file_upload').files[0].path;
    //     fileName = document.getElementById('file_upload').files[0].name;

    //     document.getElementById("wrapper2").style.display = "block";
        
    //     var cancelButton = document.getElementById("cancel");
    //     var submitButton = document.getElementById("submit");
        
    //     if (!events) {
    //         cancelButton.addEventListener('click', function () {
    //             //document.getElementById("wrapper2").style.display = "none";
    //             location.reload();
    //         });
        
    //         submitButton.addEventListener('click', function () {
    //             filterInitialDataset(document.getElementById('form1').value, document.getElementById('form2').value, document.getElementById('form3').value, document.getElementById('form4').value);
    //         });
    //     }

    //     events = true;
        
    //     //ipcRen.send('get-input'); for opening seperate input window

    // });


    async function filterInitialDataset(month, year, adjReason, note) {
        tomorrowsDate = new Date(Date.now());
        tomorrowsDate.setDate(tomorrowsDate.getDate() + 1)
        tomorrowsDate = tomorrowsDate.toLocaleDateString();
        console.log(tomorrowsDate);
        
        loadingIcon(true);
        //document.getElementById("wrapper2").style.display = "none";
        monthFN = month;
        yearFN = year;
        adjReasonG = adjReason;
        projectNoteG = note;
        console.log("month: " + monthFN + " year: " + yearFN + " adjReason: " + adjReasonG + " note: " + projectNoteG);

        const options = {
            sharedStrings: 'cache',
            hyperlinks: 'emit',
            worksheets: 'emit',
            entries: 'emit'
        };
        const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, options);
        workbookReader.read();
        
        var rowCount = 0;
        var rowNumber = 0;
        var rowCellData = [];
        
        workbookReader.on('worksheet', worksheet => {
            
            worksheet.on('row', row => {
                
                rowCellData = [];
                if (rowNumber == 0) {
                    row.eachCell(function(cell) {
                        columnHeaders.push(cell.text);
                    });
                }

                
                if (rowNumber > 1) {
                    if (row.getCell(42).text === 'N' && row.getCell(43).text === '') {
                        var rowData = {};
                        
                        for (var i = 1; i <= row.values.length; i++) {
                            if ((String(row.values[i])) === 'undefined') {
                                rowData[columnHeaders[i - 1]] = '';
                            } else {
                                rowData[columnHeaders[i - 1]] = (String(row.values[i]));
                            }
                            
                        }
                        
                        filteredData.push(rowData);
                        rawData.push(rowData);
                    }

                }
                rowNumber++;
                rowCount++;
            });
        });
        
        
        workbookReader.on('end', () => {
            seperateData();
        });
        workbookReader.on('error', (err) => {
            console.log(err);
        });

        
        
    } 

    function seperateData() {
        console.log("Start of Seperate Data: " + filteredData.length);
        var count = 0;
        filteredData.forEach((row, index) => {
            // if (row['PAGR_NAME'] === 'Fallon Community Health Plan') {   INITIIALLY FOR REMOVING FALLON CLAIMS
            //     filteredData.splice(row, 1);
                
            // }
            var newRow = {};

            if (row['GRGR_STATE'] === 'RI' && row['PAGR_NAME'] != 'Fallon Community Health Plan') {
                riClaims.push({"CLCL_ID":row.CLCL_ID});
                count++;
            }

            if (row['DPDP_ID'].substring(0, 2).includes('D8') && row['GRGR_STATE'] != 'RI' && !row['PAGR_NAME'] != 'Fallon Community Health Plan') {
        
                columnHeaders.forEach(value => {
                    if (value != "ACTIVE_FLAG" && value != "J11_FLAG") {
                        newRow[value] = row[value];
                    } 
                });

                newRow['RCVRY_NUM'] = stateList[row['GRGR_STATE']];
                    
                    //console.log(value);
                orthoClaims.push(newRow);
                
            }
        });
        

        filteredData.forEach((row, index) => {
            var newRow = {};
            
            if (row['CLCL_ID'].substring(row['CLCL_ID'].length - 2, row['CLCL_ID'].length + 1) === '00' && !row['DPDP_ID'].substring(0, 2).includes('D8') && row['GRGR_STATE'] != 'RI' && row['PAGR_NAME'] != 'Fallon Community Health Plan') {
                columnHeaders.forEach(value => {
                    if (value != "ACTIVE_FLAG" && value != "J11_FLAG") {
                        newRow[value] = row[value];
                    } 
                });

                newRow['RCVRY_NUM'] = stateList[row['GRGR_STATE']];
                nonAdjusted.push(newRow);
                mca.push({"CLCL_ID":row.CLCL_ID});
                
                
            } 
        });
        
        removeDupes();
    }

    function removeDupes() {

        riClaims = Array.from(new Set(riClaims.map(a => a.CLCL_ID)))
        .map(CLCL_ID => {
        return riClaims.find(a => a.CLCL_ID === CLCL_ID)
        })

        orthoClaims = Array.from(new Set(orthoClaims.map(a => a.CLCL_ID)))
        .map(CLCL_ID => {
        return orthoClaims.find(a => a.CLCL_ID === CLCL_ID)
        })

        // adjustedOrtho = Array.from(new Set(adjustedOrtho.map(a => a.CLCL_ID)))
        // .map(CLCL_ID => {
        // return adjustedOrtho.find(a => a.CLCL_ID === CLCL_ID)
        // })

        nonAdjusted = Array.from(new Set(nonAdjusted.map(a => a.CLCL_ID)))
        .map(CLCL_ID => {
        return nonAdjusted.find(a => a.CLCL_ID === CLCL_ID)
        })

        // filteredData = Array.from(new Set(filteredData.map(a => a.CLCL_ID)))
        // .map(CLCL_ID => {
        // return filteredData.find(a => a.CLCL_ID === CLCL_ID)
        // })

        mca = Array.from(new Set(mca.map(a => a.CLCL_ID)))
        .map(CLCL_ID => {
        return mca.find(a => a.CLCL_ID === CLCL_ID)
        })

        // mcaFINAL = Array.from(new Set(mcaFINAL.map(a => a.CLCL_ID)))
        // .map(CLCL_ID => {
        // return mcaFINAL.find(a => a.CLCL_ID === CLCL_ID)
        // })

        

        console.log(riClaims);
        console.log(orthoClaims);
        console.log(nonAdjusted);
        console.log(mca);
        //console.log(mcaFINAL);
        //console.log(filteredData);

        createReport();

    }

    async function createReport() {
        const fs = require('fs');

        try {
            if (!fs.existsSync(desktopDir + 'OUT')) {
                fs.mkdirSync(desktopDir + 'OUT');
            }
            } catch (err) {
            console.error(err);
            }

        const options = {
            filename: desktopDir + 'OUT\\Retroterm ' + monthFN + ' ' + yearFN + '.xlsx',
            useStyles: true,
            useSharedStrings: true
        };

        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
        const rawDataWS = workbook.addWorksheet("RawData");
        const orthoWS = workbook.addWorksheet("Ortho");
        const nonAdjustWS = workbook.addWorksheet("NonAdjustment");
        const rhodeIWS = workbook.addWorksheet("Rhode Island");
        const mcaWS = workbook.addWorksheet("MCA_Final");
        var excelColumns = [];
        var excelColumnsNonRaw = [];
        //var testCount = 0;

        // const options2 = {
        //     filename: desktopDir + 'OUT\\MCA_'+ monthFN + '_' + yearFN +'.xlsx',
        //     useStyles: true,
        //     useSharedStrings: true
        // };

        // const workbookMCA = new ExcelJS.stream.xlsx.WorkbookWriter(options2);
        // const mcaFinalWS = workbookMCA.addWorksheet("MCA");



        columnHeaders.forEach(value => {
            excelColumns.push({header: value, key: value, width: 15});
        });

        columnHeaders.forEach(value => {
            if (value != "ACTIVE_FLAG" && value != "J11_FLAG") {
                excelColumnsNonRaw.push({header: value, key: value, width: 15});
            }  
        });
        excelColumnsNonRaw.push({header: 'RCVRY_NUM', key: 'RCVRY_NUM', width: 15});
        //console.log(excelColumns);
        rawDataWS.columns = excelColumns;
        orthoWS.columns = excelColumnsNonRaw;
        nonAdjustWS.columns = excelColumnsNonRaw;
        rhodeIWS.columns = [{header: 'CLCL_ID', key: 'CLCL_ID', width: 15}];
        mcaWS.columns = [{header: 'CLCL_ID', key: 'CLCL_ID', width: 15}];
        //mcaFINAL.push({"CLCL_ID":row.CLCL_ID, "ADJ_R":adjReasonG, "col_3":"39", "col_4":"23", "col_5":"", "col_6":"", "date_add_1":"DATE", "proj_note":projectNoteG, "col_9": "A"});
        // mcaFinalWS.columns = [{header: '', key: 'CLCL_ID', width: 15}, {header: '', key: 'ADJ_R', width: 15}, {header: '', key: 'col_3', width: 15},
        //                     {header: '', key: 'col_4', width: 15}, {header: '', key: 'col_5', width: 15}, {header: '', key: 'col_6', width: 15},
        //                     {header: '', key: 'date_add_1', width: 15}, {header: '', key: 'proj_note', width: 15}, {header: '', key: 'col_9', width: 15}];

        console.log("STARTING RAW DATA INSERT!");
        for (let i = 0; i < rawData.length; i++) {
            rawDataWS.addRow(rawData[i]).commit();
        }

        rawDataWS.commit();
        
        console.log("STARTING ORTHO");
        for (let i = 0; i < orthoClaims.length; i++) {
            orthoWS.addRow(orthoClaims[i]).commit();
        }
        
        orthoWS.commit();
        
        console.log("STARTING NonAdjust");
        for (let i = 0; i < nonAdjusted.length; i++) {
            nonAdjustWS.addRow(nonAdjusted[i]).commit();
        }
        
        nonAdjustWS.commit();

        console.log("STARING RI CLAIMS");
        for (let i = 0; i < riClaims.length; i++) {
            rhodeIWS.addRow(riClaims[i]).commit();
            
        }
        
        rhodeIWS.commit();

        console.log("STARTING MCA");
        for (let i = 0; i < mca.length; i++) {
            mcaWS.addRow(mca[i]).commit();
            mcaFINALString += mca[i].CLCL_ID + "\t" + adjReasonG + "\t" + "39" + "\t" + "23" + "\t" + "" + "\t" + "" + "\t" + tomorrowsDate + "\t\"" + projectNoteG + "\"" + "\t" + "A" + "\n";
            
        }
        
        mcaWS.commit();
        console.log("STARTING SAVE");
        await workbook.commit();
        console.log("COMPLETE!");

        console.log("STARTING MCA TEMPLATE");
        fs.writeFile(desktopDir + 'OUT\\MCA_Input_File.txt', mcaFINALString, function (err) {
            if (err) throw err;
        });
        // for (let i = 0; i < mcaFINAL.length; i++) {
        //     mcaFinalWS.addRow(mcaFINAL[i]).commit();
        // }

        // mcaFinalWS.commit();
        // await workbookMCA.commit();
        //console.log(mcaFINALString);
        console.log("MCA FILE COMPLETE");


        rawData = [];
        filteredData = [];
        columnHeaders = [];
        riClaims = [];
        orthoClaims = []
        nonAdjusted = [];
        mca = [];
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
