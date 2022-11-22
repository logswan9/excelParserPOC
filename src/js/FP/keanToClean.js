export function loadKean() {
    const ExcelJS = require('exceljs');
    const homeDir = require('os').homedir(); 
    const desktopDir = `${homeDir}\\Desktop\\`;
    const electron = require('electron');
    const ipcRen = electron.ipcRenderer;
    const fs = require('fs');
    const { stringify } = require('querystring');


    var filePath = "";
    var fileName = "";
    var filePath2 = "";
    var fileName2 = "";
    var events = false;
    var started = false;
    //var submitBtn = document.getElementById('keanClean').addEventListener("click", startKeanClean); // Set listener on Keen Button
    document.getElementById("selectionDiv").style.display = "none";
    document.getElementById("wrapper2FP").style.display = "block";
    document.getElementById("title2").style.display = "block";
    document.getElementById("title2").innerHTML = "Kean To Clean";

    var cancelButton = document.getElementById("cancelFPKean");
    var submitButton = document.getElementById("submitFPKean");
    
    if (!events) {
        cancelButton.addEventListener('click', function () {
            //document.getElementById("wrapper2").style.display = "none";
            location.reload();
        });
    
        submitButton.addEventListener('click', function () {
            if (!started) {
                document.getElementById("noFileMessageKean").textContent = "";

                try {
                    filePath = document.getElementById('file_upload').files[0].path;
                    fileName = document.getElementById('file_upload').files[0].name;

                    try {
                        filePath2 = document.getElementById('file_upload2').files[0].path;
                        fileName2 = document.getElementById('file_upload2').files[0].name;
                    } catch (error) {
                        setTimeout(function () {
                            document.getElementById("noFileMessageKean").textContent = "No Life Claims Paid File selected! Please select a file.";
                        }, 1000);
                    }
                } catch (error) {
                    setTimeout(function () {
                        document.getElementById("noFileMessageKean").textContent = "No File selected! Please select a file.";
                    }, 1000);
                    
                }
                


                if (filePath !== "") {
                    if (fileName.endsWith(".xlsx")) {
                        if (filePath2 !== "") {
                            if (fileName2.endsWith("xlsx")) {
                                startKeanClean();
                                started = true;
                                submitButton.style.display = "none";
                            } else {
                                setTimeout(function () {
                                    document.getElementById("noFileMessageKean").textContent = "ERROR! Life Claims Paid file must be a xlsx file.";
                                }, 1000);
                            }
                        }
                        
                    } else {
                        setTimeout(function () {
                            document.getElementById("noFileMessageKean").textContent = "ERROR! You must select a xlsx file.";
                        }, 1000);
                    }
                    
                } 

            }
        });
    }

    events = true;


    // Options for Workbook Writer. Pass it file name. useStyles and useSharedStrings must be true.
    var workbook;
    var levelMatchesWS;
    var nRemovedWS;
    var duplicatesRemovedWS;
    var dodRemovedWS;
    var lifeClaimsRemovedWS;
    var quarter;
    //startKeanClean();

    async function startKeanClean() {
        console.log("COMMENCING");
        var levelMatches = []; // Sheet data var. We pass this our hashtables from the parsed excel file to write to the new file.
        var nRemoved = [];
        var duplicatesRemoved = [];
        var dodRemoved = [];
        var lifeClaimsPaidRemoved = [];
        quarter = parseInt(document.getElementById("form1FP").value);
        console.log(quarter);
        var lookBack = 0;
        var lookBackDate;
        var quarterEndDate;
        var currentYear = new Date;
        var monthLB;
        var dayLB;
        currentYear = currentYear.getFullYear();
        switch (quarter) {
            case 1:
                quarterEndDate = new Date(+currentYear, 3 - 1, +31); //December 31, 2021 - March 31, 2022
                monthLB = quarterEndDate.getMonth() + 1;
                dayLB = quarterEndDate.getDate();
                lookBackDate = new Date(+currentYear, monthLB - 4, +dayLB);
                console.log(lookBackDate);
                break;
            case 2:
                quarterEndDate = new Date(+currentYear, 6 - 1, +30); //December 31, 2020 - June 30, 2022
                monthLB = quarterEndDate.getMonth() + 1;
                dayLB = quarterEndDate.getDate();
                lookBackDate = new Date(+currentYear, monthLB - 19, (+dayLB) + 1);
                console.log(lookBackDate);
                break;
            case 3:
                quarterEndDate = new Date(+currentYear, 9 - 1, +30); //June 30 - September 30
                monthLB = quarterEndDate.getMonth() + 1;
                dayLB = quarterEndDate.getDate();
                lookBackDate = new Date(+currentYear, monthLB - 4, +dayLB);
                console.log(lookBackDate);
                break;
                
            case 4:
                quarterEndDate = new Date(+currentYear, 12 - 1, +31); //June 30, 2021 - December 31, 2022
                monthLB = quarterEndDate.getMonth() + 1;
                dayLB = quarterEndDate.getDate();
                lookBackDate = new Date(+currentYear, monthLB - 19, (+dayLB) - 1);
                console.log(lookBackDate);
                break;
            default:
                break;
        }
        // if (quarter == 2 || quarter == 4) {
        //     lookBack = 18;
        // } else {
        //     lookBack = 3;
        // }
        console.log(lookBackDate + "\n" + quarterEndDate);
        var currentWS = ''; // Current Worksheet being iterated over
        var wsCount = 0; // To check if on next worksheet
        var dupeCheck = {}; // Pass this hashtable/associative array SSNs as the Key, then assign a value of "". This is to check for Duplicates
        var dupeCheck2ndFile = {};
        document.getElementById("loadingMessage").innerHTML = "Processing: ";
        loadingIcon(true);
        // Grab file path and name.
        // filePath = document.getElementById('file_upload').files[0].path;
        // fileName = document.getElementById('file_upload').files[0].name;

        // filePath2 = document.getElementById('file_upload2').files[0].path;
        // fileName2 = document.getElementById('file_upload2').files[0].name;

        // START EXCEL LOGIC

        // Options for the workbook reader. Options must be like this.
        const options = {
            sharedStrings: 'cache',
            hyperlinks: 'emit',
            worksheets: 'emit',
            entries: 'emit'
        };
        // workbookReader object. Pass the filePath of Excel File to Parse, pass the options.
        const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, options); 
        workbookReader.read(); // Opens Excel File to parse

        const workbookReader2 = new ExcelJS.stream.xlsx.WorkbookReader(filePath2, options); 
        workbookReader2.read();

        workbookReader2.on('worksheet', worksheet => {
            
            worksheet.on('row', row => { // ExcelJS loop to iterate over rows
                dupeCheck2ndFile[row.getCell(2).text] = "";
                
            }); // End on new row
            
            
        }); // End on new Worksheet
        
        
        workbookReader2.on('end', () => {
            console.log("Parsed 2nd file");
            console.log(dupeCheck2ndFile);
        });
        workbookReader2.on('error', (err) => {
            console.log(err);
        });
        
        var rowCount = 0; // To check if pass headers. SUBJECT TO CHANGE
        var rowCellData = []; // Pass individual cell values as strings to this array of the current row to build a final hash table off of.
        var finalRowData = {}; // Hash Table to put in our global array of sheetData[]. Each entry contains SSN, Ln, Fn, DOB, Group#, and Group.
        var columnHeaders = []; // Column headers for new Excel File
        var headersSet = false;
        var dodCount = 0;
        var headerIndex = 1;
        var mColumnIndex;
        
        // ExcelJS loop to iterate over each work sheet
        workbookReader.on('worksheet', worksheet => {
            
            
            // If we have iterated over the first worksheet, commit that data to a new worksheet for new Excel File
            if (wsCount > 0) {
                if (wsCount == 1) { // Create Excel File and OUT Folder on desktop to write to
                    console.log("Creating Excel File");
                    messageToUser("Creating Excel File");
                    console.log(levelMatches);
                    console.log(nRemoved);
                    console.log(dupeCheck);
                    console.log(duplicatesRemoved);
                    
                    setWritableWorkbook();
                    levelMatchesWS = workbook.addWorksheet("Level Matches");
                    var excelColumns = [];
                    columnHeaders.forEach(value => {
                        excelColumns.push({header: value, key: value, width: 15});
                    });
                    levelMatchesWS.columns = excelColumns;
                }
                console.log(levelMatches);
                messageToUser("STARTING COMMIT");
                console.log("STARTING COMMIT");
                commitDataToLevelMatches(currentWS, levelMatches); // Function call for writing to new excel file. Pass Worksheet Name
                //
                // console.log("END OF WORKBOOK");
                // messageToUser("STARTING FINAL WS COMMIT");
                // console.log("STARTING FINAL WS COMMIT");
                // commitDataToLevelMatches(currentWS, levelMatches);
                // levelMatchesWS.commit();
                // var excelColumns = [];
                // columnHeaders.forEach(value => {
                //     excelColumns.push({header: value, key: value, width: 15});
                // });
                // nRemovedWS = workbook.addWorksheet("N Removed");
                // duplicatesRemovedWS = workbook.addWorksheet("Duplicates Removed");
                // dodRemovedWS = workbook.addWorksheet("DODs Removed");
                // lifeClaimsRemovedWS = workbook.addWorksheet("Life Claims Paid Removed");
                // nRemovedWS.columns = excelColumns;
                // duplicatesRemovedWS.columns = excelColumns;
                // dodRemovedWS.columns = excelColumns;
                // lifeClaimsRemovedWS.columns = excelColumns;
                // commitOtherWorksheets(nRemoved, duplicatesRemoved, dodRemoved, lifeClaimsPaidRemoved)
                // nRemovedWS.commit();
                // duplicatesRemovedWS.commit();
                // dodRemovedWS.commit();
                // lifeClaimsRemovedWS.commit();
                // console.log("END COMMITTING WORKBOOK");
                // messageToUser("END COMMITTING WORKBOOK");
                // console.log(dodCount);
                // workbook.commit();
                // loadingIcon(false);
                // workbook = null;
                // console.log(inFileCount);
                // //
                // workbookReader.end();
                levelMatches = []; // Reset sheetData to save memory
                rowCount = 0; // Set row count to 0 for new sheet
            }
            
            currentWS = worksheet.name; // Get Worksheet Name
            console.log('Current sheet: ' + currentWS);
            messageToUser('Sheet: ' + currentWS);
            
            worksheet.on('row', row => { // ExcelJS loop to iterate over rows
                finalRowData = {}; // Reset hashtable for new row
                
                if (rowCount == 0 && !headersSet) {
                    row.eachCell(function(cell) {
                        columnHeaders.push(cell.text.trim());
                        if (cell.text === "M") {
                            mColumnIndex = headerIndex;  // Account for headers not always being the same index/order
                        }
                        headerIndex++;
                    });
                    console.log(row.getCell(52).text); //MIGHT NEED TO CHANGE
                    //console.log(columnHeaders[51]);
                    headersSet = true;
                    console.log(columnHeaders);
                }
                
            
                
                if (rowCount > 0) {
                    // We need to push these values to an Array so we can index what specific cell to grab. Otherwise rows may not have the same amount of cells
                    // due to empty cells
                    if (row.getCell(mColumnIndex).text != '') { 
                        //console.log(row.getCell(53).text);
                        row.eachCell({includeEmpty: true}, function(cell, cellIndex) { // Iterate over each cell, include empty cells
                            finalRowData[columnHeaders[(cellIndex - 1)]] = String(cell.text);
                            
                        });
                        levelMatches.push(finalRowData);
                        if (finalRowData['YNM'] != 'N') {
                            nRemoved.push(finalRowData);
                            
                            if (finalRowData['DMF SS'] in dupeCheck === false) {
                                dupeCheck[finalRowData['DMF SS']] = "";
                                duplicatesRemoved.push(finalRowData);

                                if (finalRowData['DMFDOD'].length == 8) {
                                    var month = finalRowData['DMFDOD'].substring(0, 2);
                                    var day = finalRowData['DMFDOD'].substring(2, 4);
                                    var year = finalRowData['DMFDOD'].substring(4);
                                    var date = new Date(+year, month - 1, +day);
                                    //console.log(date2);
                                    if (date.getTime() >= lookBackDate.getTime()) {
                                        //dodRemoved.push(finalRowData);
                                        dodRemoved.push({"SSN":finalRowData['SSN'], "Last Name":finalRowData['Last Name'], "First name":finalRowData['First name'], "DOB":finalRowData['DOB']
                                        , "Address":finalRowData['Address'], "City":finalRowData['City'], "State":finalRowData['State'], "Zip":finalRowData['Zip']
                                        , "Ind effective date":finalRowData['Ind effective date'], "Ind term date":finalRowData['Ind term date'], "Group #":finalRowData['Group #'], "Group":finalRowData['Group']
                                        , "Situs":finalRowData['Situs'], "Group Eff date":finalRowData['Group Eff date'], "Group term date":finalRowData['Group term date'], "Platform":finalRowData['Platform']
                                        , "DMFDOD":finalRowData['DMFDOD']});
                                        var checkSSN = finalRowData['SSN'].replaceAll("-", "");
                                        //console.log(checkSSN);
                                        if (checkSSN in dupeCheck2ndFile === false) {
                                            //console.log(finalRowData);
                                            //lifeClaimsPaidRemoved.push(finalRowData); //16 columns
                                            lifeClaimsPaidRemoved.push({"SSN":finalRowData['SSN'], "Last Name":finalRowData['Last Name'], "First name":finalRowData['First name'], "DOB":finalRowData['DOB']
                                            , "Address":finalRowData['Address'], "City":finalRowData['City'], "State":finalRowData['State'], "Zip":finalRowData['Zip']
                                            , "Ind effective date":finalRowData['Ind effective date'], "Ind term date":finalRowData['Ind term date'], "Group #":finalRowData['Group #'], "Group":finalRowData['Group']
                                            , "Situs":finalRowData['Situs'], "Group Eff date":finalRowData['Group Eff date'], "Group term date":finalRowData['Group term date'], "Platform":finalRowData['Platform']
                                            , "DMFDOD":finalRowData['DMFDOD']});
                                        }

                                    }
                                    
                                } else {
                                    console.log("DOD: " + finalRowData['DMFDOD'] + " Platform: " + finalRowData['Platform']);
                                    dodCount += 1;
                                }

                            }
                        }
                        //console.log(finalRowData);
                    }
                    
                    

                    // if (rowCount >450000) {
                    //     console.log(sheetData);
                    //     workbookReader.end();     //FOR DEBUGGING STOPS AFTER 28 ROWS
                    // }
                }
                rowCount++;
                
            }); // End on new row
            
            wsCount++;
            
        }); // End on new Worksheet
        
        
        workbookReader.on('end', () => {
                console.log("END OF WORKBOOK");
                messageToUser("STARTING FINAL WS COMMIT");
                console.log("STARTING FINAL WS COMMIT");
                commitDataToLevelMatches(currentWS, levelMatches);
                levelMatchesWS.commit();
                var excelColumns = [];
                var excelColumnsDOD = [];
                var columnCount = 0;
                columnHeaders.forEach(value => {
                    excelColumns.push({header: value, key: value, width: 15});
                });
                columnHeaders.forEach(value => {
                    if (columnCount <= 15) {
                        excelColumnsDOD.push({header: value, key: value, width: 15});
                        columnCount ++;
                    }
                    if (value === "DMFDOD") {
                        excelColumnsDOD.push({header: value, key: value, width: 15});
                    }
                    
                });
                nRemovedWS = workbook.addWorksheet("N Removed");
                duplicatesRemovedWS = workbook.addWorksheet("Duplicates Removed");
                dodRemovedWS = workbook.addWorksheet("DODs Removed");
                lifeClaimsRemovedWS = workbook.addWorksheet("Life Claims Paid Removed");
                nRemovedWS.columns = excelColumns;
                duplicatesRemovedWS.columns = excelColumns;
                dodRemovedWS.columns = excelColumnsDOD;
                lifeClaimsRemovedWS.columns = excelColumnsDOD;
                commitOtherWorksheets(nRemoved, duplicatesRemoved, dodRemoved, lifeClaimsPaidRemoved)
                nRemovedWS.commit();
                duplicatesRemovedWS.commit();
                dodRemovedWS.commit();
                lifeClaimsRemovedWS.commit();
                console.log("END COMMITTING WORKBOOK");
                messageToUser("END COMMITTING WORKBOOK");
                console.log(dodCount);
                workbook.commit();
                loadingIcon(false);
                workbook = null;
                console.log("COMPLETE!");
        });
        workbookReader.on('error', (err) => {
            console.log(err);
        });

    }

    async function commitDataToLevelMatches(workSheetName, workSheetData) {

        cancelButton.style.display = "none";

        console.log("Commiting Rows for: " + workSheetName);
        messageToUser("Commiting Rows for: " + String(workSheetName));
        for (let i = 0; i < workSheetData.length; i++) {
            levelMatchesWS.addRow(workSheetData[i]).commit();
        }

        // Commit the Worksheet after all rows have been committed
        //worksheetToCommit.commit();
        //messageToUser("WORKSHEET COMMITTED");
        console.log("WORKSHEET COMMITTED");
        workSheetData = null;
        //columnHeaders = null;

    }

    async function commitOtherWorksheets(n, dupesR, dod, lifeClaims) {
        for (let i = 0; i < n.length; i++) {
            nRemovedWS.addRow(n[i]).commit();
        }

        for (let i = 0; i < dupesR.length; i++) {
            duplicatesRemovedWS.addRow(dupesR[i]).commit();
        }

        for (let i = 0; i < dod.length; i++) {
            dodRemovedWS.addRow(dod[i]).commit();
        }

        for (let i = 0; i < lifeClaims.length; i++) {
            lifeClaimsRemovedWS.addRow(lifeClaims[i]).commit();
        }

    }

    async function setWritableWorkbook() {

        // To create OUT folder on desktop for testing
        try {
            if (!fs.existsSync(desktopDir + 'OUT')) {
                fs.mkdirSync(desktopDir + 'OUT');
            }
        } catch (err) {
            console.error(err);
        }

        const options = {
            filename: desktopDir + 'OUT\\All Results Q' + quarter + ' 2022.xlsx',
            useStyles: true,
            useSharedStrings: true
        };
        workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options); // Create workbook object for writing new Excel file. Pass options
        // worksheetToCommit = workbook.addWorksheet("Level Matches");

        // // Create the columns from an Array that already has the column values. Do this to style the headers in Excel
        // var excelColumns = [];
        // columnHeaders.forEach(value => {
        //     excelColumns.push({header: value, key: value, width: 15});
        // });
        // worksheetToCommit.columns = excelColumns;
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
            }, 8000);
        }
    }

    async function messageToUser(message) {

        document.getElementById("consoleMessage").innerHTML = message.substring(0, 18);
    }
}