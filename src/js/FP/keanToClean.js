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
    //var submitBtn = document.getElementById('keanClean').addEventListener("click", startKeanClean); // Set listener on Keen Button


    // Options for Workbook Writer. Pass it file name. useStyles and useSharedStrings must be true.
    var workbook;
    var levelMatchesWS;
    startKeanClean();

    async function startKeanClean() {
        console.log("startng");
        var levelMatches = []; // Sheet data var. We pass this our hashtables from the parsed excel file to write to the new file.
        var nRemoved = [];
        var duplicatesRemoved = [];
        var currentWS = ''; // Current Worksheet being iterated over
        var wsCount = 0; // To check if on next worksheet
        var dupeCheck = {}; // Pass this hashtable/associative array SSNs as the Key, then assign a value of "". This is to check for Duplicates
        document.getElementById("loadingMessage").innerHTML = "Processing: ";
        loadingIcon(true);
        // Grab file path and name.
        filePath = document.getElementById('file_upload').files[0].path;
        fileName = document.getElementById('file_upload').files[0].name; 

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
        
        var rowCount = 0; // To check if pass headers. SUBJECT TO CHANGE
        var rowCellData = []; // Pass individual cell values as strings to this array of the current row to build a final hash table off of.
        var finalRowData = {}; // Hash Table to put in our global array of sheetData[]. Each entry contains SSN, Ln, Fn, DOB, Group#, and Group.
        var columnHeaders = []; // Column headers for new Excel File
        
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
                    //workbookReader.end();
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
                levelMatches = []; // Reset sheetData to save memory
                rowCount = 0; // Set row count to 0 for new sheet
            }
            
            currentWS = worksheet.name; // Get Worksheet Name
            console.log('Current sheet: ' + currentWS);
            messageToUser('Sheet: ' + currentWS);
            
            worksheet.on('row', row => { // ExcelJS loop to iterate over rows
                finalRowData = {}; // Reset hashtable for new row
                
                if (rowCount == 0) {
                    row.eachCell(function(cell) {
                        columnHeaders.push(cell.text);
                    });
                    console.log(columnHeaders);
                    console.log(columnHeaders[55]);
                }
                
            
                if (rowCount > 0) {
                    // We need to push these values to an Array so we can index what specific cell to grab. Otherwise rows may not have the same amount of cells
                    // due to empty cells
                    if (row.getCell(53).text != '') {
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
                            }
                        }
                        //console.log(finalRowData);
                    }
                    
                    //////////52!
                    //55!

                    // if (String(rowCellData[0]).length < 9 && String(rowCellData[0]) != "") {
                    //     if (String(rowCellData[0]) === "NULL") {
                    //         finalRowData['SSN'] = String(rowCellData[0]);
                    //     } else {
                    //         var ssnLength = 9;
                    //         var newSSN = "";
                    //         ssnLength -= String(rowCellData[0]).length;
                    //         for (var i = 0; i < ssnLength; i++) {
                    //             newSSN += "0";
                    //         }
                            
                    //         newSSN += String(rowCellData[0]);
                    //         newSSN = newSSN.slice(0, 3) + "-" + newSSN.slice(3)
                    //         newSSN = newSSN.slice(0, 6) + "-" + newSSN.slice(6)
                    //         //console.log("ADDED ZEROS: " + newSSN);
                    //         finalRowData['SSN'] = newSSN;
                    //     }
                    // } else {
                    //     if (String(rowCellData[0]) === "") {
                    //         finalRowData['SSN'] = String(rowCellData[0]);
                    //     } else {
                    //         if (String(rowCellData[0]).includes('-')) {
                    //             finalRowData['SSN'] = String(rowCellData[0]);
                    //         } else {
                    //             var newSSN = String(rowCellData[0]);
                    //             newSSN = newSSN.slice(0, 3) + "-" + newSSN.slice(3)
                    //             newSSN = newSSN.slice(0, 6) + "-" + newSSN.slice(6)
                    //             //console.log("NO ADDED ZEROS: " + newSSN);
                    //             finalRowData['SSN'] = newSSN;
                    //         }
                            
                    //     }
                        
                    // }
                    
                    // finalRowData['Last Name'] = String(rowCellData[1]);
                    // finalRowData['First Name'] = String(rowCellData[2]);

                    // if (rowCellData[3] != "") {
                    //     var DOB = new Date(0, 0, rowCellData[3] - 1, 0, 0, 0);
                    //     var day = DOB.getDate();
                    //     var month = DOB.getMonth() + 1;
                    //     var year = DOB.getFullYear();
                    //     DOB = (month + "/" + day + "/" + year);
                    //     if (isNaN(day) || isNaN(month) || isNaN(year)){
                    //         finalRowData['DOB'] = rowCellData[3];
                    //     } else {
                    //         finalRowData['DOB'] = String(DOB);
                    //     }
                    // } else {
                    //     finalRowData['DOB'] = "";
                    // }
                    

                    
                    // finalRowData['Group #'] = String(rowCellData[10]);
                    // finalRowData['Group'] = String(rowCellData[11]);

            
                    //dupeCheck[String(rowCellData[0])] = "";   Was initially for filtering Duplicates
                    //sheetData.push(finalRowData);
                    //rowCellData = [];

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
            console.log(levelMatches);
            messageToUser("STARTING FINAL WS COMMIT");
            console.log("STARTING FINAL WS COMMIT");
            commitDataToLevelMatches(currentWS, levelMatches, columnHeaders);
            levelMatchesWS.commit();
            console.log("END COMMITTING WORKBOOK");
            messageToUser("END COMMITTING WORKBOOK");
            workbook.commit();
            loadingIcon(false);
            workbook = null;
        });
        workbookReader.on('error', (err) => {
        // ...
        });

    }

    async function commitDataToLevelMatches(workSheetName, workSheetData) {

        

        // Create the columns from an Array that already has the column values. Do this to style the headers in Excel
        // var excelColumns = [];
        // columnHeaders.forEach(value => {
        //     excelColumns.push({header: value, key: value, width: 15});
        // });
        // levelMatchesWS.columns = excelColumns;
        //For Commiting Worksheets to the new Excel File. Pass the name of the sheet    
        // being iterated over.

        // Create Worksheet and add it to the Workbook we created as a writable stream on Line 32
        //var worksheetToCommit = workbook.addWorksheet(workSheetName);

        // // Create the columns from an Array that already has the column values. Do this to style the headers in Excel
        // var excelColumns = [];
        // columnHeaders.forEach(value => {
        //     excelColumns.push({header: value, key: value, width: 15});
        // });
        // worksheetToCommit.columns = excelColumns;

        // Commit each row of data to the worksheet
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
            filename: desktopDir + 'OUT\\All Results Q2 2022.xlsx',
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
            }, 8000);
        }
    }

    async function messageToUser(message) {

        document.getElementById("consoleMessage").innerHTML = message.substring(0, 18);
    }
}