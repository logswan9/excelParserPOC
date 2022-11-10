export function loadLost() {
    const ExcelJS = require('exceljs');
    const homeDir = require('os').homedir(); 
    const desktopDir = `${homeDir}\\Desktop\\`;
    const electron = require('electron');
    const ipcRen = electron.ipcRenderer;
    const fs = require('fs');
    const { stringify } = require('querystring');
    //import {loadKean} from "../FP/keanToClean.js";

    var filePath = "";
    var fileName = "";
    //var submitBtn = document.getElementById('lostPolicy').addEventListener("click", startLostPolicy); // Set listener on Lost Policy Button
    //var submitBtn2 = document.getElementById('keanClean').addEventListener("click", startKeanClean);

    var events = false;
    var started = false;
    //var submitBtn = document.getElementById('keanClean').addEventListener("click", startKeanClean); // Set listener on Keen Button
    document.getElementById("selectionDiv").style.display = "none";
    document.getElementById("wrapper2FPLost").style.display = "block";
    document.getElementById("title2").style.display = "block";
    document.getElementById("title2").innerHTML = "Lost Policy";
    var exportFileName;
    

    var cancelButton = document.getElementById("cancelFPLost");
    var submitButton = document.getElementById("submitFPLost");
    
    if (!events) {
        cancelButton.addEventListener('click', function () {
            //document.getElementById("wrapper2").style.display = "none";
            location.reload();
        });
    
        submitButton.addEventListener('click', function () {
            if (!started) {
                document.getElementById("noFileMessageLost").textContent = "";

                try {
                    filePath = document.getElementById('file_upload').files[0].path;
                    fileName = document.getElementById('file_upload').files[0].name;
                } catch (error) {
                    setTimeout(function () {
                        document.getElementById("noFileMessageLost").textContent = "No File selected! Please select a file.";
                    }, 1000);
                    
                }
                


                if (filePath !== "") {
                    if (fileName.endsWith(".xlsx")) {
                        exportFileName = document.getElementById("lostPolicyName").value;
                        if (exportFileName !== "") {
                            startLostPolicy();
                            started = true;
                            submitButton.style.display = "none";
                        } else {
                            setTimeout(function () {
                                document.getElementById("noFileMessageLost").textContent = "No file/sheet name set. Please name the file/sheet.";
                            }, 1000);
                        }
                        
                    } else {
                        setTimeout(function () {
                            document.getElementById("noFileMessageLost").textContent = "ERROR! You must select a xlsx file.";
                        }, 1000);
                    }
                    
                } 
                
            }
        });
    }

    events = true;

    // Options for Workbook Writer. Pass it file name. useStyles and useSharedStrings must be true.
    var workbook;
    var worksheetToCommit;
    var worksheetToCommit2;
    var sheetThreshholdCheck = 1;
    var secondWorkSheet = false;

    // function startKeanClean() {

    //     loadKean();
    // }
    
    function startLostPolicy() {
        var sheetData = []; // Sheet data var. We pass this our hashtables from the parsed excel file to write to the new file.
        var currentWS = ''; // Current Worksheet being iterated over
        var wsCount = 0; // To check if on next worksheet
        //var dupeCheck = {}; // Pass this hashtable SSNs as the Key, then assign a value of "". This is to check for Duplicates
        document.getElementById("loadingMessage").innerHTML = "Processing: ";
        loadingIcon(true);
        // Grab file path and name.
        // filePath = document.getElementById('file_upload').files[0].path;
        // fileName = document.getElementById('file_upload').files[0].name;

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
        var columnHeaders = ['SSN', 'Last Name', 'First Name', 'DOB', 'Group #', 'Group']; // Column headers for new Excel File
        
        
        // ExcelJS loop to iterate over each work sheet
        workbookReader.on('worksheet', worksheet => {
            
            
            // If we have iterated over the first worksheet, commit that data to a new worksheet for new Excel File
            if (wsCount > 0) {
                if (wsCount == 1) { // Create Excel File and OUT Folder on desktop to write to
                    console.log("Creating Excel File");
                    messageToUser("Creating Excel File");
                    setWritableWorkbook(columnHeaders);
                }
                console.log(sheetData);
                messageToUser("STARTING COMMIT");
                console.log("STARTING COMMIT");
                commitSheetToWorkbook(currentWS, sheetData, columnHeaders); // Function call for writing to new excel file. Pass Worksheet Name
                sheetData = []; // Reset sheetData to save memory
                rowCount = 0; // Set row count to 0 for new sheet
            }
            
            currentWS = worksheet.name; // Get Worksheet Name
            console.log('Current sheet: ' + currentWS);
            messageToUser('Sheet: ' + currentWS);
            
            worksheet.on('row', row => { // ExcelJS loop to iterate over rows
                finalRowData = {}; // Reset hashtable for new row
                
            
                if (rowCount > 0) {
                    // We need to push these values to an Array so we can index what specific cell to grab. Otherwise rows may not have the same amount of cells
                    // due to empty cells
                    row.eachCell({includeEmpty: true}, function(cell, cellIndex) { // Iterate over each cell, include empty cells
                        if (cell.value === null || cell.value === "") {
                            rowCellData.push('');
                        } else {
                            rowCellData.push(String(cell.text));
                        }
                    });

                    if (String(rowCellData[0]).length < 9 && String(rowCellData[0]) != "") {
                        if (String(rowCellData[0]) === "NULL") {
                            finalRowData['SSN'] = String(rowCellData[0]);
                        } else {
                            var ssnLength = 9;
                            var newSSN = "";
                            ssnLength -= String(rowCellData[0]).length;
                            for (var i = 0; i < ssnLength; i++) {
                                newSSN += "0";
                            }
                            
                            newSSN += String(rowCellData[0]);
                            newSSN = newSSN.slice(0, 3) + "-" + newSSN.slice(3)
                            newSSN = newSSN.slice(0, 6) + "-" + newSSN.slice(6)
                            //console.log("ADDED ZEROS: " + newSSN);
                            finalRowData['SSN'] = newSSN;
                        }
                    } else {
                        if (String(rowCellData[0]) === "") {
                            finalRowData['SSN'] = String(rowCellData[0]);
                        } else {
                            if (String(rowCellData[0]).includes('-')) {
                                finalRowData['SSN'] = String(rowCellData[0]);
                            } else {
                                var newSSN = String(rowCellData[0]);
                                newSSN = newSSN.slice(0, 3) + "-" + newSSN.slice(3)
                                newSSN = newSSN.slice(0, 6) + "-" + newSSN.slice(6)
                                //console.log("NO ADDED ZEROS: " + newSSN);
                                finalRowData['SSN'] = newSSN;
                            }
                            
                        }
                        
                    }
                    
                    finalRowData['Last Name'] = String(rowCellData[1]);
                    finalRowData['First Name'] = String(rowCellData[2]);

                    if (String(rowCellData[3]) != "") {
                        var DOB = new Date(0, 0, rowCellData[3] - 1, 0, 0, 0);
                        var day = DOB.getDate();
                        var month = DOB.getMonth() + 1;
                        var year = DOB.getFullYear();
                        DOB = (month + "/" + day + "/" + year);
                        if (isNaN(day) || isNaN(month) || isNaN(year)){
                            finalRowData['DOB'] = String(rowCellData[3]);
                        } else {
                            finalRowData['DOB'] = String(DOB);
                        }
                    } else {
                        finalRowData['DOB'] = "";
                    }
                    

                    
                    finalRowData['Group #'] = String(rowCellData[10]);
                    finalRowData['Group'] = String(rowCellData[11]);

            
                    //dupeCheck[String(rowCellData[0])] = "";   Was initially for filtering Duplicates
                    sheetData.push(finalRowData);
                    rowCellData = [];

                    // if (rowCount >28) {
                    //     workbookReader.end();     //FOR DEBUGGING STOPS AFTER 28 ROWS
                    // }
                }
                rowCount++;
                
            }); // End on new row
            
            wsCount++;
            
        }); // End on new Worksheet
        
        
        workbookReader.on('end', () => {
            console.log(sheetData);
            messageToUser("STARTING FINAL WS COMMIT");
            console.log("STARTING FINAL WS COMMIT");
            console.log(sheetData);
            commitSheetToWorkbook(currentWS, sheetData, columnHeaders);
            worksheetToCommit.commit();
            if (secondWorkSheet) {
                worksheetToCommit2.commit();
            }
            console.log("END COMMITTING WORKBOOK");
            messageToUser("END COMMITTING WORKBOOK");
            workbook.commit();
            loadingIcon(false);
            workbook = null;
        });
        workbookReader.on('error', (err) => {
            console.log(err);
        });

    }

    function commitSheetToWorkbook(workSheetName, workSheetData, columnHeaders) {
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
        cancelButton.style.display = "none";
        for (let i = 0; i < workSheetData.length; i++) {
            if (sheetThreshholdCheck < 1048576) {
                //console.log("UNDER THRESHOLD");
                var blankFieldCount = 0;
                for (var key in workSheetData[i]) {
                    var value = workSheetData[i][key];
                    if (value === "") {
                        blankFieldCount++;      // Checking for mostly blank rows
                    }
                }
                if (blankFieldCount <= 3) {
                    worksheetToCommit.addRow(workSheetData[i]).commit();
                    sheetThreshholdCheck++;
                }
            } else {
                //console.log("OVER THRESHOLD");
                if (secondWorkSheet === false) {
                    console.log("CREATING SECONDWORKSHEET");
                    worksheetToCommit2 = workbook.addWorksheet(exportFileName + ' 2');
                    // Create the columns from an Array that already has the column values. Do this to style the headers in Excel
                    var excelColumns = [];
                    columnHeaders.forEach(value => {
                        excelColumns.push({header: value, key: value, width: 15});
                    });
                    worksheetToCommit2.columns = excelColumns;
                    secondWorkSheet = true;
                }
                var blankFieldCount = 0;
                for (var key in workSheetData[i]) {
                    var value = workSheetData[i][key];
                    if (value === "") {
                        blankFieldCount++;      // Checking for mostly blank rows
                    }
                }
                if (blankFieldCount <= 3) {
                    worksheetToCommit2.addRow(workSheetData[i]).commit();
                    sheetThreshholdCheck++;
                }

            }
            
        }
        
        

        // Commit the Worksheet after all rows have been committed
        //worksheetToCommit.commit();
        //messageToUser("WORKSHEET COMMITTED");
        console.log("WORKSHEET COMMITTED");
        workSheetData = null;
        columnHeaders = null;

    }

    function setWritableWorkbook(columnHeaders) {

        // To create OUT folder on desktop for testing
        try {
            if (!fs.existsSync(desktopDir + 'OUT')) {
                fs.mkdirSync(desktopDir + 'OUT');
            }
        } catch (err) {
            console.error(err);
        }

        const options = {
            filename: desktopDir + 'OUT\\' + exportFileName + '.xlsx',
            useStyles: true,
            useSharedStrings: true
        };
        workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options); // Create workbook object for writing new Excel file. Pass options
        worksheetToCommit = workbook.addWorksheet(exportFileName);

        // Create the columns from an Array that already has the column values. Do this to style the headers in Excel
        var excelColumns = [];
        columnHeaders.forEach(value => {
            excelColumns.push({header: value, key: value, width: 15});
        });
        worksheetToCommit.columns = excelColumns;
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