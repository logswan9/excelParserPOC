import { Base } from './Base.js';

export function loadBlankCheck() {
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
    //document.getElementById("wrapper2CheckRun").style.display = "block";
    document.getElementById("title2").style.display = "block";
    document.getElementById("title2").innerHTML = "Blank Field Search";
    document.getElementById("wrapper2BlankField").style.display = "block";
    

    var cancelButton = document.getElementById("cancelBlankField");
    var submitButton = document.getElementById("submitBlankField");

    var monthFN = "";
    var yearFN = "";
    var adjReasonG = "";
    var projectNoteG = "";
    var rawData = [];
    var tomorrowsDate;
    var events = false;
    var columnHeaders = [];
    var files = [];

    var workbook;
    var reportWS;

    if (!events) {
        cancelButton.addEventListener('click', function () {
            //document.getElementById("wrapper2").style.display = "none";
            location.reload();
        });
    
        submitButton.addEventListener('click', function () {
            if (!started) {
                document.getElementById("noFileMessageBlankField").textContent = "";

                try {
                    filePath = document.getElementById('file_upload').files[0].path;
                    fileName = document.getElementById('file_upload').files[0].name;
                } catch (error) {
                    setTimeout(function () {
                        document.getElementById("noFileMessageBlankField").textContent = "No File selected! Please select a file.";
                    }, 1000);
                    
                }
                


                if (filePath !== "") {
                    // if (fileName.endsWith(".xlsx")) {
                        
                    //     started = true;
                    //     submitButton.style.display = "none";
                        
                        
                    // } else {
                    //     setTimeout(function () {
                    //         document.getElementById("noFileMessageCheck").textContent = "ERROR! You must select a xlsx file.";
                    //     }, 1000);
                    // }
                    loadingIcon(true);
                    beginBlankFieldSearch();
                } 
                
            }
        });
    }

    events = true;
    var finalData = [];

    async function beginBlankFieldSearch() {

        files = document.getElementById('file_upload').files;
        
        var fileData = [];
        var lineData = [];
        var data;
        
        var masterCount = 0;
        var employerGroupFile = "";

        var countSSN = 0;
        var rowsSSNEmpty = [];

        var countHireDate = 0;
        var rowsHireDateEmpty = [];

        var countWorkState = 0;
        var rowsWorkStateEmpty = [];

        var countEmployStatus = 0;
        var rowsEmployStatusEmpty = [];

        var countFTIndicator = 0;
        var rowsFTIndicatorEmpty = [];

        var countAnnHours = 0;
        var rowsAnnHoursEmpty = [];

        var countExemptStatus = 0;
        var rowsExemptStatusEmpty = [];

        var countDivClass = 0;
        var rowsDivClassEmpty = [];

        var countAnnSalary = 0;
        var rowsAnnSalaryEmpty = [];

        var countPayCycle = 0;
        var rowsPayCycleEmpty = [];

        var countDisPlanCode = 0;
        var rowsDisPlanCodeEmpty = [];
        var validRow = false;
        
        const fs = require('fs');

        createFile();

        for (let i = 0; i < document.getElementById('file_upload').files.length; i++) { // Loop through files
            data = String(fs.readFileSync(document.getElementById('file_upload').files[i].path)); // Read/parse file
            //if (err) throw err;
            fileData = data.split(/\r?\n/);
            for (var z = 0; z < fileData.length; z++) { // Loop through Line of data
                if (fileData[z].trim() == '') {
                    continue;
                }
                lineData = fileData[z].split('|');
                console.log(lineData.length);
                if (lineData.length == 76) {
                    try {
                        finalData.push({"SSN of Claimant":lineData[0], "SSN of Employee":lineData[1], "Claimant Last Name":lineData[2], "Claimant First Name":lineData[3], "Claimant Middle Name":lineData[4],
                        "Employee Last Name":lineData[5], "Employee First Name":lineData[6], "Employee Middle Name":lineData[7], "Home/Mailing Address 1":lineData[8], "Home/Mailing Address 2":lineData[9], "Home/Mailing City":lineData[10],
                        "Home/Mailing State":lineData[11], "Home/Mailing Zip":lineData[12], "Home/Contact Phone":lineData[13], "Gender":lineData[14], "DOB":lineData[15], "Effective Date of Coverage CIPP":lineData[16],
                        "Coverage Amount CIPP":lineData[17], "Plan Code APP":lineData[18], "Hire Date":lineData[19], "Work State":lineData[20], "Weekly Work Hours":lineData[21], "Employment Status":lineData[22],
                        "Employment Termination Date":lineData[23], "Home Email":lineData[24], "Latest payroll period":lineData[25], "FT/PT Indicator":lineData[26], "Annual Hours":lineData[27], "Job Title":lineData[28],
                        "Job Code":lineData[29], "Exempt Status":lineData[30],"Employers Department Level":lineData[31], "Division/Classes":lineData[32], "Cost Center Code":lineData[33], "Location/Subgroup":lineData[34],
                        "Building Code":lineData[35], "G/L Company Code":lineData[36], "Organization Level 1":lineData[37], "Organization Level 2":lineData[38], "Organization Level 3":lineData[39], "Pay Rate":lineData[40],
                        "Annual Salary":lineData[41], "Supervisor Last Name":lineData[42], "Supervisor First Name":lineData[43], "Supervisor Email":lineData[44], "Supervisor EE Identifier":lineData[45],
                        "Supervisor EE Identifier Code":lineData[46], "HR Rep EE Identifier":lineData[47], "HR Rep EE Identifier Code":lineData[48], "HR Rep Last Name":lineData[49], "HR Rep First Name":lineData[50],
                        "HR Rep Email address":lineData[51], "Rehire Date":lineData[52], "Pay Cycle":lineData[53], "Disability Plan Code":lineData[54], "Employer Name":lineData[55], "Null":lineData[56], 
                        "Effective Date of Coverage APP":lineData[57], "Coverage Term Date – CIPP":lineData[58], "Coverage Term Date – APP":lineData[59], "Prior Effective Date of Coverage CIPP":lineData[60],
                        "Prior Effective End Date of Coverage CIPP":lineData[61], "Prior Coverage Amount CIPP":lineData[62], "Prior – 2 Effective Date of Coverage CIPP":lineData[63], 
                        "Prior Effective End Date of Coverage – 2 CIPP":lineData[64], "Prior – 2 Coverage Amount CIPP":lineData[65], "Effective Date of Coverage HIPP":lineData[66], "Plan Selection HIPP":lineData[67], "HIPP Option":lineData[68],
                        "Effective Date of Coverage Term HIPP":lineData[69], "Prior Effective Date of Coverage HIPP":lineData[70],"Prior Effective End Date of Coverage – HIPP":lineData[71], "Prior Plan Selection(HIPP)":lineData[72], 
                        "Prior – 2 Effective Date of Coverage – HIPP":lineData[73], "Prior Effective End Date of Coverage – 2 HIPP":lineData[74], "Prior – 2 Plan Selection(HIPP)":lineData[75]});
                        if (z == 0 || z == fileData.length - 1) {
                            console.log({"SSN of Claimant":lineData[0], "SSN of Employee":lineData[1], "Claimant Last Name":lineData[2], "Claimant First Name":lineData[3], "Claimant Middle Name":lineData[4],
                            "Employee Last Name":lineData[5], "Employee First Name":lineData[6], "Employee Middle Name":lineData[7], "Home/Mailing Address 1":lineData[8], "Home/Mailing Address 2":lineData[9], "Home/Mailing City":lineData[10],
                            "Home/Mailing State":lineData[11], "Home/Mailing Zip":lineData[12], "Home/Contact Phone":lineData[13], "Gender":lineData[14], "DOB":lineData[15], "Effective Date of Coverage CIPP":lineData[16],
                            "Coverage Amount CIPP":lineData[17], "Plan Code APP":lineData[18], "Hire Date":lineData[19], "Work State":lineData[20], "Weekly Work Hours":lineData[21], "Employment Status":lineData[22],
                            "Employment Termination Date":lineData[23], "Home Email":lineData[24], "Latest payroll period":lineData[25], "FT/PT Indicator":lineData[26], "Annual Hours":lineData[27], "Job Title":lineData[28],
                            "Job Code":lineData[29], "Exempt Status":lineData[30],"Employers Department Level":lineData[31], "Division/Classes":lineData[32], "Cost Center Code":lineData[33], "Location/Subgroup":lineData[34],
                            "Building Code":lineData[35], "G/L Company Code":lineData[36], "Organization Level 1":lineData[37], "Organization Level 2":lineData[38], "Organization Level 3":lineData[39], "Pay Rate":lineData[40],
                            "Annual Salary":lineData[41], "Supervisor Last Name":lineData[42], "Supervisor First Name":lineData[43], "Supervisor Email":lineData[44], "Supervisor EE Identifier":lineData[45],
                            "Supervisor EE Identifier Code":lineData[46], "HR Rep EE Identifier":lineData[47], "HR Rep EE Identifier Code":lineData[48], "HR Rep Last Name":lineData[49], "HR Rep First Name":lineData[50],
                            "HR Rep Email address":lineData[51], "Rehire Date":lineData[52], "Pay Cycle":lineData[53], "Disability Plan Code":lineData[54], "Employer Name":lineData[55], "Null":lineData[56], 
                            "Effective Date of Coverage APP":lineData[57], "Coverage Term Date – CIPP":lineData[58], "Coverage Term Date – APP":lineData[59], "Prior Effective Date of Coverage CIPP":lineData[60],
                            "Prior Effective End Date of Coverage CIPP":lineData[61], "Prior Coverage Amount CIPP":lineData[62], "Prior – 2 Effective Date of Coverage CIPP":lineData[63], 
                            "Prior Effective End Date of Coverage – 2 CIPP":lineData[64], "Prior – 2 Coverage Amount CIPP":lineData[65], "Effective Date of Coverage HIPP":lineData[66], "Plan Selection HIPP":lineData[67], "HIPP Option":lineData[68],
                            "Effective Date of Coverage Term HIPP":lineData[69], "Prior Effective Date of Coverage HIPP":lineData[70],"Prior Effective End Date of Coverage – HIPP":lineData[71], "Prior Plan Selection(HIPP)":lineData[72], 
                            "Prior – 2 Effective Date of Coverage – HIPP":lineData[73], "Prior Effective End Date of Coverage – 2 HIPP":lineData[74], "Prior – 2 Plan Selection(HIPP)":lineData[75]});
                        }
                    } catch (error) {
                        console.log(error);
                    }
                } else {
                    console.log("ROW IS NOT CORRECT LENGTH");
                }
                
                
                
                if (lineData.length == 76) {
                    if (lineData[55] != '') {
                        try {
                            //console.log(document.getElementById('file_upload').files[i].name);
                            lineData[55].trim();
                            //employerGroupFile = lineData[55].replace('.', '');
                            //employerGroupFile = lineData[55].replaceAll('(', '');
                            employerGroupFile = lineData[55].replace(/["'().-]/g,"");
                            // employerGroupFile = employerGroupFile.trim();
                            // if (employerGroupFile === '') {
                            //     employerGroupFile = 'test';
                            // }
                            console.log(employerGroupFile);
                        } catch (error) {
                            console.log(error);
                        }
                        
                    } else {
                        employerGroupFile = document.getElementById('file_upload').files[i].name;
                        employerGroupFile = employerGroupFile.replace(/["'().-]/g,"");
                        employerGroupFile = employerGroupFile.substring(0, 12);
                        employerGroupFile = employerGroupFile.replace('_',"");
                        console.log(employerGroupFile);
                    }
                } else {
                    employerGroupFile = "Error: Invalid Row Data. FileName: " + document.getElementById('file_upload').files[i].name;
                }
                
            }
            masterCount = finalData.length;

            console.log(document.getElementById('file_upload').files[i].name);
            for (var t = 0; t < finalData.length; t++) {
                
                if (finalData[t]['SSN of Claimant'].trim() == '') {
                    countSSN++;
                    if (fileData.length < 20000) {
                        rowsSSNEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Hire Date'].trim() == '') {
                    countHireDate++;
                    if (fileData.length < 20000) {
                        rowsHireDateEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Work State'].trim() == '') {
                    countWorkState++;
                    if (fileData.length < 20000) {
                        rowsWorkStateEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Employment Status'].trim() == '') {
                    countEmployStatus++;
                    if (fileData.length < 20000) {
                        rowsEmployStatusEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['FT/PT Indicator'].trim() == '') {
                    countFTIndicator++;
                    if (fileData.length < 20000) {
                        rowsFTIndicatorEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Annual Hours'].trim() == '') {
                    countAnnHours++;
                    if (fileData.length < 20000) {
                        rowsAnnHoursEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Exempt Status'].trim() == '') {
                    countExemptStatus++;
                    if (fileData.length < 20000) {
                        rowsExemptStatusEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Division/Classes'].trim() == '') {
                    countDivClass++;
                    if (fileData.length < 20000) {
                        rowsDivClassEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Annual Salary'].trim() == '') {
                    countAnnSalary++;
                    if (fileData.length < 20000) {
                        rowsAnnSalaryEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Pay Cycle'].trim() == '') {
                    countPayCycle++;
                    if (fileData.length < 20000) {
                        rowsPayCycleEmpty.push(finalData[t]);
                    }
                }

                if (finalData[t]['Disability Plan Code'].trim() == '') {
                    countDisPlanCode++;
                    if (fileData.length < 20000) {
                        rowsDisPlanCodeEmpty.push(finalData[t]);
                    }
                }
                
                
                //console.log(finalData[t]['Claimant Middle Name'].trim());
                
                
            }

            let testRow = reportWS.addRow({'Employer Group':employerGroupFile, 'Total Records':masterCount, 'SSN of Claimant':countSSN, 'Hire Date':countHireDate, 'Work State':countWorkState,
            'Employment Status':countEmployStatus, 'FT/PT Indicator':countFTIndicator, 'Annual Hours':countAnnHours, 'Exempt Status':countExemptStatus, 'Division/Classes':countDivClass,
            'Annual Salary':countAnnSalary, 'Pay Cycle':countPayCycle, 'Disability Plan Code':countDisPlanCode});

            testRow.eachCell(function(cell) {
                console.log(cell);
                console.log(cell._column._key);
                if (cell._column._key === 'SSN of Claimant') {
                    cell.font = {
                        color: { argb: "f00a0a" },
                        bold: true,
                      };
                }
                
            });

            testRow.commit();


            if (fileData.length < 20000 && (countAnnHours > 0 || countAnnSalary > 0 || countDisPlanCode > 0 || countDivClass > 0 || countEmployStatus > 0 || countExemptStatus > 0 || countFTIndicator > 0 || countHireDate > 0 || 
                countPayCycle > 0 || countSSN > 0 || countWorkState > 0)) {
                console.log('Creating sheet for ' + employerGroupFile + ' File: ' + document.getElementById('file_upload').files[i].name);
                var groupWS = workbook.addWorksheet(employerGroupFile);
                groupWS.columns = [{header: 'SSN of Claimant', key: 'SSN of Claimant', width: 20}, {header: 'SSN of Employee', key: 'SSN of Employee', width: 15}, {header: 'Claimant Last Name', key: 'Claimant Last Name', width: 15}, 
                {header: 'Claimant First Name', key: 'Claimant First Name', width: 15}, {header: 'Claimant Middle Name', key: 'Claimant Middle Name', width: 15}, {header: 'Employee Last Name', key: 'Employee Last Name', width: 15}, 
                {header: 'Employee First Name', key: 'Employee First Name', width: 15}, {header: 'Employee Middle Name', key: 'Employee Middle Name', width: 15}, {header: 'Home/Mailing Address 1', key: 'Home/Mailing Address 1', width: 15}, 
                {header: 'Home/Mailing Address 2', key: 'Home/Mailing Address 2', width: 15}, {header: 'Home/Mailing City', key: 'Home/Mailing City', width: 15}, {header: 'Home/Mailing State', key: 'Home/Mailing State', width: 15}, 
                {header: 'Home/Mailing Zip', key: 'Home/Mailing Zip', width: 15}, {header: 'Home/Contact Phone', key: 'Home/Contact Phone', width: 15}, {header: 'Gender', key: 'Gender', width: 15}, 
                {header: 'DOB', key: 'DOB', width: 15}, {header: 'Effective Date of Coverage CIPP', key: 'Effective Date of Coverage CIPP', width: 15}, {header: 'Coverage Amount CIPP', key: 'Coverage Amount CIPP', width: 15}, 
                {header: 'Plan Code APP', key: 'Plan Code APP', width: 15}, {header: 'Hire Date', key: 'Hire Date', width: 15}, {header: 'Work State', key: 'Work State', width: 15}, 
                {header: 'Weekly Work Hours', key: 'Weekly Work Hours', width: 15}, {header: 'Employment Status', key: 'Employment Status', width: 15}, {header: 'Employment Termination Date', key: 'Employment Termination Date', width: 15}, 
                {header: 'Home Email', key: 'Home Email', width: 15}, {header: 'Latest payroll period', key: 'Latest payroll period', width: 15}, {header: 'FT/PT Indicator', key: 'FT/PT Indicator', width: 15}, 
                {header: 'Annual Hours', key: 'Annual Hours', width: 15}, {header: 'Job Title', key: 'Job Title', width: 15}, {header: 'Job Code', key: 'Job Code', width: 15}, 
                {header: 'Exempt Status', key: 'Exempt Status', width: 15}, {header: 'Employers Department Level', key: 'Employers Department Level', width: 15}, {header: 'Division/Classes', key: 'Division/Classes', width: 15}, 
                {header: 'Cost Center Code', key: 'Cost Center Code', width: 15}, {header: 'Location/Subgroup', key: 'Location/Subgroup', width: 15}, {header: 'Building Code', key: 'Building Code', width: 15}, 
                {header: 'G/L Company Code', key: 'G/L Company Code', width: 15}, {header: 'Organization Level 1', key: 'Organization Level 1', width: 15}, {header: 'Organization Level 2', key: 'Organization Level 2', width: 15}, 
                {header: 'Organization Level 3', key: 'Organization Level 3', width: 15}, {header: 'Pay Rate', key: 'Pay Rate', width: 15}, {header: 'Annual Salary', key: 'Annual Salary', width: 15}, 
                {header: 'Supervisor Last Name', key: 'Supervisor Last Name', width: 15}, {header: 'Supervisor First Name', key: 'Supervisor First Name', width: 15}, {header: 'Supervisor Email', key: 'Supervisor Email', width: 15}, 
                {header: 'Supervisor EE Identifier', key: 'Supervisor EE Identifier', width: 15}, {header: 'Supervisor EE Identifier Code', key: 'Supervisor EE Identifier Code', width: 15}, {header: 'HR Rep EE Identifier', key: 'HR Rep EE Identifier', width: 15}, 
                {header: 'HR Rep EE Identifier Code', key: 'HR Rep EE Identifier Code', width: 15}, {header: 'HR Rep Last Name', key: 'HR Rep Last Name', width: 15}, {header: 'HR Rep First Name', key: 'HR Rep First Name', width: 15}, 
                {header: 'HR Rep Email address', key: 'HR Rep Email address', width: 15}, {header: 'Rehire Date', key: 'Rehire Date', width: 15}, {header: 'Pay Cycle', key: 'Pay Cycle', width: 15}, 
                {header: 'Disability Plan Code', key: 'Disability Plan Code', width: 15}, {header: 'Employer Name', key: 'Employer Name', width: 15}, {header: 'Null', key: 'Null', width: 15}, 
                {header: 'Effective Date of Coverage APP', key: 'Effective Date of Coverage APP', width: 15}, {header: 'Coverage Term Date – CIPP', key: 'Coverage Term Date – CIPP', width: 15}, {header: 'Coverage Term Date – APP', key: 'Coverage Term Date – APP', width: 15}, 
                {header: 'Prior Effective Date of Coverage CIPP', key: 'Prior Effective Date of Coverage CIPP', width: 15}, {header: 'Prior Effective End Date of Coverage CIPP', key: 'Prior Effective End Date of Coverage CIPP', width: 15}, {header: 'Prior Coverage Amount CIPP', key: 'Prior Coverage Amount CIPP', width: 15}, 
                {header: 'Prior – 2 Effective Date of Coverage CIPP', key: 'Prior – 2 Effective Date of Coverage CIPP', width: 15}, {header: 'Prior Effective End Date of Coverage – 2 CIPP', key: 'Prior Effective End Date of Coverage – 2 CIPP', width: 15}, {header: 'Prior – 2 Coverage Amount CIPP', key: 'Prior – 2 Coverage Amount CIPP', width: 15}, 
                {header: 'Effective Date of Coverage HIPP', key: 'Effective Date of Coverage HIPP', width: 15}, {header: 'Plan Selection HIPP', key: 'Plan Selection HIPP', width: 15}, {header: 'HIPP Option', key: 'HIPP Option', width: 15}, 
                {header: 'Effective Date of Coverage Term HIPP', key: 'Effective Date of Coverage Term HIPP', width: 15}, {header: 'Prior Effective Date of Coverage HIPP', key: 'Prior Effective Date of Coverage HIPP', width: 15}, {header: 'Prior Effective End Date of Coverage – HIPP', key: 'Prior Effective End Date of Coverage – HIPP', width: 15}, 
                {header: 'Prior Plan Selection(HIPP)', key: 'Prior Plan Selection(HIPP)', width: 15}, {header: 'Prior – 2 Effective Date of Coverage – HIPP', key: 'Prior – 2 Effective Date of Coverage – HIPP', width: 15}, {header: 'Prior Effective End Date of Coverage – 2 HIPP', key: 'Prior Effective End Date of Coverage – 2 HIPP', width: 15}, 
                {header: 'Prior – 2 Plan Selection(HIPP)', key: 'Prior – 2 Plan Selection(HIPP)', width: 15}];
                

                if (rowsSSNEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing SSN", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsSSNEmpty.length; i++) {
                        groupWS.addRow(rowsSSNEmpty[i]).commit();
                    }
                }
                

                if (rowsHireDateEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Hire Date", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsHireDateEmpty.length; i++) {
                        groupWS.addRow(rowsHireDateEmpty[i]).commit();
                    }
                }
                

                if (rowsWorkStateEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Work State", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsWorkStateEmpty.length; i++) {
                        groupWS.addRow(rowsWorkStateEmpty[i]).commit();
                    }
                }
                

                if (rowsEmployStatusEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Employment Status", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsEmployStatusEmpty.length; i++) {
                        groupWS.addRow(rowsEmployStatusEmpty[i]).commit();
                    }
                }
                

                if (rowsFTIndicatorEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing FT/PT Indicator", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsFTIndicatorEmpty.length; i++) {
                        groupWS.addRow(rowsFTIndicatorEmpty[i]).commit();
                    }
                }
                

                if (rowsAnnHoursEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Annual Hours", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsAnnHoursEmpty.length; i++) {
                        groupWS.addRow(rowsAnnHoursEmpty[i]).commit();
                    }
                }
                

                if (rowsExemptStatusEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Exempt Status", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsExemptStatusEmpty.length; i++) {
                        groupWS.addRow(rowsExemptStatusEmpty[i]).commit();
                    }
                }
                

                if (rowsDivClassEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Division/Classes", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsDivClassEmpty.length; i++) {
                        groupWS.addRow(rowsDivClassEmpty[i]).commit();
                    }
                }
                

                if (rowsAnnSalaryEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Annual Salary", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsAnnSalaryEmpty.length; i++) {
                        groupWS.addRow(rowsAnnSalaryEmpty[i]).commit();
                    }
                }
                

                if (rowsPayCycleEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                        "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                        "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                        "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                        "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                        "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                        "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                        "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                        "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                        "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                        "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                        "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                        "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                        "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Pay Cycle", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsPayCycleEmpty.length; i++) {
                        groupWS.addRow(rowsPayCycleEmpty[i]).commit();
                    }
                }
                

                if (rowsDisPlanCodeEmpty.length > 0) {
                    groupWS.addRow({"SSN of Claimant":"", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    groupWS.addRow({"SSN of Claimant":"Missing Disability Plan Code", "SSN of Employee":"", "Claimant Last Name":"", "Claimant First Name":"", "Claimant Middle Name":"",
                    "Employee Last Name":"", "Employee First Name":"", "Employee Middle Name":"", "Home/Mailing Address 1":"", "Home/Mailing Address 2":"", "Home/Mailing City":"",
                    "Home/Mailing State":"", "Home/Mailing Zip":"", "Home/Contact Phone":"", "Gender":"", "DOB":"", "Effective Date of Coverage CIPP":"",
                    "Coverage Amount CIPP":"", "Plan Code APP":"", "Hire Date":"", "Work State":"", "Weekly Work Hours":"", "Employment Status":"",
                    "Employment Termination Date":"", "Home Email":"", "Latest payroll period":"", "FT/PT Indicator":"", "Annual Hours":"", "Job Title":"",
                    "Job Code":"", "Exempt Status":"","Employers Department Level":"", "Division/Classes":"", "Cost Center Code":"", "Location/Subgroup":"",
                    "Building Code":"", "G/L Company Code":"", "Organization Level 1":"", "Organization Level 2":"", "Organization Level 3":"", "Pay Rate":"",
                    "Annual Salary":"", "Supervisor Last Name":"", "Supervisor First Name":"", "Supervisor Email":"", "Supervisor EE Identifier":"",
                    "Supervisor EE Identifier Code":"", "HR Rep EE Identifier":"", "HR Rep EE Identifier Code":"", "HR Rep Last Name":"", "HR Rep First Name":"",
                    "HR Rep Email address":"", "Rehire Date":"", "Pay Cycle":"", "Disability Plan Code":"", "Employer Name":"", "Null":"", 
                    "Effective Date of Coverage APP":"", "Coverage Term Date – CIPP":"", "Coverage Term Date – APP":"", "Prior Effective Date of Coverage CIPP":"",
                    "Prior Effective End Date of Coverage CIPP":"", "Prior Coverage Amount CIPP":"", "Prior – 2 Effective Date of Coverage CIPP":"", 
                    "Prior Effective End Date of Coverage – 2 CIPP":"", "Prior – 2 Coverage Amount CIPP":"", "Effective Date of Coverage HIPP":"", "Plan Selection HIPP":"", "HIPP Option":"",
                    "Effective Date of Coverage Term HIPP":"", "Prior Effective Date of Coverage HIPP":"","Prior Effective End Date of Coverage – HIPP":"", "Prior Plan Selection(HIPP)":"", 
                    "Prior – 2 Effective Date of Coverage – HIPP":"", "Prior Effective End Date of Coverage – 2 HIPP":"", "Prior – 2 Plan Selection(HIPP)":""}).commit();
                    for (let i = 0; i < rowsDisPlanCodeEmpty.length; i++) {
                        groupWS.addRow(rowsDisPlanCodeEmpty[i]).commit();
                    }
                }
                groupWS.commit();
            }
            
            
            rowsSSNEmpty = [];
            rowsHireDateEmpty = [];
            rowsWorkStateEmpty = [];
            rowsEmployStatusEmpty = [];
            rowsFTIndicatorEmpty = [];
            rowsAnnHoursEmpty = [];
            rowsExemptStatusEmpty = [];
            rowsDivClassEmpty = [];
            rowsAnnSalaryEmpty = [];
            rowsPayCycleEmpty = [];
            rowsDisPlanCodeEmpty = [];
            finalData = [];
            employerGroupFile = '';

            masterCount = 0;
            countSSN = 0;
            countHireDate = 0;
            countWorkState = 0;
            countEmployStatus = 0;
            countFTIndicator = 0;
            countAnnHours = 0;
            countExemptStatus = 0;
            countDivClass = 0;
            countAnnSalary = 0;
            countPayCycle = 0;
            countDisPlanCode = 0;

            // if (fileData.length < 20000) {
            //     groupWS.commit();
            // }
            

                
            // console.log(finalData);
            // console.log(result.toString());
                
              
        }

        // console.log("SSN: " + countSSN);
        // console.log("AnnHOURS: " + countAnnHours);
        // console.log("AnnSala: " + countAnnSalary);
        // console.log("Displancode: " + countDisPlanCode);
        // console.log("DivClass: " + countDivClass);
        // console.log("Emply Status: " + countEmployStatus);
        // console.log("Exemptstatus: " + countExemptStatus);
        // console.log("FTINDICATRO: " + countFTIndicator);
        // console.log("HireDate: " + countHireDate);
        // console.log("PayCyce: " + countPayCycle);
        // console.log("workState: " + countWorkState);
        // console.log(finalData);
        // console.log(employerGroupFile);
        // console.log(masterCount);

        reportWS.commit();
        
        
        
        //console.log(finalData);


        workbook.commit();
        console.log('COMEPLETE');

        loadingIcon(false);
        

        
        
    } 


    async function createFile() {
        submitButton.style.display = "none";
        cancelButton.style.display = "none";
        const fs = require('fs');

        try {
            if (!fs.existsSync(desktopDir + 'OUT')) {
                fs.mkdirSync(desktopDir + 'OUT');
            }
            } catch (err) {
            console.error(err);
            }

        const options = {
            filename: desktopDir + 'OUT\\Blank Field Report.xlsx',
            useStyles: true,
            useSharedStrings: true
        };

        workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
        reportWS = workbook.addWorksheet("Report");
        
        var checkNumbPro = 1;
        var excelColumns = [];

        
        
        //console.log(excelColumns);
        reportWS.columns = [{header: 'Employer Group', key: 'Employer Group', width: 25}, {header: 'Total Records', key: 'Total Records', width: 15}, {header: 'SSN of Claimant', key: 'SSN of Claimant', width: 15}, {header: 'Hire Date', key: 'Hire Date', width: 15},
        {header: 'Work State', key: 'Work State', width: 15}, {header: 'Employment Status', key: 'Employment Status', width: 15}, {header: 'FT/PT Indicator', key: 'FT/PT Indicator', width: 15},
        {header: 'Annual Hours', key: 'Annual Hours', width: 15}, {header: 'Exempt Status', key: 'Exempt Status', width: 15}, {header: 'Division/Classes', key: 'Division/Classes', width: 15},
        {header: 'Annual Salary', key: 'Annual Salary', width: 15}, {header: 'Pay Cycle', key: 'Pay Cycle', width: 15}, {header: 'Disability Plan Code', key: 'Disability Plan Code', width: 15}];

        //await workbook.commit();
        // while (checkNumbPro <= numberProcessors) {
        //     var 
        // }
        

        // console.log("STARTING RAW DATA INSERT!");
        // for (let i = 0; i < rawData.length; i++) {
        //     rawDataWS.addRow(rawData[i]).commit();
        // }

        //rawDataWS.commit();
        //console.log(Math.trunc(rowCount / numberProcessors));


        
        // for (let i = 0; i < mcaFINAL.length; i++) {
        //     mcaFinalWS.addRow(mcaFINAL[i]).commit();
        // }

        // mcaFinalWS.commit();
        // await workbookMCA.commit();
        //console.log(mcaFINALString);


        //rawData = [];
        
        //loadingIcon(false);

    }

    function loadingIcon(loading) {
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
                //location.reload();
            }, 10000);
        }
    }


}
