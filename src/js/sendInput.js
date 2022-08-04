const ExcelJS = require('exceljs');
const homeDir = require('os').homedir(); 
const desktopDir = `${homeDir}\\Desktop\\`;
const electron = require('electron');
const ipcRen = electron.ipcRenderer;
const fs = require('fs');




var submitBtn = document.getElementById('submitBtn').addEventListener("click", sendInput);
var year = ""

async function sendInput() {
    year = document.getElementById('year').value;
    month = document.getElementById('month').value;
    adjReason = document.getElementById('adjReason').value;
    
    ipcRen.send('input-gathered', year, month, adjReason);

}