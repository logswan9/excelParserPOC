const ExcelJS = require('exceljs');
const homeDir = require('os').homedir(); 
const desktopDir = `${homeDir}\\Desktop\\`;
const electron = require('electron');
const ipcRen = electron.ipcRenderer;
const fs = require('fs');
const { stringify } = require('querystring');
import {loadKean} from "../FP/keanToClean.js";
import {loadLost} from "../FP/lostPolicy.js"

var filePath = "";
var fileName = "";
var submitBtn = document.getElementById('lostPolicy').addEventListener("click", startLostPolicy); // Set listener on Lost Policy Button
var submitBtn2 = document.getElementById('keanClean').addEventListener("click", startKeanClean);

function startKeanClean() {
    loadKean();
}

function startLostPolicy() {
    loadLost();
}