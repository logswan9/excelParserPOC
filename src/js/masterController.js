const ExcelJS = require('exceljs');
const homeDir = require('os').homedir(); 
const desktopDir = `${homeDir}\\Desktop\\`;
const electron = require('electron');
const ipcRen = electron.ipcRenderer;
const fs = require('fs');
const { stringify } = require('querystring');
import {loadKean} from "../js/FP/keanToClean.js";
import {loadLost} from "../js/FP/lostPolicy.js";
import {loadRetro} from "../js/retrotermMCA.js";
import {loadCheck} from "../js/checkRun.js";
import {loadBlankCheck} from "../js/blankFieldsSearch.js";
import {loadEmployeeChanges} from "../js/employeeChanges.js";

var submitBtn = document.getElementById('lostPolicy').addEventListener("click", startLostPolicy); // Set listener on Lost Policy Button
var submitBtn2 = document.getElementById('keanClean').addEventListener("click", startKeanClean);
var submitBtn3 = document.getElementById('retroterm').addEventListener("click", startRetroterm);
var submitBtn4 = document.getElementById('checkRun').addEventListener("click", startCheckRun);
var submitBtn5 = document.getElementById('blankField').addEventListener("click", startBlankFieldSearch);
var submitBtn6 = document.getElementById('employeeChanges').addEventListener("click", startEmployeeChanges);

function startKeanClean() {
    loadKean();
}

function startLostPolicy() {
    loadLost();
}

function startRetroterm() {
    loadRetro();
}

function startCheckRun() {
    loadCheck();
}

function startBlankFieldSearch() {
    loadBlankCheck();
}

function startEmployeeChanges() {
    loadEmployeeChanges();
}