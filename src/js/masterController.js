const ExcelJS = require('exceljs');
const homeDir = require('os').homedir(); 
const desktopDir = `${homeDir}\\Desktop\\`;
const electron = require('electron');
const ipcRen = electron.ipcRenderer;
const fs = require('fs');
const { stringify } = require('querystring');
import {loadKean} from "../js/FP/keanToClean.js";
import {loadLost} from "../js/FP/lostPolicy.js"
import {loadRetro} from "../js/retrotermMCA.js"

var submitBtn = document.getElementById('lostPolicy').addEventListener("click", startLostPolicy); // Set listener on Lost Policy Button
var submitBtn2 = document.getElementById('keanClean').addEventListener("click", startKeanClean);
var submitBtn3 = document.getElementById('retroterm').addEventListener("click", startRetroterm);

function startKeanClean() {
    loadKean();
}

function startLostPolicy() {
    loadLost();
}

function startRetroterm() {
    loadRetro();
}