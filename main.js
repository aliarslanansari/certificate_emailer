const electron = require("electron");
const url = require("url");
const path = require("path");
const Excel = require('exceljs');
require('electron-reload')(__dirname);
const {app,BrowserWindow,Menu,ipcMain, session} = electron;
const nodemailer = require('nodemailer');

