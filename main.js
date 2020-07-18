const electron = require("electron");
const url = require("url");
const path = require("path");
const Excel = require('exceljs');
require('electron-reload')(__dirname);
const {app,BrowserWindow,Menu,ipcMain, session} = electron;
const nodemailer = require('nodemailer');

let mainWindow;
let htmlPreviewWindow;
//Set Environment
process.env.NODE_ENV = 'development';

//listen for app to be ready

app.on('ready',function(){
    // create new window
    mainWindow = new BrowserWindow({
        frame:false,
        height:800,
        width:800,
        minHeight:800,
        minWidth:600,
        webPreferences: {
            nodeIntegration: true
        }});
    // load html into window
    mainWindow.loadURL(url.format({
        pathname:path.join(__dirname,"mainWindow.html"),
        protocol:"file",
        slashes:true
    }));
    mainWindow.maximize()
    //Quit app when closed
    mainWindow.on('closed', function(){
        app.quit();
    })

    //Build Menu from template
    const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
    //Insert menu
    Menu.setApplicationMenu(mainMenu);
    // Menu.setApplicationMenu(null);
    // mainWindow.setMenu(mainMenu)
});
