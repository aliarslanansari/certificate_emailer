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
//Catch item:add
ipcMain.on('send_email',function(e,item){
    var ws;
    //console.log(item);
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(item.excel_path)
    .then(function(){
        ws = workbook.getWorksheet('Sheet1');
        var cell = ws.getCell('A1').value;
        //console.log(cell);
        ws.eachRow(function(row, rowNumber) {
            if(rowNumber==1){
                return;
            }
            console.log(row.values[item.emailHeader]);
            sendEmail(item.host, item.port, item.email, item.pass, item.subject,row.values[item.emailHeader],item.text,rowNumber);
        })
    })
    .catch((err)=>{console.log(err)});
});

const mainMenuTemplate = [{
    label:'File',
    submenu:[
        {   
            label: 'Open Excel',
            accelerator: process.platform == 'darwin' ? 'Command+F':'Ctrl+F',
            click(){
                mainWindow.webContents.send('openexcel');
            }
        },
        {   
            label: 'Quit',
            accelerator: process.platform == 'darwin' ? 'Command+Q':'Ctrl+Q',
            click(){
                app.quit();
            }
        },
    ]
}];

//if mac, add empty object to menu
if(process.platform === "darwin"){
    mainMenuTemplate.unshift({});
}

//add developer tools item if not in production
if(process.env.NODE_ENV !== "production"){
    mainMenuTemplate.push({
        label:'Developer Tools',
        submenu:[
            {
                label: 'Toggle DevTools',
                accelerator: process.platform == 'darwin' ? 'Command+I':'Ctrl+I',
                click(item,focusedWindow){
                    focusedWindow.toggleDevTools();
                }
            },
            {
                role:'reload'
            }
        ]
    })
}
