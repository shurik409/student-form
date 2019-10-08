const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const path = require('path');
const rimraf = require("rimraf");
const fs = require('fs');

const url = require('url');
const creds = require('./secret2.json');
const faculties = new Map();

const { app, BrowserWindow } = require('electron');

var mainWindow = null;

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('ready', function () {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true
        }
    });
    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file',
        slashes: true
    }))

    // mainWindow.webContents.openDevTools();

    mainWindow.on('closed', function () {
        mainWindow = null;
    });
});