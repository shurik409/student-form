const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const mkdirp = require('mkdirp');
const path = require('path');
const rimraf = require("rimraf");
const fs = require('fs');

const url = require('url');
const { app, BrowserWindow } = require('electron');

const creds = require('./secret2.json');
const faculties = new Map();

let win;

function createWindow() {
    win = new BrowserWindow({
        width: 700,
        height: 500,
        icon: __dirname + '/img/icon.jpg'
    })

    win.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file',
        slashes: true
    }))


    win.webContents.openDevTools();

    win.on('closed', () => {
        win = null;
    });
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
    app.quit();
})

