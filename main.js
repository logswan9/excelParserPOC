const electron = require('electron');
const { app, BrowserWindow } = electron;
const ipcMain = electron.ipcMain;

let mainWindow;

app.on('ready', function() {
  mainWindow = new BrowserWindow({
    width: 500,
    height: 800,
    icon: 'src/images/sword.png',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    resizable: false
  });
  mainWindow.loadFile('src/pages/indexMaster.html')
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
})



let inputWindow;

ipcMain.on('get-input', function () {
  inputWindow = new BrowserWindow({
    width: 600,
    height: 400,
    icon: 'src/images/sword.png',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
    resizable: false
  });
  inputWindow.loadFile('src/pages/input.html')
});

ipcMain.on('input-gathered', function (e, year, month, adjReason) {
  mainWindow.webContents.send('input-window-close', year, month, adjReason);
  inputWindow.close();
  inputWindow = null;
  
});




