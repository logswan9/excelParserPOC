const electron = require('electron');
const { app, BrowserWindow } = electron;
const ipc = electron.ipcMain;

let mainWindow;

app.on('ready', function() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 1000,
    icon: 'src/images/cat.png',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    resizable: false
  });
  mainWindow.loadFile('index.html')
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
})



let inputWindow;

ipc.on('get-input', function () {
  inputWindow = new BrowserWindow({
    width: 600,
    height: 400,
    icon: 'src/images/cat.png',
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    resizable: false
  });
  inputWindow.loadFile('test.html')
});


