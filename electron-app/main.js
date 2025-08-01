const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const xlsx = require('xlsx');

let mainWindow;

// 🧹 GPU fallback (for older machines)
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('in-process-gpu');
app.disableHardwareAcceleration();

// 🪟 Create main window
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      enableRemoteModule: false,
      worldSafeExecuteJavaScript: true
    }
  });

  // Uncomment for debugging:
  // mainWindow.webContents.openDevTools();

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
}

app.whenReady().then(createWindow);

// 📂 Folder Picker
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  return result.canceled ? null : result.filePaths[0];
});

// 🚀 Run Python Extractor
ipcMain.handle('run-python', async (_, folderPath) => {
  return new Promise((resolve) => {
    const isDev = !app.isPackaged;
    const possiblePaths = isDev
      ? [path.join(__dirname, 'extractor.exe')]
      : [
          path.join(__dirname, 'electron-app', 'extractor.exe'),
          path.join(process.resourcesPath, 'app', 'electron-app', 'extractor.exe')
        ];

    const exePath = possiblePaths.find(p => fs.existsSync(p));
    if (!exePath) return resolve({ success: false, error: 'Extractor executable not found.' });

    execFile(exePath, [folderPath], { cwd: path.dirname(exePath), timeout: 300000 }, (error, stdout, stderr) => {
      if (error) {
        resolve({ success: false, error: (stderr || error.message).toString() });
      } else {
        resolve({ success: true, output: stdout });
      }
    });
  });
});

// 📊 Load Results Data
ipcMain.handle('load-results', async (_, folderPath) => {
  try {
    const files = fs.readdirSync(folderPath);
    let filePath = files.find(f => /^00 Summary of Tender - .+\.xlsx$/.test(f));

    filePath = filePath
      ? path.join(folderPath, filePath)
      : path.join(folderPath, '00 Extracted_Answers.xlsx');

    if (!fs.existsSync(filePath)) {
      return { success: false, error: 'No output Excel file found.' };
    }

    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets['output'];
    const json = xlsx.utils.sheet_to_json(sheet);

    return { success: true, data: json, file: path.basename(filePath) };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// 📄 Show Results Page
ipcMain.handle('show-results-page', async (_, folderPath) => {
  const resultsPath = path.join(__dirname, 'results.html');
  if (!fs.existsSync(resultsPath)) {
    console.error('❌ results.html not found:', resultsPath);
    return;
  }

  if (mainWindow) {
    try {
      await mainWindow.loadFile(resultsPath);
      console.log('✅ Loaded results.html');
      mainWindow.webContents.send('set-folder-path', folderPath);
    } catch (err) {
      console.error('❌ Failed to load results page:', err);
    }
  }
});

// 🏠 Go Home
ipcMain.on('show-home-page', () => {
  const homePath = path.join(__dirname, 'index.html');

  if (mainWindow) {
    mainWindow.loadFile(homePath)
      .then(() => console.log('✅ Loaded index.html'))
      .catch(err => console.error('❌ Failed to load home page:', err));
  } else {
    console.error('❌ mainWindow not available');
  }
});
