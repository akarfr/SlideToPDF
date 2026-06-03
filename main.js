const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { exec } = require('child_process');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 480, // slightly taller to accommodate options
    resizable: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.setMenu(null);
  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// IPC: Select files dialog
ipcMain.on('select-files', (event) => {
  dialog.showOpenDialog(mainWindow, {
    title: 'Select PowerPoint Files',
    filters: [
      { name: 'PowerPoint Presentations', extensions: ['pptx'] }
    ],
    properties: ['openFile', 'multiSelections']
  }).then(result => {
    if (!result.canceled) {
      event.sender.send('selected-files-reply', result.filePaths);
    }
  }).catch(err => {
    event.sender.send('conversion-result', `Error choosing files: ${err.message}`);
  });
});

// IPC: Select output folder dialog
ipcMain.on('select-output-dir', (event) => {
  dialog.showOpenDialog(mainWindow, {
    title: 'Select Output Destination Folder',
    properties: ['openDirectory']
  }).then(result => {
    if (!result.canceled && result.filePaths.length > 0) {
      event.sender.send('selected-output-dir-reply', result.filePaths[0]);
    }
  }).catch(err => {
    event.sender.send('conversion-result', `Error choosing output folder: ${err.message}`);
  });
});

// IPC: Convert files
ipcMain.on('convert-files', (event, { files, outputDir }) => {
  const platform = os.platform();

  if (!files || files.length === 0) {
    event.sender.send('conversion-result', 'Error: No files selected.');
    return;
  }

  // If a single folder path was passed (e.g. from drag and drop of a folder)
  let resolvedFiles = [];
  if (files.length === 1) {
    try {
      const stats = fs.statSync(files[0]);
      if (stats.isDirectory()) {
        const folder = files[0];
        const dirFiles = fs.readdirSync(folder);
        resolvedFiles = dirFiles
          .filter(f => f.toLowerCase().endsWith('.pptx'))
          .map(f => path.join(folder, f));
      } else {
        resolvedFiles = [files[0]];
      }
    } catch (err) {
      event.sender.send('conversion-result', `Error reading drop path: ${err.message}`);
      return;
    }
  } else {
    resolvedFiles = files;
  }

  if (resolvedFiles.length === 0) {
    event.sender.send('conversion-result', 'Error: No PowerPoint (.pptx) files found.');
    return;
  }

  // Check if custom outputDir exists
  if (outputDir) {
    try {
      if (!fs.existsSync(outputDir)) {
        event.sender.send('conversion-result', 'Error: Selected output directory does not exist.');
        return;
      }
    } catch (err) {
      event.sender.send('conversion-result', `Error verifying output directory: ${err.message}`);
      return;
    }
  }

  if (platform === 'win32') {
    // Windows Track: PowerPoint COM Interface via PowerShell Script File
    const escapedFiles = resolvedFiles.map(f => `'${f.replace(/'/g, "''")}'`).join(', ');
    const escapedOutputDir = outputDir ? `'${outputDir.replace(/'/g, "''")}'` : "''";

    const psCommand = `
$files = @(${escapedFiles})
$outputDir = ${escapedOutputDir}
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.DisplayAlerts = 1
} catch {
    Write-Error "Failed to instantiate PowerPoint COM interface. Please ensure Microsoft PowerPoint is installed."
    exit 1
}
$total = $files.Count
$current = 1
foreach ($file in $files) {
    if (-not (Test-Path $file)) { 
        Write-Error "File not found: $file"
        continue 
    }
    $item = Get-Item $file
    Write-Output "PROGRESS:Converting ($current/$total): $($item.Name)"
    if ($outputDir -ne '') {
        $PDFPath = Join-Path $outputDir ($item.BaseName + '.pdf')
    } else {
        $PDFPath = Join-Path $item.DirectoryName ($item.BaseName + '.pdf')
    }
    $PDFFullPath = [System.IO.Path]::GetFullPath($PDFPath)
    try {
        $presentation = $ppt.Presentations.Open($item.FullName, 1, 0, 0)
        $presentation.SaveAs($PDFFullPath, 32)
        $presentation.Close()
    } catch {
        Write-Error "Error converting file $($item.Name): $($_.Exception.Message)"
        $ppt.Quit()
        exit 1
    }
    $current++
}
try {
    $ppt.Quit()
} catch {}
    `.trim();

    // Write to a temporary .ps1 file to avoid command-line length and escaping issues
    const tempDir = app.getPath('temp');
    const tempScriptPath = path.join(tempDir, `antigravity-convert-${Date.now()}.ps1`);

    try {
      fs.writeFileSync(tempScriptPath, psCommand, 'utf8');
    } catch (writeErr) {
      event.sender.send('conversion-result', `Error: Failed to write temp execution script: ${writeErr.message}`);
      return;
    }

    const { spawn } = require('child_process');
    const child = spawn('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', tempScriptPath]);

    let stderrData = '';

    child.stdout.on('data', (data) => {
      const output = data.toString();
      const match = output.match(/PROGRESS:(.*)/);
      if (match) {
        event.sender.send('conversion-result', match[1].trim());
      }
    });

    child.stderr.on('data', (data) => {
      stderrData += data.toString();
    });

    child.on('close', (code) => {
      // Clean up the temporary file
      try {
        if (fs.existsSync(tempScriptPath)) {
          fs.unlinkSync(tempScriptPath);
        }
      } catch (cleanupErr) {
        console.error('Failed to clean up temp script:', cleanupErr);
      }

      if (code !== 0) {
        const errorMsg = stderrData.trim() || `Script exited with code ${code}`;
        event.sender.send('conversion-result', `Error: ${errorMsg}`);
      } else {
        event.sender.send('conversion-result', 'Success! All files converted.');
      }
    });

  } else if (platform === 'darwin') {
    // macOS Track: headless LibreOffice
    // Run conversions sequentially to handle custom output directories and individual files cleanly
    let index = 0;
    
    function convertNext() {
      if (index >= resolvedFiles.length) {
        event.sender.send('conversion-result', 'Success! All files converted.');
        return;
      }

      const file = resolvedFiles[index];
      const fileDir = path.dirname(file);
      const outDir = outputDir ? outputDir : fileDir;

      const escapedFile = file.replace(/'/g, "'\\''");
      const escapedOutDir = outDir.replace(/'/g, "'\\''");

      const cmd = `soffice --headless --convert-to pdf --outdir '${escapedOutDir}' '${escapedFile}'`;

      exec(cmd, (error, stdout, stderr) => {
        if (error) {
          const errorMsg = stderr ? stderr.trim() : error.message;
          event.sender.send('conversion-result', `Error converting ${path.basename(file)}: ${errorMsg}`);
        } else {
          index++;
          convertNext();
        }
      });
    }

    convertNext();

  } else {
    event.sender.send('conversion-result', `Error: Platform "${platform}" is not supported.`);
  }
});
