const { app, BrowserWindow, Menu, ipcMain, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');

function createWindow() {
  const win = new BrowserWindow({
    width: 1100,
    height: 720,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http://') || url.startsWith('https://')) {
      shell.openExternal(url);
    }
    return { action: 'deny' };
  });

  win.loadFile('index.html');
}

function runAgentCommand(command) {
  return new Promise((resolve) => {
    const trimmed = command.trim();
    if (!trimmed) {
      resolve({ ok: false, output: '请输入命令内容。' });
      return;
    }

    const args = [
      '-p',
      '--trust',
      '--output-format',
      'text',
      trimmed,
    ];

    const child = spawn('agent', args, {
      cwd: process.cwd(),
      shell: true,
      env: process.env,
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    child.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    child.on('error', (error) => {
      resolve({
        ok: false,
        output: `无法启动 agent：${error.message}\n请确认已安装 Cursor Agent CLI 且 agent 在 PATH 中。`,
      });
    });

    child.on('close', (code) => {
      const output = [stdout, stderr].filter(Boolean).join('\n').trim()
        || (code === 0 ? '命令已执行，无输出。' : `命令退出，代码：${code}`);

      resolve({ ok: code === 0, output });
    });
  });
}

ipcMain.handle('agent:run', (_event, command) => runAgentCommand(command));

ipcMain.handle('shell:openExternal', async (_event, url) => {
  if (typeof url === 'string' && /^https?:\/\//i.test(url)) {
    await shell.openExternal(url);
    return { ok: true };
  }
  return { ok: false };
});

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
