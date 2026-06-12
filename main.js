const { app, BrowserWindow, Menu, ipcMain, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn, execSync } = require('child_process');
const {
  checkForUpdates,
  downloadUpdate,
  formatBytes,
} = require('./updater');
const XLSX = require('xlsx');
const {
  fillAndroidExcel,
  formatFillOutput,
  isAndroidFeatureSheet,
  describeAndroidSheetIssue,
  hasAndroidPropertyMarker,
  buildReportPath,
  formatReportTimestamp,
  getProductName,
} = require('./fill-android-excel');
const {
  applyAgentFills,
  extractFillsFromAgentOutput,
} = require('./excel-preserve-fill');

const AGENT_TIMEOUT_MS = 30 * 60 * 1000;
const MAX_FILL_HISTORY = 30;
const UPDATE_CHECK_DELAY_MS = 3000;
let activeAgentChild = null;
let updateDownloadInProgress = false;

function getHistoryFilePath() {
  return path.join(app.getPath('userData'), 'fill-history.json');
}

function loadFillHistory() {
  try {
    const raw = fs.readFileSync(getHistoryFilePath(), 'utf8');
    const list = JSON.parse(raw);
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

function saveFillHistoryItem(entry) {
  const history = loadFillHistory();
  history.unshift(entry);
  if (history.length > MAX_FILL_HISTORY) {
    history.length = MAX_FILL_HISTORY;
  }
  fs.mkdirSync(path.dirname(getHistoryFilePath()), { recursive: true });
  fs.writeFileSync(getHistoryFilePath(), JSON.stringify(history, null, 2), 'utf8');
  return history;
}

function checkAdbDevices() {
  try {
    const out = execSync('adb devices', {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      timeout: 10000,
    });
    const devices = out
      .split(/\r?\n/)
      .slice(1)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => line.split(/\s+/))
      .filter((parts) => parts[1] === 'device')
      .map((parts) => parts[0]);

    return { connected: devices.length > 0, devices };
  } catch (error) {
    return { connected: false, devices: [], error: error.message };
  }
}

function sendProgress(webContents, message, partial = '') {
  if (webContents && !webContents.isDestroyed()) {
    webContents.send('agent:progress', { message, partial });
  }
}

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
  return win;
}

function getMainWindow() {
  const windows = BrowserWindow.getAllWindows();
  return windows.find((win) => !win.isDestroyed()) || null;
}

function formatReleaseNotes(notes) {
  const text = String(notes || '').trim();
  if (!text) return '是否立即下载并安装更新？';
  return text.length > 500 ? `${text.slice(0, 500)}...\n\n是否立即下载并安装更新？` : `${text}\n\n是否立即下载并安装更新？`;
}

async function runInstallerAndQuit(installerPath) {
  if (process.platform === 'win32') {
    spawn(installerPath, [], {
      detached: true,
      stdio: 'ignore',
    }).unref();
    app.quit();
    return { ok: true };
  }

  await shell.openPath(installerPath);
  return { ok: true, opened: true };
}

async function downloadAndInstallUpdate(release, parentWin) {
  if (updateDownloadInProgress) {
    return { ok: false, error: '正在下载更新，请稍候' };
  }

  updateDownloadInProgress = true;
  const win = parentWin && !parentWin.isDestroyed() ? parentWin : getMainWindow();

  try {
    if (win) {
      win.setProgressBar(2);
    }

    const installerPath = await downloadUpdate(
      release,
      path.join(app.getPath('temp'), 'tools-updates'),
      (received, total) => {
        if (!win || win.isDestroyed()) return;
        if (total > 0) {
          win.setProgressBar(received / total);
        }
      },
    );

    if (win && !win.isDestroyed()) {
      win.setProgressBar(-1);
    }

    const { response } = await dialog.showMessageBox(win, {
      type: 'info',
      title: '下载完成',
      message: '更新包已下载完成',
      detail: '安装程序即将启动，请按提示完成安装。安装完成后请重新打开应用。',
      buttons: ['立即安装', '取消'],
      defaultId: 0,
      cancelId: 1,
    });

    if (response !== 0) {
      return { ok: true, downloaded: true, installed: false, installerPath };
    }

    await runInstallerAndQuit(installerPath);
    return { ok: true, downloaded: true, installed: true, installerPath };
  } catch (error) {
    if (win && !win.isDestroyed()) {
      win.setProgressBar(-1);
    }
    await dialog.showMessageBox(win, {
      type: 'error',
      title: '更新失败',
      message: '下载或安装更新失败',
      detail: error.message,
      buttons: ['确定'],
    });
    return { ok: false, error: error.message };
  } finally {
    updateDownloadInProgress = false;
  }
}

async function promptForUpdate(updateInfo, parentWin, { silentWhenNoUpdate = false } = {}) {
  const win = parentWin && !parentWin.isDestroyed() ? parentWin : getMainWindow();

  if (!updateInfo.available) {
    if (!silentWhenNoUpdate) {
      await dialog.showMessageBox(win, {
        type: 'info',
        title: '检查更新',
        message: '当前已是最新版本',
        detail: `当前版本：v${updateInfo.currentVersion}`,
        buttons: ['确定'],
      });
    }
    return { ok: true, available: false, ...updateInfo };
  }

  const release = updateInfo.release;
  const sizeHint = release.size ? `安装包大小约 ${formatBytes(release.size)}` : '';
  const { response } = await dialog.showMessageBox(win, {
    type: 'info',
    title: '发现新版本',
    message: `当前版本 v${updateInfo.currentVersion}，最新版本 v${updateInfo.latestVersion}`,
    detail: [sizeHint, formatReleaseNotes(release.notes)].filter(Boolean).join('\n\n'),
    buttons: ['立即更新', '稍后'],
    defaultId: 0,
    cancelId: 1,
  });

  if (response !== 0) {
    return { ok: true, available: true, declined: true, ...updateInfo };
  }

  const installResult = await downloadAndInstallUpdate(release, win);
  return { ok: installResult.ok, available: true, ...updateInfo, ...installResult };
}

async function checkUpdateAndPrompt(parentWin, options = {}) {
  try {
    const updateInfo = await checkForUpdates(app.getVersion());
    return promptForUpdate(updateInfo, parentWin, options);
  } catch (error) {
    const win = parentWin && !parentWin.isDestroyed() ? parentWin : getMainWindow();
    if (!options.silentWhenNoUpdate) {
      await dialog.showMessageBox(win, {
        type: 'error',
        title: '检查更新失败',
        message: '无法从 GitHub Releases 获取版本信息',
        detail: error.message,
        buttons: ['确定'],
      });
    }
    return { ok: false, error: error.message };
  }
}

function runAgentCommand(command, options = {}) {
  const {
    webContents,
    workspace,
    timeoutMs = AGENT_TIMEOUT_MS,
    model,
    trust = false,
  } = options;

  return new Promise((resolve) => {
    const trimmed = command.trim();
    if (!trimmed) {
      resolve({ ok: false, output: '请输入命令内容。' });
      return;
    }

    const args = [
      '-p',
      '-f',
      '--output-format',
      'text',
    ];

    if (model) {
      args.push('--model', model);
    }

    if (trust) {
      args.push('--trust');
    }

    if (workspace) {
      args.push('--workspace', workspace);
    }

    args.push(trimmed);

    let child;
    try {
      child = spawn('agent', args, {
        cwd: workspace || process.cwd(),
        shell: true,
        env: process.env,
        windowsHide: true,
      });
    } catch (error) {
      resolve({
        ok: false,
        output: `无法启动 agent：${error.message}`,
      });
      return;
    }

    activeAgentChild = child;

    let stdout = '';
    let stderr = '';
    let settled = false;

    const finish = (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (activeAgentChild === child) {
        activeAgentChild = null;
      }
      resolve(result);
    };

    const timer = setTimeout(() => {
      child.kill();
      finish({
        ok: false,
        output: [stdout, stderr, '\n[超时] Agent 执行超过 30 分钟，已终止。'].filter(Boolean).join('\n').trim(),
      });
    }, timeoutMs);

    child.stdout.on('data', (data) => {
      const text = data.toString();
      stdout += text;
      sendProgress(webContents, null, text);
    });

    child.stderr.on('data', (data) => {
      const text = data.toString();
      stderr += text;
      sendProgress(webContents, null, text);
    });

    child.on('error', (error) => {
      finish({
        ok: false,
        output: `无法启动 agent：${error.message}\n请确认已安装 Cursor Agent CLI 且 agent 在 PATH 中。`,
      });
    });

    child.on('close', (code) => {
      const output = [stdout, stderr].filter(Boolean).join('\n').trim()
        || (code === 0 ? '命令已执行，无输出。' : `命令退出，代码：${code}`);

      finish({ ok: code === 0, output });
    });
  });
}

function readExcelFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheets = {};
  for (const name of workbook.SheetNames) {
    sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      defval: '',
      raw: false,
    });
  }
  return { sheetNames: workbook.SheetNames, sheets };
}

function buildOutputPath(filePath, time = new Date()) {
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);
  return path.join(dir, `${base}_filled_${formatReportTimestamp(time)}${ext}`);
}

function writeTableDataFile(tableData) {
  const dataFilePath = path.join(
    os.tmpdir(),
    `tools-excel-fill-${Date.now()}.json`,
  );
  fs.writeFileSync(dataFilePath, JSON.stringify(tableData, null, 2), 'utf8');
  return dataFilePath;
}

function buildFillPrompt(sourcePath, outputPath, dataFilePath) {
  return [
    `填写该特性表格(${outputPath})`,
    '',
    `源文件（只读）：${sourcePath}`,
    `表格数据 JSON：${dataFilePath}`,
    '',
    '要求：',
    '1. 只填写空白单元格，不修改已有非空内容',
    '2. 不要读写、创建或覆盖任何 Excel 文件',
    '3. 仅输出一个 JSON 对象，不要输出其它说明文字',
    '4. JSON 格式：{"fills":[{"sheet":"工作表名","cell":"B3","value":"内容"}]}',
    '5. cell 使用 Excel 单元格地址（如 A1、C10）；也可用 row/col（从 0 开始）',
  ].join('\n');
}

function prepareOutputCopy(sourcePath, time = new Date()) {
  const outputPath = buildOutputPath(sourcePath, time);
  fs.copyFileSync(sourcePath, outputPath);
  return outputPath;
}

function formatAgentCommand(prompt, options = {}) {
  const parts = ['agent', '-p', '-f', '--output-format', 'text'];
  if (options.model) {
    parts.push('--model', options.model);
  }
  if (options.workspace) {
    parts.push('--workspace', options.workspace);
  }
  parts.push(`"${prompt}"`);
  return parts.join(' ');
}

ipcMain.handle('agent:run', (event, command) => runAgentCommand(command, {
  webContents: event.sender,
  trust: true,
}));

ipcMain.handle('agent:cancel', () => {
  if (activeAgentChild) {
    activeAgentChild.kill();
    activeAgentChild = null;
    return { ok: true };
  }
  return { ok: false };
});

ipcMain.handle('fill-history:list', () => loadFillHistory());

ipcMain.handle('shell:openPath', async (_event, filePath) => {
  if (typeof filePath !== 'string' || !filePath.trim()) {
    return { ok: false, error: '无效路径' };
  }
  if (!fs.existsSync(filePath)) {
    return { ok: false, error: '文件不存在' };
  }
  const error = await shell.openPath(filePath);
  return error ? { ok: false, error } : { ok: true };
});

ipcMain.handle('agent:fill-excel', async (event) => {
  const webContents = event.sender;
  const win = BrowserWindow.fromWebContents(webContents);

  const adb = checkAdbDevices();
  if (!adb.connected) {
    await dialog.showMessageBox(win, {
      type: 'warning',
      title: '未连接设备',
      message: '未检测到已连接的 Android 设备',
      detail: '请先通过 USB 连接设备并开启 USB 调试，确认 adb devices 可见设备后再填写特性表。',
      buttons: ['确定'],
    });
    return { ok: false, output: '未连接 Android 设备，已取消填表。', filePath: null };
  }

  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    title: '选择要填写的 Excel 表格',
    properties: ['openFile'],
    filters: [
      { name: 'Excel 表格', extensions: ['xlsx', 'xls', 'xlsm'] },
    ],
  });

  if (canceled || !filePaths.length) {
    return { ok: false, output: '未选择文件。', filePath: null };
  }

  const filePath = filePaths[0];
  const workspace = path.dirname(filePath);

  if (path.extname(filePath).toLowerCase() === '.xls') {
    return {
      ok: false,
      output: '暂不支持 .xls 格式保留样式，请先在 Excel 中另存为 .xlsx 后再填写。',
      filePath,
    };
  }

  sendProgress(webContents, `已选择：${filePath}\n正在准备输出副本...`);

  try {
    const fillTime = new Date();
    const outputPath = prepareOutputCopy(filePath, fillTime);
    const tableData = readExcelFile(filePath);

    const productName = getProductName(new Map());
    if (isAndroidFeatureSheet(tableData, productName)) {
      const productHint = productName ? `设备型号：${productName}\n` : '';
      sendProgress(webContents, `${productHint}已识别 Android 特性验证表，正在本地验证并填写...\n输出文件：${outputPath}`);
      const fillResult = await fillAndroidExcel({
        excelPath: outputPath,
        tableData,
        fillTime,
        preserveFormat: true,
        getpropFile: path.join(os.homedir(), 'Desktop', 'ai', 'as.txt'),
        packageFile: path.join(os.homedir(), 'Desktop', '脚本', 'path.txt'),
      });
      const reportPath = fillResult.reportPath || buildReportPath(outputPath, fillTime);
      const historyEntry = {
        id: Date.now(),
        time: fillTime.toISOString(),
        filePath,
        outputPath,
        reportPath: fs.existsSync(reportPath) ? reportPath : null,
        ok: true,
      };
      saveFillHistoryItem(historyEntry);

      return {
        ok: true,
        output: `${formatFillOutput(fillResult)}\n\n源文件未修改：${filePath}`,
        filePath,
        outputPath,
        reportPath: historyEntry.reportPath,
        history: loadFillHistory(),
      };
    }

    const hasPropertySheet = tableData.sheetNames.some(
      (name) => hasAndroidPropertyMarker(tableData.sheets[name] || []),
    );
    if (hasPropertySheet) {
      const issue = describeAndroidSheetIssue(tableData, productName);
      return {
        ok: false,
        output: `已识别为 Android 配置属性表，但当前表头无法完成本地验证：\n${issue}`,
        filePath,
        outputPath,
      };
    }

    let dataFilePath = null;
    try {
      dataFilePath = writeTableDataFile({ ...tableData, sourcePath: filePath, outputPath });
      const prompt = buildFillPrompt(filePath, outputPath, dataFilePath);
      const commandPreview = formatAgentCommand(prompt, {
        model: 'auto',
        workspace,
      });

      sendProgress(
        webContents,
        `源文件：${filePath}\n`
          + `输出文件：${outputPath}\n`
          + `正在调用 Agent 分析空白单元格（不直接改 Excel）...\n`
          + `执行：${commandPreview}\n`
          + `（首次响应通常需要 30 秒～数分钟）\n\n`,
      );

      const result = await runAgentCommand(prompt, {
        webContents,
        workspace,
        model: 'auto',
        trust: true,
      });

      if (!result.ok) {
        return { ...result, filePath, outputPath, command: commandPreview };
      }

      sendProgress(webContents, `Agent 分析完成，正在写入单元格（保留格式）...\n\n${result.output}\n\n`);

      const fills = extractFillsFromAgentOutput(result.output);
      const writeResult = await applyAgentFills(outputPath, fills);

      const summary = [
        result.output,
        '',
        `已写入 ${writeResult.applied} 个空白单元格，跳过 ${writeResult.skipped} 个非空/无效单元格`,
        `源文件未修改：${filePath}`,
        `输出文件：${outputPath}`,
      ].join('\n');

      const reportPath = buildReportPath(outputPath, fillTime);
      const historyEntry = {
        id: Date.now(),
        time: fillTime.toISOString(),
        filePath,
        outputPath,
        reportPath: fs.existsSync(reportPath) ? reportPath : null,
        ok: true,
      };
      saveFillHistoryItem(historyEntry);

      return {
        ok: true,
        filePath,
        outputPath,
        reportPath: historyEntry.reportPath,
        command: commandPreview,
        output: summary,
        history: loadFillHistory(),
      };
    } finally {
      if (dataFilePath) {
        fs.unlink(dataFilePath, () => {});
      }
    }
  } catch (error) {
    return {
      ok: false,
      output: `处理 Excel 失败：${error.message}`,
      filePath,
    };
  }
});

ipcMain.handle('app:get-version', () => app.getVersion());

ipcMain.handle('app:check-update', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  return checkUpdateAndPrompt(win, { silentWhenNoUpdate: false });
});

ipcMain.handle('shell:openExternal', async (_event, url) => {
  if (typeof url === 'string' && /^https?:\/\//i.test(url)) {
    await shell.openExternal(url);
    return { ok: true };
  }
  return { ok: false };
});

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  const mainWindow = createWindow();

  if (app.isPackaged) {
    setTimeout(() => {
      checkUpdateAndPrompt(mainWindow, { silentWhenNoUpdate: true });
    }, UPDATE_CHECK_DELAY_MS);
  }

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
