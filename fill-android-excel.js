#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const XLSX = require('xlsx');

const TEMP_JSON = process.env.TEMP_JSON
  || 'C:\\Users\\weikang.yang\\AppData\\Local\\Temp\\tools-excel-fill-1781096383855.json';
const GETPROP_FILE = process.env.GETPROP_FILE
  || 'C:\\Users\\weikang.yang\\Desktop\\ai\\as.txt';
const PACKAGE_FILE = process.env.PACKAGE_FILE
  || 'C:\\Users\\weikang.yang\\Desktop\\脚本\\path.txt';
const SEARCH_ROOTS = (process.env.SEARCH_ROOTS
  || 'C:\\Users\\weikang.yang\\Desktop;C:\\Users\\weikang.yang\\Downloads;C:\\Users\\weikang.yang\\Documents;D:\\project')
  .split(';')
  .map((s) => s.trim())
  .filter(Boolean);
const SEARCH_DEPTH = Number(process.env.SEARCH_DEPTH || 6);

const ANDROID_PROP_MARKER = '配置属性名';

function isBlank(v) {
  return v === undefined || v === null || String(v).trim() === '';
}

function norm(v) {
  return String(v ?? '').trim();
}

function headerEquals(a, b) {
  const left = norm(a).toLowerCase();
  const right = norm(b).toLowerCase();
  return left && right && left === right;
}

function findHeaderIndex(header, name) {
  const idx = header.findIndex((cell) => headerEquals(cell, name));
  return idx >= 0 ? idx : -1;
}

function findHeaderIndexAny(header, names) {
  for (const name of names) {
    const idx = findHeaderIndex(header, name);
    if (idx >= 0) return idx;
  }
  return -1;
}

function scoreProductHeaderMatch(headerCell, productName) {
  const header = norm(headerCell);
  const product = norm(productName);
  if (!header || !product) return -1;
  if (headerEquals(header, product)) return 1000 + header.length;
  const headerLower = header.toLowerCase();
  const productLower = product.toLowerCase();
  if (headerLower.includes(productLower)) return 800 + header.length;
  const minHeaderLen = Math.max(4, Math.floor(product.length * 0.6));
  if (productLower.includes(headerLower) && header.length >= minHeaderLen) {
    return 400 + header.length;
  }
  return -1;
}

function resolveSupportColumn(header, productName) {
  if (productName) {
    let bestIdx = -1;
    let bestScore = -1;
    let bestHeader = '';
    header.forEach((cell, idx) => {
      const score = scoreProductHeaderMatch(cell, productName);
      if (score > bestScore) {
        bestScore = score;
        bestIdx = idx;
        bestHeader = norm(cell);
      }
    });
    if (bestIdx >= 0) {
      return { idx: bestIdx, source: 'product', header: bestHeader };
    }
  }

  const defaultIdx = findHeaderIndex(header, '是否支持');
  if (defaultIdx >= 0) {
    return { idx: defaultIdx, source: 'default', header: norm(header[defaultIdx]) };
  }

  const fallbackIdx = header.findIndex((cell) => {
    const text = norm(cell);
    return text.includes('支持') && text.includes('不支持');
  });
  if (fallbackIdx >= 0) {
    return { idx: fallbackIdx, source: 'fallback', header: norm(header[fallbackIdx]) };
  }

  return { idx: -1, source: 'none', header: '' };
}

function findFirstHeaderIndex(header, name) {
  for (let i = 0; i < header.length; i += 1) {
    if (headerEquals(header[i], name)) return i;
  }
  return -1;
}

function buildColumnMap(headerRow, productName = '') {
  const header = headerRow.map(norm);
  const support = resolveSupportColumn(header, productName);
  const stepIdx = findHeaderIndex(header, '验证步骤');
  const remarkIdx = findFirstHeaderIndex(header, '备注');
  const resolvedStepIdx = stepIdx >= 0 ? stepIdx : remarkIdx;
  const resultIdx = findHeaderIndexAny(header, ['验证结果', '研发自检']);

  return {
    NAME: findHeaderIndexAny(header, ['特性名称', '四级特性名称', '特性编码']),
    SUPPORT: support.idx,
    RESULT: resultIdx,
    STEP: resolvedStepIdx,
    PROP: findHeaderIndex(header, '配置属性名'),
    SUPPORTED: findHeaderIndex(header, '特性支持的属性值'),
    UNSUPPORTED: findHeaderIndex(header, '特性不支持的属性值'),
    PACKAGES: findHeaderIndex(header, '关联应用包名'),
    supportSource: support.source,
    supportHeader: support.header,
    stepHeader: resolvedStepIdx >= 0 ? norm(header[resolvedStepIdx]) : '',
    resultHeader: resultIdx >= 0 ? norm(header[resultIdx]) : '',
    productName: norm(productName),
    appendedResult: false,
  };
}

function ensureResultColumn(rows, col) {
  if (!rows.length || col.RESULT >= 0) return col;
  const header = rows[0];
  const newIdx = header.length;
  header.push('验证结果');
  col.RESULT = newIdx;
  col.resultHeader = '验证结果';
  col.appendedResult = true;
  return col;
}

function columnMapValid(col) {
  return col.SUPPORT >= 0
    && col.RESULT >= 0
    && col.STEP >= 0
    && col.PROP >= 0;
}

function hasAndroidPropertyMarker(rows) {
  if (!rows.length) return false;
  return findHeaderIndex(rows[0].map(norm), ANDROID_PROP_MARKER) >= 0;
}

function getFeatureName(row, col) {
  if (col.NAME >= 0) return norm(row[col.NAME]);
  return '';
}

function minRowWidth(col) {
  return Math.max(
    col.NAME,
    col.SUPPORT,
    col.RESULT,
    col.STEP,
    col.PROP,
    col.SUPPORTED,
    col.UNSUPPORTED,
    col.PACKAGES,
  ) + 1;
}

function getProductName(props) {
  const fromMap = norm(props.get('ro.product.name'));
  if (fromMap) return fromMap;

  try {
    const out = execSync('adb shell getprop ro.product.name', {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      timeout: 30000,
    }).trim();
    const m = out.match(/^\[([^\]]+)\]:\s*\[(.*)\]\s*$/);
    if (m) return norm(m[2]);
    return norm(out);
  } catch {
    return '';
  }
}

function parseGetpropFile(filePath) {
  const map = new Map();
  if (!fs.existsSync(filePath)) return map;
  const text = fs.readFileSync(filePath, 'utf8');
  for (const line of text.split(/\r?\n/)) {
    const m = line.match(/^\[([^\]]+)\]:\s*\[(.*)\]\s*$/);
    if (m) map.set(m[1], m[2]);
  }
  return map;
}

function tryAdbGetprop() {
  try {
    const out = execSync('adb shell getprop', {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      timeout: 120000,
    });
    const map = new Map();
    for (const line of out.split(/\r?\n/)) {
      const m = line.match(/^\[([^\]]+)\]:\s*\[(.*)\]\s*$/);
      if (m) map.set(m[1], m[2]);
    }
    return map;
  } catch {
    return new Map();
  }
}

function parsePackagePathLine(line) {
  const trimmed = norm(line);
  if (!trimmed) return null;

  const listFormat = trimmed.match(/^(package:\S+)=(.+)$/);
  if (listFormat) {
    return { pkg: listFormat[2], path: listFormat[1] };
  }

  const assignFormat = trimmed.match(/^(com\.[a-zA-Z0-9_.]+)\s*[=:]\s*(package:\S+)$/);
  if (assignFormat) {
    return { pkg: assignFormat[1], path: assignFormat[2] };
  }

  if (/^com\.[a-zA-Z0-9_.]+$/.test(trimmed)) {
    return { pkg: trimmed, path: null };
  }

  return null;
}

function loadAllPackagePathsFromAdb() {
  const map = new Map();
  try {
    const out = execSync('adb shell pm list packages -f', {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      timeout: 120000,
    });
    for (const line of out.split(/\r?\n/)) {
      const parsed = parsePackagePathLine(line);
      if (parsed?.pkg && parsed.path) {
        map.set(parsed.pkg, parsed.path);
      }
    }
  } catch {
    /* ignore */
  }
  return map;
}

function tryAdbPackagePath(pkg, adbPathCache) {
  if (adbPathCache?.has(pkg)) {
    return adbPathCache.get(pkg);
  }

  try {
    const out = execSync(`adb shell pm path ${pkg}`, {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      timeout: 30000,
    }).trim();

    const lines = out.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    const pathLine = lines.find((line) => /^package:/i.test(line));
    if (pathLine && !/error|not found|Exception/i.test(pathLine)) {
      return pathLine;
    }
  } catch {
    /* ignore */
  }

  return null;
}

function resolvePackagePaths(packages, installedSet, adbPathCache, pathTxtMap) {
  const paths = [];
  for (const pkg of packages) {
    let adbPath = tryAdbPackagePath(pkg, adbPathCache);
    if (!adbPath && pathTxtMap?.has(pkg)) {
      adbPath = pathTxtMap.get(pkg);
    }

    paths.push({
      pkg,
      path: adbPath,
      inPathTxt: installedSet.has(pkg),
    });
  }
  return paths;
}

function splitValues(raw) {
  const s = norm(raw);
  if (!s) return [];
  return s.split(/[,，;；|]/).map((x) => x.trim()).filter(Boolean);
}

function valuesMatch(actual, expected) {
  const a = norm(actual);
  const e = norm(expected);
  if (!e) return false;
  if (a === e) return true;
  if (a.toLowerCase() === e.toLowerCase()) return true;
  return false;
}

function valueInList(actual, listRaw) {
  const list = splitValues(listRaw);
  return list.some((item) => valuesMatch(actual, item));
}

function isValidPropName(name) {
  const n = norm(name);
  if (!n) return false;
  if (/^[\u4e00-\u9fff·]+/.test(n) && !n.includes('.')) return false;
  return /^[a-zA-Z][a-zA-Z0-9_.]*$/.test(n) || n.includes('.');
}

function extractPackages(text) {
  const matches = norm(text).match(/com\.[a-zA-Z0-9_.]+/g);
  return matches ? [...new Set(matches)] : [];
}

function classifyPropValue(actual, supportedRaw, unsupportedRaw) {
  if (actual === null || actual === undefined) return 'missing';
  if (valueInList(actual, supportedRaw)) return 'supported';
  if (valueInList(actual, unsupportedRaw)) return 'unsupported';
  return 'other';
}

function hasConfigProp(row, col) {
  return col.PROP >= 0 && isValidPropName(norm(row[col.PROP]));
}

function hasPackages(row, col) {
  return col.PACKAGES >= 0 && extractPackages(row[col.PACKAGES]).length > 0;
}

function validateConfig(row, props, col) {
  const propName = norm(row[col.PROP]);
  const supportedRaw = col.SUPPORTED >= 0 ? row[col.SUPPORTED] : '';
  const unsupportedRaw = col.UNSUPPORTED >= 0 ? row[col.UNSUPPORTED] : '';
  const expectSupport = norm(row[col.SUPPORT]);

  if (!isValidPropName(propName)) {
    return { status: 'skip', info: '', propName: '', propValue: null, missingProp: false };
  }

  const propValue = props.has(propName) ? props.get(propName) : null;
  const hasSupportedList = !isBlank(supportedRaw);
  const hasUnsupportedList = !isBlank(unsupportedRaw);
  const classification = propValue === null
    ? 'missing'
    : classifyPropValue(propValue, supportedRaw, unsupportedRaw);

  if (classification === 'missing') {
    if (expectSupport === '支持') {
      return {
        status: 'fail',
        info: `未获取到属性值: ${propName}，需添加属性值为: ${norm(supportedRaw)}`,
        propName,
        propValue: null,
        missingProp: true,
      };
    }
    if (expectSupport === '不支持') {
      return {
        status: 'fail',
        info: `未获取到属性值: ${propName}，需添加属性值为: ${norm(unsupportedRaw)}`,
        propName,
        propValue: null,
        missingProp: true,
      };
    }
    return { status: 'NA', info: `未获取到属性值: ${propName}`, propName, propValue: null, missingProp: true };
  }

  if (!hasSupportedList && !hasUnsupportedList) {
    return {
      status: 'NA',
      info: `获取配置属性结果: ${propName}=${propValue}（无支持/不支持对照值）`,
      propName,
      propValue,
      missingProp: false,
    };
  }

  if (classification === 'supported') {
    if (expectSupport === '支持') {
      return {
        status: 'pass',
        info: `获取配置属性结果: ${propName}=${propValue}`,
        propName,
        propValue,
        missingProp: false,
      };
    }
    return {
      status: 'fail',
      info: `获取配置属性结果: ${propName}=${propValue}，需要修改为: ${norm(unsupportedRaw)}`,
      propName,
      propValue,
      missingProp: false,
    };
  }

  if (classification === 'unsupported') {
    if (expectSupport === '不支持') {
      return {
        status: 'pass',
        info: `获取配置属性结果: ${propName}=${propValue}`,
        propName,
        propValue,
        missingProp: false,
      };
    }
    return {
      status: 'fail',
      info: `获取配置属性结果: ${propName}=${propValue}，需要修改为: ${norm(supportedRaw)}`,
      propName,
      propValue,
      missingProp: false,
    };
  }

  return {
    status: 'NA',
    info: `获取配置属性结果: ${propName}=${propValue}（值不在支持/不支持列表中）`,
    propName,
    propValue,
    missingProp: false,
  };
}

function formatPackagePathInfo(resolved) {
  return resolved.map((r) => {
    if (r.path) return `${r.pkg}=${r.path}`;
    return `${r.pkg}=未获取到安装路径`;
  }).join('；');
}

function packagePathMessage(resolved) {
  return `应用包名路径: ${formatPackagePathInfo(resolved)}`;
}

function validatePackages(row, installedSet, adbPathCache, pathTxtMap, col) {
  const packages = col.PACKAGES >= 0 ? extractPackages(row[col.PACKAGES]) : [];
  const expectSupport = norm(row[col.SUPPORT]);

  if (packages.length === 0) {
    return { status: 'skip', info: '', paths: [] };
  }

  const resolved = resolvePackagePaths(packages, installedSet, adbPathCache, pathTxtMap);
  const pathMessage = packagePathMessage(resolved);
  const allHaveRealPath = resolved.every((r) => r.path);
  const noneInstalled = resolved.every((r) => !r.path && !r.inPathTxt);
  const found = resolved.filter((r) => r.path);

  if (expectSupport === '支持') {
    return {
      status: allHaveRealPath ? 'pass' : 'fail',
      info: pathMessage,
      paths: resolved,
    };
  }

  if (noneInstalled || (!allHaveRealPath && found.length === 0)) {
    return { status: 'pass', info: pathMessage, paths: resolved };
  }
  if (found.length > 0) {
    return { status: 'fail', info: pathMessage, paths: resolved };
  }
  return { status: 'NA', info: pathMessage, paths: resolved };
}

function effectiveValidationResult(result) {
  if (result.status === 'skip') {
    return { status: 'pass', info: '' };
  }
  return result;
}

function combineValidation(configResult, pkgResult) {
  if (configResult.status === 'skip' && pkgResult.status === 'skip') {
    return { result: 'NA', step: 'NA' };
  }

  const config = effectiveValidationResult(configResult);
  const pkg = effectiveValidationResult(pkgResult);
  const parts = [];
  if (config.info) parts.push(config.info);
  if (pkg.info) parts.push(pkg.info);

  if (config.status === 'NA' || pkg.status === 'NA') {
    return { result: 'NA', step: parts.join('；') || 'NA' };
  }
  if (config.status === 'pass' && pkg.status === 'pass') {
    return { result: 'pass', step: parts.join('；') };
  }
  return { result: 'fail', step: parts.join('；') };
}

function sheetHasHeaders(rows, productName = '') {
  if (!hasAndroidPropertyMarker(rows)) return false;
  const col = ensureResultColumn(rows, buildColumnMap(rows[0], productName));
  return columnMapValid(col);
}

function loadWorkbookFromJson(jsonPath) {
  const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
  const wb = XLSX.utils.book_new();
  for (const name of data.sheetNames) {
    const ws = XLSX.utils.aoa_to_sheet(data.sheets[name]);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return { wb, source: 'json', jsonPath };
}

function loadWorkbookFromExcel(excelPath) {
  return { wb: XLSX.readFile(excelPath), source: 'excel', excelPath };
}

function findMatchingExcel() {
  const exts = new Set(['.xlsx', '.xls', '.xlsm']);
  const hits = [];

  function walk(dir, depth) {
    if (depth > SEARCH_DEPTH) return;
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const ent of entries) {
      const full = path.join(dir, ent.name);
      if (ent.isDirectory()) {
        if (['node_modules', '.git', '$RECYCLE.BIN'].includes(ent.name)) continue;
        walk(full, depth + 1);
        continue;
      }
      const ext = path.extname(ent.name).toLowerCase();
      if (!exts.has(ext)) continue;
      try {
        const wb = XLSX.readFile(full, { sheetRows: 2 });
        const sheetName = wb.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
          header: 1,
          defval: '',
          raw: false,
        });
        if (sheetHasHeaders(rows)) {
          hits.push({ path: full, mtime: fs.statSync(full).mtimeMs });
        }
      } catch {
        /* skip unreadable */
      }
    }
  }

  for (const root of SEARCH_ROOTS) walk(root, 0);
  hits.sort((a, b) => b.mtime - a.mtime);
  return hits[0]?.path || null;
}

function aoaFromWorkbook(wb) {
  const sheets = {};
  for (const name of wb.SheetNames) {
    sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], {
      header: 1,
      defval: '',
      raw: false,
    });
  }
  return sheets;
}

function writeWorkbook(wb, outPath) {
  XLSX.writeFile(wb, outPath);
}

function updateSheetRows(rows, props, installedSet, summary, sheetName, updates, adbPathCache, pathTxtMap, col) {
  if (!rows.length || !columnMapValid(col)) return 0;

  let updated = 0;
  const rowWidth = minRowWidth(col);
  if (col.appendedResult && updates) {
    updates.push({ sheet: sheetName, row: 0, col: col.RESULT, value: col.resultHeader });
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    while (row.length < rowWidth) row.push('');

    if (!hasConfigProp(row, col) && !hasPackages(row, col)) continue;
    if (!isBlank(row[col.RESULT]) && !isBlank(row[col.STEP])) continue;

    const configResult = validateConfig(row, props, col);
    const pkgResult = validatePackages(row, installedSet, adbPathCache, pathTxtMap, col);
    const final = combineValidation(configResult, pkgResult);

    if (isBlank(row[col.RESULT])) {
      row[col.RESULT] = final.result;
      if (updates) {
        updates.push({ sheet: sheetName, row: i, col: col.RESULT, value: final.result });
      }
    }
    if (isBlank(row[col.STEP])) {
      row[col.STEP] = final.step;
      if (updates) {
        updates.push({ sheet: sheetName, row: i, col: col.STEP, value: final.step });
      }
    }

    summary.total += 1;
    summary[final.result] = (summary[final.result] || 0) + 1;
    if (configResult.missingProp) summary.missingProp += 1;

    if (final.result === 'fail') {
      if (configResult.status === 'fail') {
        summary.propChanges.push({
          feature: getFeatureName(row, col),
          prop: configResult.propName,
          message: configResult.info,
        });
      }
      if (pkgResult.status === 'fail') {
        summary.pkgChanges.push({
          feature: getFeatureName(row, col),
          packages: col.PACKAGES >= 0 ? extractPackages(row[col.PACKAGES]) : [],
          message: pkgResult.info,
        });
      }
    }

    updated += 1;
  }
  return updated;
}

function formatReportTimestamp(time = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${time.getFullYear()}${pad(time.getMonth() + 1)}${pad(time.getDate())}_${pad(time.getHours())}${pad(time.getMinutes())}${pad(time.getSeconds())}`;
}

function buildReportPath(excelPath, time = new Date()) {
  if (!excelPath) return null;
  return path.join(path.dirname(excelPath), `validation_results_${formatReportTimestamp(time)}.txt`);
}

function buildSummaryText(summary, excelPath, sourceNote) {
  const lines = [
    'Android 特性验证摘要',
    '='.repeat(40),
    `Excel 文件: ${excelPath}`,
    `数据来源: ${sourceNote}`,
  ];

  if (summary.productName) {
    lines.push(`设备 ro.product.name: ${summary.productName}`);
  }
  if (summary.supportHeader) {
    lines.push(`是否支持列: ${summary.supportHeader}${summary.supportSource === 'product' ? '（匹配设备型号）' : ''}`);
  }
  if (summary.stepHeader) {
    lines.push(`验证步骤列: ${summary.stepHeader}`);
  }
  if (summary.resultHeader) {
    lines.push(`验证结果列: ${summary.resultHeader}${summary.appendedResult ? '（已追加）' : ''}`);
  }

  lines.push(
    '',
    `总验证项数: ${summary.total}`,
    `pass 项数: ${summary.pass || 0}`,
    `fail 项数: ${summary.fail || 0}`,
    `NA 项数: ${summary.NA || 0}`,
    `未获取属性值项数: ${summary.missingProp || 0}`,
    '',
  );

  if (summary.propChanges.length) {
    lines.push('需要添加或修改的属性:', '-'.repeat(30));
    for (const item of summary.propChanges) {
      lines.push(`- [${item.feature}] ${item.message}`);
    }
    lines.push('');
  }

  if (summary.pkgChanges.length) {
    lines.push('需要移除或预置的应用包名:', '-'.repeat(30));
    for (const item of summary.pkgChanges) {
      lines.push(`- [${item.feature}] ${item.message} (${item.packages.join(', ')})`);
    }
    lines.push('');
  }

  if (summary.blockers.length) {
    lines.push('阻塞项:', '-'.repeat(30));
    for (const b of summary.blockers) lines.push(`- ${b}`);
  }

  return lines.join('\n');
}

function isAndroidFeatureSheet(tableData, productName = '') {
  if (!tableData?.sheetNames?.length) return false;
  return tableData.sheetNames.some((name) => sheetHasHeaders(tableData.sheets[name] || [], productName));
}

function describeAndroidSheetIssue(tableData, productName = '') {
  if (!tableData?.sheetNames?.length) return '表格为空';
  const lines = [];
  for (const name of tableData.sheetNames) {
    const rows = tableData.sheets[name] || [];
    if (!hasAndroidPropertyMarker(rows)) continue;
    const col = buildColumnMap(rows[0] || [], productName);
    const missing = [];
    if (col.SUPPORT < 0) missing.push('是否支持列（需匹配 ro.product.name 或包含支持/不支持）');
    if (col.STEP < 0) missing.push('验证步骤列或备注列');
    if (col.PROP < 0) missing.push('配置属性名');
    if (missing.length) lines.push(`工作表「${name}」缺少：${missing.join('、')}`);
  }
  if (!lines.length) return '未识别到包含「配置属性名」的 Android 特性验证工作表';
  return lines.join('\n');
}

async function fillAndroidExcel(options = {}) {
  const getpropFile = options.getpropFile || GETPROP_FILE;
  const packageFile = options.packageFile || PACKAGE_FILE;
  const tempJson = options.tempJson || TEMP_JSON;
  const blockers = [];
  let excelPath = options.excelPath || process.env.EXCEL_PATH || findMatchingExcel();
  let wb;
  let sourceNote;

  if (excelPath && fs.existsSync(excelPath)) {
    ({ wb } = loadWorkbookFromExcel(excelPath));
    sourceNote = `原始 Excel (${excelPath})`;
  } else if (tempJson && fs.existsSync(tempJson)) {
    blockers.push(`未找到 Excel 文件，已从临时 JSON 重建: ${tempJson}`);
    const fallbackDir = path.dirname(getpropFile);
    excelPath = path.join(fallbackDir, 'result.xlsx');
    ({ wb } = loadWorkbookFromJson(tempJson));
    sourceNote = `临时 JSON 重建 -> ${excelPath}`;
  } else if (options.tableData) {
    const data = options.tableData;
    wb = XLSX.utils.book_new();
    for (const name of data.sheetNames) {
      const ws = XLSX.utils.aoa_to_sheet(data.sheets[name]);
      XLSX.utils.book_append_sheet(wb, ws, name);
    }
    excelPath = options.excelPath;
    if (!excelPath) {
      throw new Error('缺少 Excel 文件路径');
    }
    sourceNote = `内存表格 -> ${excelPath}`;
  } else {
    throw new Error('找不到 Excel 文件、临时 JSON 或表格数据');
  }

  const props = loadPropsFromSources(getpropFile);
  if (props.size === 0) {
    blockers.push('未能从 getprop 文件或 adb getprop 加载任何属性');
  }

  const productName = getProductName(props);
  if (!productName) {
    blockers.push('未能读取 ro.product.name，将尝试使用表头中的「是否支持」或包含支持/不支持的列');
  }

  const { installedSet, pathTxtMap } = loadPackageData(packageFile);
  const adbPathCache = loadAllPackagePathsFromAdb();
  if (adbPathCache.size === 0) {
    blockers.push('adb 未返回任何已安装应用路径，请确认设备已连接 (adb devices)');
  }
  const sheets = options.tableData?.sheets || aoaFromWorkbook(wb);
  const summary = {
    total: 0,
    pass: 0,
    fail: 0,
    NA: 0,
    missingProp: 0,
    propChanges: [],
    pkgChanges: [],
    blockers,
    productName,
    supportHeader: '',
    supportSource: '',
    stepHeader: '',
    resultHeader: '',
    appendedResult: false,
  };

  const preserveFormat = options.preserveFormat !== false
    && excelPath
    && fs.existsSync(excelPath)
    && path.extname(excelPath).toLowerCase() !== '.xls';

  let totalUpdated = 0;
  const cellUpdates = preserveFormat ? [] : null;

  const sheetNames = options.tableData?.sheetNames || wb.SheetNames;
  for (const name of sheetNames) {
    const rows = sheets[name] || [];
    let col = buildColumnMap(rows[0] || [], productName);
    col = ensureResultColumn(rows, col);
    if (!summary.supportHeader && col.supportHeader) {
      summary.supportHeader = col.supportHeader;
      summary.supportSource = col.supportSource;
      summary.stepHeader = col.stepHeader;
      summary.resultHeader = col.resultHeader;
      summary.appendedResult = col.appendedResult;
    }
    if (!columnMapValid(col)) {
      blockers.push(`工作表「${name}」未能识别是否支持列或验证步骤/备注列`);
      continue;
    }
    totalUpdated += updateSheetRows(
      rows,
      props,
      installedSet,
      summary,
      name,
      cellUpdates,
      adbPathCache,
      pathTxtMap,
      col,
    );
    if (!preserveFormat) {
      wb.Sheets[name] = XLSX.utils.aoa_to_sheet(rows);
    }
  }

  if (preserveFormat) {
    const { fillCellsInPlace } = require('./excel-preserve-fill');
    const writeResult = await fillCellsInPlace(excelPath, cellUpdates);
    summary.formatPreserved = true;
    summary.cellsApplied = writeResult.applied;
    summary.cellsSkipped = writeResult.skipped;
  } else {
    writeWorkbook(wb, excelPath);
  }

  const reportPath = buildReportPath(excelPath, options.fillTime || new Date());
  fs.writeFileSync(reportPath, buildSummaryText(summary, excelPath, sourceNote), 'utf8');

  return {
    excelPath,
    reportPath,
    sourceNote,
    productName,
    supportHeader: summary.supportHeader,
    stepHeader: summary.stepHeader,
    resultHeader: summary.resultHeader,
    appendedResult: summary.appendedResult,
    updatedRows: totalUpdated,
    cellsApplied: summary.cellsApplied,
    cellsSkipped: summary.cellsSkipped,
    counts: {
      total: summary.total,
      pass: summary.pass || 0,
      fail: summary.fail || 0,
      NA: summary.NA || 0,
      missingProp: summary.missingProp || 0,
    },
    blockers: summary.blockers,
    summaryText: buildSummaryText(summary, excelPath, sourceNote),
  };
}

function loadPropsFromSources(getpropFile) {
  const fileProps = parseGetpropFile(getpropFile);
  const adbProps = tryAdbGetprop();
  const merged = new Map(fileProps);
  for (const [k, v] of adbProps) {
    if (!merged.has(k)) merged.set(k, v);
  }
  return merged;
}

function loadPackageData(packageFile = PACKAGE_FILE) {
  const installedSet = new Set();
  const pathTxtMap = new Map();
  if (!fs.existsSync(packageFile)) {
    return { installedSet, pathTxtMap };
  }

  for (const line of fs.readFileSync(packageFile, 'utf8').split(/\r?\n/)) {
    const parsed = parsePackagePathLine(line);
    if (!parsed?.pkg) continue;
    installedSet.add(parsed.pkg);
    if (parsed.path) {
      pathTxtMap.set(parsed.pkg, parsed.path);
    }
  }

  return { installedSet, pathTxtMap };
}

function formatFillOutput(result) {
  const {
    counts,
    excelPath,
    reportPath,
    updatedRows,
    blockers,
    cellsApplied,
    productName,
    supportHeader,
    stepHeader,
  } = result;
  const lines = [
    'Android 特性验证完成',
    `Excel: ${excelPath}`,
    `报告: ${reportPath}`,
  ];
  if (productName) lines.push(`设备 ro.product.name: ${productName}`);
  if (supportHeader) lines.push(`是否支持列: ${supportHeader}`);
  if (stepHeader) lines.push(`验证步骤列: ${stepHeader}`);
  lines.push(`已填写 ${updatedRows} 行（验证结果 / 验证步骤）`);
  if (cellsApplied != null) {
    lines.push(`已写入 ${cellsApplied} 个单元格（保留原格式）`);
  }
  lines.push('', `pass: ${counts.pass}  fail: ${counts.fail}  NA: ${counts.NA}  总计: ${counts.total}`);
  if (blockers.length) {
    lines.push('', '说明:', ...blockers.map((b) => `- ${b}`));
  }
  return lines.join('\n');
}

async function main() {
  try {
    const output = await fillAndroidExcel();
    console.log(JSON.stringify(output, null, 2));
  } catch (error) {
    console.error(`ERROR: ${error.message}`);
    process.exit(1);
  }
}

module.exports = {
  fillAndroidExcel,
  formatFillOutput,
  isAndroidFeatureSheet,
  describeAndroidSheetIssue,
  hasAndroidPropertyMarker,
  buildReportPath,
  formatReportTimestamp,
  buildColumnMap,
  resolveSupportColumn,
  getProductName,
};

if (require.main === module) {
  main();
}
