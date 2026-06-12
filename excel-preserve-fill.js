'use strict';

const ExcelJS = require('exceljs');

function isBlank(value) {
  if (value === undefined || value === null) return true;
  if (typeof value === 'string') return value.trim() === '';
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || '').join('').trim() === '';
    }
    if (value.result !== undefined && value.result !== null) {
      return String(value.result).trim() === '';
    }
    if (value.text !== undefined) {
      return String(value.text).trim() === '';
    }
  }
  return false;
}

function getCellText(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || '').join('');
    }
    if (value.result !== undefined && value.result !== null) {
      return String(value.result);
    }
    if (value.text !== undefined) {
      return String(value.text);
    }
  }
  return String(value);
}

async function fillCellsInPlace(filePath, updates) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  let applied = 0;
  let skipped = 0;

  for (const update of updates) {
    const sheet = workbook.getWorksheet(update.sheet);
    if (!sheet) {
      skipped += 1;
      continue;
    }

    const cell = sheet.getCell(update.row + 1, update.col + 1);
    if (!isBlank(cell.value)) {
      skipped += 1;
      continue;
    }

    cell.value = update.value;
    applied += 1;
  }

  await workbook.xlsx.writeFile(filePath);
  return { applied, skipped };
}

async function applyAgentFills(filePath, fills) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  let applied = 0;
  let skipped = 0;

  for (const fill of fills) {
    const sheet = workbook.getWorksheet(fill.sheet);
    if (!sheet) {
      skipped += 1;
      continue;
    }

    let cell;
    if (fill.cell) {
      cell = sheet.getCell(fill.cell);
    } else if (fill.row != null && fill.col != null) {
      cell = sheet.getCell(fill.row + 1, fill.col + 1);
    } else {
      skipped += 1;
      continue;
    }

    if (!isBlank(cell.value)) {
      skipped += 1;
      continue;
    }

    cell.value = fill.value;
    applied += 1;
  }

  await workbook.xlsx.writeFile(filePath);
  return { applied, skipped };
}

function extractFillsFromAgentOutput(text) {
  const trimmed = String(text || '').trim();
  if (!trimmed) {
    throw new Error('Agent 未返回任何内容');
  }

  const codeBlock = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = codeBlock ? codeBlock[1].trim() : trimmed;
  const jsonMatch = candidate.match(/\{[\s\S]*"fills"\s*:\s*\[[\s\S]*?\]\s*\}/);
  if (!jsonMatch) {
    throw new Error('Agent 未返回有效 JSON（需包含 fills 数组）');
  }

  const parsed = JSON.parse(jsonMatch[0]);
  if (!Array.isArray(parsed.fills)) {
    throw new Error('Agent JSON 格式错误：fills 必须是数组');
  }

  return parsed.fills.map((fill) => ({
    sheet: fill.sheet,
    cell: fill.cell,
    row: fill.row,
    col: fill.col,
    value: fill.value,
  }));
}

module.exports = {
  isBlank,
  getCellText,
  fillCellsInPlace,
  applyAgentFills,
  extractFillsFromAgentOutput,
};
