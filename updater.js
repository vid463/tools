'use strict';

const https = require('https');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

const GITHUB_OWNER = 'vid463';
const GITHUB_REPO = 'tools';
const API_BASE = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}`;

function normalizeVersion(version) {
  return String(version || '').trim().replace(/^v/i, '');
}

function compareVersions(current, latest) {
  const parse = (value) => normalizeVersion(value)
    .split('.')
    .map((part) => Number.parseInt(part, 10) || 0);
  const left = parse(current);
  const right = parse(latest);
  const length = Math.max(left.length, right.length);

  for (let i = 0; i < length; i += 1) {
    const a = left[i] || 0;
    const b = right[i] || 0;
    if (a > b) return 1;
    if (a < b) return -1;
  }
  return 0;
}

function requestJson(url) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, {
      headers: {
        Accept: 'application/vnd.github+json',
        'User-Agent': 'tools-updater',
      },
    }, (response) => {
      if (response.statusCode >= 300
        && response.statusCode < 400
        && response.headers.location) {
        requestJson(response.headers.location).then(resolve).catch(reject);
        return;
      }

      if (response.statusCode !== 200) {
        reject(new Error(`GitHub API 请求失败：HTTP ${response.statusCode}`));
        response.resume();
        return;
      }

      const chunks = [];
      response.on('data', (chunk) => chunks.push(chunk));
      response.on('end', () => {
        try {
          resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
        } catch (error) {
          reject(new Error(`解析 GitHub 响应失败：${error.message}`));
        }
      });
    });

    request.on('error', reject);
    request.setTimeout(30000, () => {
      request.destroy(new Error('检查更新超时'));
    });
  });
}

function pickWindowsInstaller(assets = []) {
  const exeAssets = assets.filter((asset) => /\.exe$/i.test(asset.name)
    && !/blockmap/i.test(asset.name));

  return exeAssets.find((asset) => /setup/i.test(asset.name))
    || exeAssets[0]
    || null;
}

async function fetchLatestRelease() {
  const release = await requestJson(`${API_BASE}/releases/latest`);
  const version = normalizeVersion(release.tag_name);
  const asset = pickWindowsInstaller(release.assets || []);

  if (!version) {
    throw new Error('GitHub Release 缺少有效版本号');
  }
  if (!asset?.browser_download_url) {
    throw new Error('GitHub Release 未找到 Windows 安装包（.exe）');
  }

  return {
    version,
    tagName: release.tag_name,
    name: release.name || release.tag_name,
    notes: release.body || '',
    htmlUrl: release.html_url,
    downloadUrl: asset.browser_download_url,
    fileName: asset.name,
    size: asset.size || 0,
  };
}

async function checkForUpdates(currentVersion) {
  const latest = await fetchLatestRelease();
  const current = normalizeVersion(currentVersion);
  const available = compareVersions(current, latest.version) < 0;

  return {
    available,
    currentVersion: current,
    latestVersion: latest.version,
    release: latest,
  };
}

function downloadFile(url, destPath, onProgress) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, {
      headers: { 'User-Agent': 'tools-updater' },
    }, (response) => {
      if (response.statusCode >= 300
        && response.statusCode < 400
        && response.headers.location) {
        downloadFile(response.headers.location, destPath, onProgress)
          .then(resolve)
          .catch(reject);
        return;
      }

      if (response.statusCode !== 200) {
        reject(new Error(`下载失败：HTTP ${response.statusCode}`));
        response.resume();
        return;
      }

      const total = Number(response.headers['content-length'] || 0);
      let received = 0;
      const file = fs.createWriteStream(destPath);

      response.on('data', (chunk) => {
        received += chunk.length;
        if (onProgress) onProgress(received, total);
      });

      response.on('error', reject);
      file.on('error', reject);
      file.on('finish', () => {
        file.close(() => resolve(destPath));
      });

      response.pipe(file);
    });

    request.on('error', reject);
    request.setTimeout(10 * 60 * 1000, () => {
      request.destroy(new Error('下载更新超时'));
    });
  });
}

function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

async function downloadUpdate(release, destDir, onProgress) {
  const safeName = path.basename(release.fileName || `tools-setup-${release.version}.exe`);
  const destPath = path.join(destDir, safeName);
  fs.mkdirSync(destDir, { recursive: true });

  if (fs.existsSync(destPath)) {
    fs.unlinkSync(destPath);
  }

  await downloadFile(release.downloadUrl, destPath, onProgress);
  return destPath;
}

module.exports = {
  GITHUB_OWNER,
  GITHUB_REPO,
  normalizeVersion,
  compareVersions,
  fetchLatestRelease,
  checkForUpdates,
  downloadUpdate,
  formatBytes,
};
