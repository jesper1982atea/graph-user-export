#!/usr/bin/env node
/*
  Update dist/ from GitHub latest release ZIP.
  Usage:
    node scripts/update-from-release.js [--tag v1.2.3] [--dest dist] [--clean] [--dry-run] [--asset-name name.zip]
  Env:
    GITHUB_TOKEN (optional, to avoid rate limiting)
*/
const fs = require('fs');
const path = require('path');
const os = require('os');
const JSZip = require('jszip');

function parseArgs(argv) {
  const args = { dest: 'dist', clean: false, dryRun: false };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--clean') args.clean = true;
    else if (a === '--dry-run') args.dryRun = true;
    else if (a === '--dest') args.dest = argv[++i] || args.dest;
    else if (a === '--tag') args.tag = argv[++i];
    else if (a === '--asset-name') args.assetName = argv[++i];
  }
  return args;
}

function parseRepoFromPackageJson() {
  try {
    const pkgPath = path.resolve(__dirname, '..', 'package.json');
    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    const url = pkg?.repository?.url || '';
    // supports formats like https://github.com/owner/repo.git
    const m = url.match(/github\.com[/:]([^/]+)\/([^/.]+)(?:\.git)?/i);
    if (m) return { owner: m[1], repo: m[2] };
  } catch {}
  return { owner: 'jesper1982atea', repo: 'graph-user-export' };
}

async function httpJson(url) {
  const headers = { 'Accept': 'application/vnd.github+json', 'User-Agent': 'update-script' };
  if (process.env.GITHUB_TOKEN) headers['Authorization'] = `Bearer ${process.env.GITHUB_TOKEN}`;
  const res = await fetch(url, { headers });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return res.json();
}

async function httpBuffer(url) {
  const headers = { 'Accept': 'application/octet-stream', 'User-Agent': 'update-script' };
  if (process.env.GITHUB_TOKEN) headers['Authorization'] = `Bearer ${process.env.GITHUB_TOKEN}`;
  const res = await fetch(url, { headers });
  if (!res.ok) throw new Error(`HTTP ${res.status} for asset download`);
  const arrayBuf = await res.arrayBuffer();
  return Buffer.from(arrayBuf);
}

function ensureDirSync(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function emptyDirSync(dir) {
  if (!fs.existsSync(dir)) return;
  for (const entry of fs.readdirSync(dir)) {
    const p = path.join(dir, entry);
    const st = fs.lstatSync(p);
    if (st.isDirectory()) {
      emptyDirSync(p);
      fs.rmdirSync(p);
    } else {
      fs.unlinkSync(p);
    }
  }
}

async function extractZipToDir(buffer, dest, { dryRun=false }={}) {
  const zip = await JSZip.loadAsync(buffer);
  const entries = [];
  zip.forEach((rel, entry) => { if (!entry.dir) entries.push(rel); });
  for (const rel of entries) {
    const fileBuf = await zip.file(rel).async('nodebuffer');
    const outPath = path.join(dest, rel);
    const outDir = path.dirname(outPath);
    if (!dryRun) ensureDirSync(outDir);
    if (!dryRun) fs.writeFileSync(outPath, fileBuf);
    console.log(`${dryRun ? '[dry]' : '[write]'} ${rel}`);
  }
  console.log(`Done ${dryRun ? '(dry-run)' : ''}. ${entries.length} files ${dryRun ? 'would be' : ''} written to ${dest}`);
}

async function main() {
  const args = parseArgs(process.argv);
  const { owner, repo } = parseRepoFromPackageJson();
  const assetName = args.assetName || 'graph-user-export-dist.zip';
  let buf = null;
  let tag = args.tag || '';
  // 1) Try stable GitHub download URL without using API (avoids rate limits)
  try {
    const url = args.tag
      ? `https://github.com/${owner}/${repo}/releases/download/${encodeURIComponent(args.tag)}/${encodeURIComponent(assetName)}`
      : `https://github.com/${owner}/${repo}/releases/latest/download/${encodeURIComponent(assetName)}`;
    console.log(`Attempting direct download: ${url}`);
    buf = await httpBuffer(url);
    console.log(`Downloaded asset via direct link: ${assetName}`);
  } catch (e) {
    console.warn(`Direct download failed (${e.message}). Falling back to API…`);
  }
  // 2) Fallback to API to discover asset URL
  if (!buf) {
    const base = `https://api.github.com/repos/${owner}/${repo}/releases`;
    const apiUrl = args.tag ? `${base}/tags/${encodeURIComponent(args.tag)}` : `${base}/latest`;
    console.log(`Fetching release via API: ${apiUrl}`);
    const rel = await httpJson(apiUrl);
    tag = rel.tag_name || rel.name || tag || 'unknown';
    const assets = Array.isArray(rel.assets) ? rel.assets : [];
    let asset = assets.find(a => a.name === assetName)
      || assets.find(a => a.name && a.name.endsWith('.zip') && a.name.includes('graph-user-export-dist'))
      || null;
    if (!asset) throw new Error('Could not find a dist ZIP in release assets');
    console.log(`Latest: ${tag}. Using asset: ${asset.name} (${Math.round(asset.size/1024)} KB)`);
    console.log(`Downloading: ${asset.browser_download_url}`);
    buf = await httpBuffer(asset.browser_download_url);
  }

  const dest = path.resolve(process.cwd(), args.dest || 'dist');
  ensureDirSync(dest);
  if (args.clean) {
    console.log(`Cleaning destination folder: ${dest}`);
    emptyDirSync(dest);
  }
  await extractZipToDir(buf, dest, { dryRun: !!args.dryRun });
}

// Node 18+ has global fetch; if not, require('node-fetch') fallback
async function withFetch() {
  if (typeof fetch === 'function') return main();
  global.fetch = (await import('node-fetch')).default;
  return main();
}

withFetch().catch(err => {
  console.error('Update failed:', err.message || err);
  process.exitCode = 1;
});
