// Simple update checker that queries GitHub Releases
const OWNER = 'jesper1982atea';
const REPO = 'graph-user-export';

const API_URL = `https://api.github.com/repos/${OWNER}/${REPO}/releases/latest`;

function normalizeVersion(v) {
  if (!v) return '0.0.0';
  const s = String(v).trim().replace(/^v/i, '');
  return s.split('-')[0]; // drop pre-release/build metadata
}

export function compareVersions(a, b) {
  const as = normalizeVersion(a).split('.').map(n => parseInt(n, 10) || 0);
  const bs = normalizeVersion(b).split('.').map(n => parseInt(n, 10) || 0);
  const len = Math.max(as.length, bs.length);
  for (let i = 0; i < len; i++) {
    const ai = as[i] || 0;
    const bi = bs[i] || 0;
    if (ai > bi) return 1;
    if (ai < bi) return -1;
  }
  return 0;
}

export async function getLatestRelease() {
  const res = await fetch(API_URL, { headers: { 'Accept': 'application/vnd.github+json' } });
  if (!res.ok) throw new Error(`GitHub API error ${res.status}`);
  const json = await res.json();
  const assets = Array.isArray(json.assets) ? json.assets : [];
  return {
    tagName: json.tag_name || json.name || '',
    htmlUrl: json.html_url,
    publishedAt: json.published_at,
    assets: assets.map(a => ({
      name: a.name,
      url: a.url,
      browser_download_url: a.browser_download_url,
      size: a.size,
      content_type: a.content_type,
    })),
  };
}

export async function checkForUpdate(currentVersion) {
  const latest = await getLatestRelease();
  const cur = normalizeVersion(currentVersion);
  const lat = normalizeVersion(latest.tagName);
  const cmp = compareVersions(lat, cur);
  // Prefer the static-named dist zip, fallback to any zip containing 'graph-user-export-dist'
  const preferred = latest.assets.find(a => a.name === 'graph-user-export-dist.zip');
  const fallback = latest.assets.find(a => a.name && a.name.endsWith('.zip') && a.name.includes('graph-user-export-dist'));
  return {
    hasUpdate: cmp === 1,
    current: cur,
    latest: lat,
    releaseHtmlUrl: latest.htmlUrl,
    assetName: (preferred || fallback)?.name || '',
    assetUrl: (preferred || fallback)?.browser_download_url || latest.htmlUrl,
    publishedAt: latest.publishedAt,
  };
}
