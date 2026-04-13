const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwGWwOrSMSieQ56AXszY_5RjO88ETNBqgIGaorpWfEEAUJ8g1GOBxtnH3LWl6SyVAUh/exec";
const FALLBACK_FAVICON = 'data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 64 64%22%3E%3Crect width=%2264%22 height=%2264%22 rx=%2214%22 fill=%22%23081224%22/%3E%3Ctext x=%2250%25%22 y=%2254%25%22 font-size=%2232%22 text-anchor=%22middle%22 fill=%22%23facc15%22 font-family=%22Arial%22 dy=%22.1em%22%3EG%3C/text%3E%3C/svg%3E';
const PUBLIC_SITE_CACHE_KEY = 'genz_cbt_public_site_content_v1';
const PUBLIC_SITE_BUST_KEY = 'genz_cbt_public_site_bust_v1';

function antiEmbedGuard(){
  try{
    if(window.top !== window.self){
      document.documentElement.style.display = 'none';
      window.top.location = window.self.location.href;
    }
  }catch(err){
    try{ window.location = window.self.location.href; }catch(e){}
  }
}
antiEmbedGuard();

function escHtml(v){ return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
function nl2br(v){ return escHtml(v).replace(/\n/g,'<br>'); }
function safeList(arr,fallback){ return Array.isArray(arr) && arr.length ? arr : (fallback || []); }
function setSiteFavicon(url){ const el = document.getElementById('siteFavicon'); if(el) el.href = normalizePublicImageUrl(url) || FALLBACK_FAVICON; }
function youtubeEmbed(url){ if(!url) return ''; const m = String(url).match(/(?:youtu\.be\/|v=|embed\/|shorts\/)([A-Za-z0-9_-]{6,})/); return m ? `https://www.youtube.com/embed/${m[1]}` : ''; }
function unwrapImageInput(url){
  const raw = String(url || '').trim();
  if(!raw) return '';
  const imgMatch = raw.match(/<img\b[^>]*\bsrc\s*=\s*(["'])(.*?)\1/i) || raw.match(/<img\b[^>]*\bsrc\s*=\s*([^\s>]+)/i);
  if(imgMatch) return String(imgMatch[2] || imgMatch[1] || '').trim();
  const cssMatch = raw.match(/url\(\s*(["']?)(.*?)\1\s*\)/i);
  if(cssMatch && cssMatch[2]) return String(cssMatch[2]).trim();
  return raw;
}
function extractDriveImageId(url){
  const raw = unwrapImageInput(url);
  if(!raw) return '';
  let m = raw.match(/drive\.google\.com\/file\/d\/([^\/]+)/i);
  if(!m) m = raw.match(/drive\.google\.com\/open\?id=([^&]+)/i);
  if(!m) m = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/i);
  if(!m) m = raw.match(/uc\?export=(?:view|download)&id=([a-zA-Z0-9_-]+)/i);
  if(!m) m = raw.match(/thumbnail\?[^#]*id=([a-zA-Z0-9_-]+)/i);
  if(!m) m = raw.match(/drive\.usercontent\.google\.com\/(?:download|u\/\d+\/uc)\?[^#]*id=([a-zA-Z0-9_-]+)/i);
  if(!m && /^[A-Za-z0-9_-]{20,}$/.test(raw)) m = [raw, raw];
  return m && m[1] ? m[1] : '';
}
function normalizePublicImageUrl(url){
  const raw = unwrapImageInput(url);
  if(!raw) return '';
  const driveId = extractDriveImageId(raw);
  if(driveId) return `https://drive.google.com/thumbnail?id=${driveId}&sz=w2000`;
  if(/dropbox\.com/i.test(raw)){
    let next = raw.replace(/\?dl=0$/i,'?raw=1').replace(/\?dl=1$/i,'?raw=1');
    if(!/[?&]raw=1/i.test(next)) next += (next.includes('?') ? '&' : '?') + 'raw=1';
    return next;
  }
  const m = raw.match(/^https:\/\/github\.com\/([^\/]+)\/([^\/]+)\/blob\/([^\/]+)\/(.+)$/i);
  if(m) return `https://raw.githubusercontent.com/${m[1]}/${m[2]}/${m[3]}/${m[4]}`;
  if(/^https:\/\/1drv\.ms\//i.test(raw)){
    try{ return `https://api.onedrive.com/v1.0/shares/u!${btoa(raw).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'')}/root/content`; }catch(err){}
  }
  return raw;
}
function publicImageCandidates(url){
  const raw = unwrapImageInput(url);
  if(!raw) return [];
  const out = [];
  const push = (v) => { const s = String(v || '').trim(); if(s && !out.includes(s)) out.push(s); };
  const normalized = normalizePublicImageUrl(raw);
  push(normalized);
  const driveId = extractDriveImageId(raw);
  if(driveId){
    push(`https://drive.google.com/thumbnail?id=${driveId}&sz=w2000`);
    push(`https://drive.google.com/uc?export=view&id=${driveId}`);
    push(`https://drive.google.com/uc?id=${driveId}`);
    push(`https://drive.usercontent.google.com/download?id=${driveId}&export=view&authuser=0`);
  }
  return out;
}
function readPublicSiteCache(){
  try{
    const raw = localStorage.getItem(PUBLIC_SITE_CACHE_KEY);
    if(!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && parsed.data ? parsed : null;
  }catch(err){ return null; }
}
function normalizePublicSiteData(data){
  const safe = Object.assign({}, data || {});
  ['siteFaviconUrl','resultBrandLogoUrl','ceoImageUrl'].forEach(key => { if(safe[key]) safe[key] = normalizePublicImageUrl(safe[key]); });
  ['contributors','institutions','testimonials','tutorialVideos','socialLinks'].forEach(listKey => {
    safe[listKey] = safeList(safe[listKey], []).map(item => {
      const out = Object.assign({}, item || {});
      ['imageUrl','logoUrl','thumbUrl','avatarUrl','iconUrl','profileUrl'].forEach(key => { if(out[key]) out[key] = normalizePublicImageUrl(out[key]); });
      return out;
    });
  });
  return safe;
}
function writePublicSiteCache(data){
  try{ localStorage.setItem(PUBLIC_SITE_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), data: normalizePublicSiteData(data || {}) })); }catch(err){}
}
function clearPublicSiteCache(){
  try{ localStorage.removeItem(PUBLIC_SITE_CACHE_KEY); }catch(err){}
}
function bindPublicImage(img, url, loadingMode, finalFallback){
  if(!img) return;
  const candidates = publicImageCandidates(url);
  if(!candidates.length){ img.style.display='none'; img.removeAttribute('src'); if(typeof finalFallback === 'function') finalFallback(false); return; }
  img.referrerPolicy = 'no-referrer';
  img.decoding = 'async';
  img.loading = loadingMode || 'eager';
  let index = 0;
  const tryNext = () => {
    if(index >= candidates.length){
      img.style.display = 'none';
      img.removeAttribute('src');
      if(typeof finalFallback === 'function') finalFallback(false);
      return;
    }
    const nextSrc = candidates[index++];
    img.onerror = tryNext;
    img.onload = function(){ this.style.display = 'block'; if(typeof finalFallback === 'function') finalFallback(true, nextSrc); };
    img.src = nextSrc;
    img.style.display = 'block';
  };
  tryNext();
}
function getCachedPublicSiteContent(){
  const cached = readPublicSiteCache();
  return cached && cached.data ? normalizePublicSiteData(cached.data) : {};
}
function preloadPublicImage(url){
  const candidates = publicImageCandidates(url);
  if(!candidates.length) return;
  try{ const img = new Image(); img.decoding = 'async'; img.referrerPolicy = 'no-referrer'; img.src = String(candidates[0]); }catch(err){}
}
function warmPublicSiteAssets(data){
  if(!data || typeof data !== 'object') return;
  [data.siteFaviconUrl, data.resultBrandLogoUrl, data.ceoImageUrl].forEach(preloadPublicImage);
}
async function fetchPublicSiteContentRemote(){
  const u = new URL(APPS_SCRIPT_URL);
  u.searchParams.set('action','getPublicSiteContent');
  u.searchParams.set('_ts', Date.now());
  const res = await fetch(u.toString(), { cache:'no-store' });
  const data = await res.json();
  return data && data.ok ? normalizePublicSiteData(data.data || {}) : {};
}
async function getPublicSiteContent(options){
  const opts = options || {};
  const bust = (() => { try{ return localStorage.getItem(PUBLIC_SITE_BUST_KEY) || ''; }catch(err){ return ''; } })();
  const cached = readPublicSiteCache();
  const normalizedCached = cached && cached.data ? normalizePublicSiteData(cached.data || {}) : {};
  const cachedVersion = normalizedCached.publicContentVersion || '';
  if(cached && !opts.preferFresh){
    warmPublicSiteAssets(normalizedCached);
    fetchPublicSiteContentRemote().then(fresh => {
      if(fresh && Object.keys(fresh).length){
        writePublicSiteCache(fresh);
        const changed = (fresh.publicContentVersion || '') !== cachedVersion || !!bust;
        if(typeof opts.onFresh === 'function'){
          try{ opts.onFresh(fresh, changed); }catch(err){}
        }
        if(changed){
          try{ window.dispatchEvent(new CustomEvent('genz:public-site-content-updated', { detail: fresh })); }catch(err){}
          try{ localStorage.removeItem(PUBLIC_SITE_BUST_KEY); }catch(err){}
        }
      }
    }).catch(() => {});
    return normalizedCached;
  }
  try{
    const fresh = await fetchPublicSiteContentRemote();
    if(fresh && Object.keys(fresh).length){
      writePublicSiteCache(fresh);
      warmPublicSiteAssets(fresh);
      try{ localStorage.removeItem(PUBLIC_SITE_BUST_KEY); }catch(err){}
      return fresh;
    }
  }catch(err){}
  return normalizedCached;
}
function renderBrand(root, data){
  if(!root) return;
  const logo = data && data.resultBrandLogoUrl ? String(data.resultBrandLogoUrl) : '';
  if(!logo){ root.innerHTML = '<div class="mark">G</div>'; return; }
  root.innerHTML = '<img alt="Brand logo" style="display:none" referrerpolicy="no-referrer" decoding="async"><div class="mark" style="display:none">G</div>';
  const img = root.querySelector('img');
  const fallback = root.querySelector('.mark');
  bindPublicImage(img, logo, 'eager', (ok) => {
    if(ok){ if(fallback) fallback.style.display = 'none'; }
    else {
      if(img) img.style.display = 'none';
      if(fallback) fallback.style.display = 'grid';
    }
  });
}
function renderFooter(data){
  const footer = document.getElementById('siteFooter');
  if(!footer) return;
  const socials = safeList(data.socialLinks, []).map(item => `<a href="${escHtml(item.url||'#')}" target="_blank" rel="noopener">${escHtml(item.label||'Social')}</a>`).join('');
  footer.innerHTML = `
    <div class="footer-links">
      <a href="index.html">Home</a>
      <a href="about.html">About</a>
      <a href="terms.html">Terms</a>
      <a href="privacy.html">Privacy</a>
      <a href="cookies.html">Cookies</a>
      <a href="contact.html">Contact</a>
      <a href="student.html">Student Portal</a>
      <a href="admin.html">Admin Portal</a>
      <a href="result_checker.html">Result Checker</a>
    </div>
    <div style="margin-top:18px" class="social">${socials || '<a href="#">Add social links from admin settings</a>'}</div>
    <div style="margin-top:16px">${escHtml(data.resultBrandName || 'Genz EduTech Innovations')} · Built for modern exam delivery and result presentation.</div>
  `;
}
function cardMedia(item, fallbackLabel){
  if(item.videoUrl){
    const yt = youtubeEmbed(item.videoUrl);
    if(yt) return `<div class="media-thumb"><iframe src="${escHtml(yt)}" allowfullscreen loading="lazy" referrerpolicy="no-referrer"></iframe></div>`;
  }
  if(item.imageUrl || item.logoUrl || item.thumbUrl || item.avatarUrl){
    const src = normalizePublicImageUrl(item.imageUrl || item.logoUrl || item.thumbUrl || item.avatarUrl);
    return `<div class="media-thumb"><img src="${escHtml(src)}" alt="${escHtml(item.name || item.title || fallbackLabel || 'Media')}" loading="lazy" decoding="async" referrerpolicy="no-referrer"></div>`;
  }
  return `<div class="media-thumb" style="display:grid;place-items:center;color:#c7d2fe">${escHtml(fallbackLabel || 'Media')}</div>`;
}
