/* ════════════════════════════════════════════
   연구실적 DOI 검증 시스템 — app.js
   ════════════════════════════════════════════ */

let rows = [];
let verifying = false;
let verifyQueue = [];
let researchers = [];   // researchers.json 로드 데이터
let scopusCache = {};   // SCOPUS ID → 영문명 캐시 (세션 내)

/* ── 설정 패널 ── */
function toggleSettings() {
  const p = document.getElementById('settingsPanel');
  p.style.display = p.style.display === 'none' ? 'block' : 'none';
  // 저장된 검증 대상자 이름 복원
  const savedName = localStorage.getItem('targetName');
  if (savedName) {
    document.getElementById('targetName').value = savedName;
    showTargetStatus(`✓ "${savedName}" 설정됨`, 'ok');
  }
}

function saveTargetName() {
  const name = document.getElementById('targetName').value.trim();
  if (!name) { showTargetStatus('이름을 입력하세요.', 'fail'); return; }
  localStorage.setItem('targetName', name);
  showTargetStatus(`✓ "${name}" 저장됨`, 'ok');
}

function getTargetName() {
  return localStorage.getItem('targetName') || '';
}

/* ── 업적평가 기간 ── */
function initEvalYear() {
  const sel = document.getElementById('evalYear');
  if (!sel) return;
  const cur = new Date().getFullYear();
  for (let y = cur; y >= cur - 10; y--) {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y + '년';
    sel.appendChild(opt);
  }
  // 저장된 값 복원
  const saved = JSON.parse(localStorage.getItem('evalPeriod') || '{}');
  if (saved.type) document.getElementById('evalType').value = saved.type;
  if (saved.year) sel.value = saved.year;
  updateEvalDisplay();
}

function saveEvalPeriod() {
  const type = document.getElementById('evalType').value;
  const year = parseInt(document.getElementById('evalYear').value);
  localStorage.setItem('evalPeriod', JSON.stringify({type, year}));
  updateEvalDisplay();
}

function updateEvalDisplay() {
  const p = getEvalPeriod();
  if (!p) return;
  const el = document.getElementById('evalPeriodDisplay');
  if (el) el.textContent = `${p.start} ~ ${p.end}`;
}

function getEvalPeriod() {
  const saved = JSON.parse(localStorage.getItem('evalPeriod') || '{}');
  const type = saved.type || 'first';
  const year = saved.year || new Date().getFullYear();
  if (type === 'first') {
    return { start: `${year}-01-01`, end: `${year}-12-31`, label: `${year}년 전반기` };
  } else {
    // 후반기: 전년도 7/1 ~ 당해 6/30
    return { start: `${year-1}-07-01`, end: `${year}-06-30`, label: `${year}년 후반기` };
  }
}

function checkPeriod(doiDate) {
  if (!doiDate) return null;
  const p = getEvalPeriod();
  if (!p) return null;
  // 날짜를 YYYYMMDD 숫자로 비교
  const toNum = s => parseInt(s.replace(/-/g,'').slice(0,8).padEnd(8,'0'));
  const d = toNum(doiDate);
  const start = toNum(p.start);
  const end = toNum(p.end);
  return { inPeriod: d >= start && d <= end, period: p.label, date: doiDate };
}

// 페이지 로드 시 연도 초기화
document.addEventListener('DOMContentLoaded', initEvalYear);

function showTargetStatus(msg, cls) {
  const el = document.getElementById('targetNameStatus');
  if (el) { el.textContent = msg; el.className = 'settings-status ' + cls; }
}

function saveScopusKey() {
  const key = document.getElementById('scopusApiKey').value.trim();
  if (!key) { showKeyStatus('API Key를 입력하세요.', 'fail'); return; }
  localStorage.setItem('scopusApiKey', key);
  showKeyStatus('✓ 저장되었습니다.', 'ok');
}

async function testScopusKey() {
  const key = getScopusKey();
  if (!key) { showKeyStatus('API Key를 먼저 입력하고 저장하세요.', 'fail'); return; }
  showKeyStatus('연결 테스트 중...', 'loading');
  try {
    const res = await fetch(
      'https://api.elsevier.com/content/author/author_id/7201711855?field=preferred-name',
      { headers: { 'X-ELS-APIKey': key, 'Accept': 'application/json' } }
    );
    if (res.ok) {
      showKeyStatus('✓ 연결 성공! Scopus API를 사용할 수 있습니다.', 'ok');
    } else {
      showKeyStatus(`✗ 연결 실패 (${res.status}) — API Key를 확인해주세요.`, 'fail');
    }
  } catch(e) {
    showKeyStatus('✗ 네트워크 오류: ' + e.message, 'fail');
  }
}

function showKeyStatus(msg, cls) {
  const el = document.getElementById('scopusKeyStatus');
  el.textContent = msg;
  el.className = 'settings-status ' + cls;
}

function getScopusKey() {
  return localStorage.getItem('scopusApiKey') || '';
}

function loadResearchers(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      researchers = JSON.parse(ev.target.result);
      const el = document.getElementById('researchersStatus');
      el.textContent = `✓ ${researchers.length}명 로드됨`;
      el.className = 'settings-status ok';
      localStorage.setItem('researchers', JSON.stringify(researchers));
    } catch(err) {
      const el = document.getElementById('researchersStatus');
      el.textContent = '✗ JSON 파싱 오류: ' + err.message;
      el.className = 'settings-status fail';
    }
  };
  reader.readAsText(file);
}

// 페이지 로드 시 researchers.json 자동 로드 (같은 저장소에서)
(function initSettings() {
  // 1. 먼저 로컬스토리지에서 복원 시도
  const saved = localStorage.getItem('researchers');
  if (saved) {
    try {
      researchers = JSON.parse(saved);
      console.log(`researchers.json 로컬 캐시 로드: ${researchers.length}명`);
    } catch(e) {}
  }
  // 2. GitHub 저장소에서 최신 researchers.json 자동 로드
  fetch('researchers.json')
    .then(r => r.ok ? r.json() : null)
    .then(data => {
      if (data && data.length) {
        researchers = data;
        // 로컬스토리지 크기 제한 대비 — 실패해도 무시
        try { localStorage.setItem('researchers', JSON.stringify(data)); } catch(e) {}
        console.log(`researchers.json 자동 로드 완료: ${researchers.length}명`);
        // 설정 패널 상태 업데이트
        const el = document.getElementById('researchersStatus');
        if (el) {
          el.textContent = `✓ ${researchers.length}명 자동 로드됨 (researchers.json)`;
          el.className = 'settings-status ok';
        }
      }
    })
    .catch(() => {
      console.log('researchers.json 자동 로드 실패 — 설정에서 수동 업로드 필요');
    });
})();

/* ── SCOPUS API ── */
function findResearcher(nameKo) {
  if (!nameKo || !researchers.length) return null;
  const name = nameKo.replace(/\s/g, '');
  return researchers.find(r => r.nameKo && r.nameKo.replace(/\s/g, '') === name) || null;
}

async function fetchScopusName(scopusId) {
  if (!scopusId) return null;
  if (scopusCache[scopusId]) return scopusCache[scopusId];
  const key = getScopusKey();
  if (!key) return null;
  try {
    const res = await fetch(
      `https://api.elsevier.com/content/author/author_id/${scopusId}?field=preferred-name`,
      { headers: { 'X-ELS-APIKey': key, 'Accept': 'application/json' } }
    );
    if (!res.ok) return null;
    const data = await res.json();
    const pref = data['author-retrieval-response']?.[0]?.['author-profile']?.['preferred-name'];
    if (!pref) return null;
    const result = {
      family: (pref['surname'] || '').trim(),
      given:  (pref['given-name'] || '').trim(),
    };
    scopusCache[scopusId] = result;
    return result;
  } catch { return null; }
}

/* ── 파일 처리 ── */
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag');
  const f = e.dataTransfer.files[0];
  if (f) processFile(f);
}
function handleFile(e) {
  const f = e.target.files[0];
  if (f) processFile(f);
}

function processFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const raw = e.target.result;
      const bytes = new Uint8Array(raw);
      const head = new TextDecoder('utf-8').decode(bytes.slice(0, 512));
      const isXml = head.includes('<?xml') || head.includes('schemas-microsoft-com');
      let json;
      if (isXml) {
        const xmlText = new TextDecoder('utf-8').decode(bytes);
        json = parseXmlSpreadsheet(xmlText);
      } else {
        const wb = XLSX.read(bytes, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      }
      if (!json || !json.length) { alert('데이터를 읽을 수 없습니다.'); return; }
      buildRows(json);
    } catch(err) {
      alert('파일 읽기 오류: ' + (err.message || String(err)));
    }
  };
  reader.onerror = function() { alert('파일을 열 수 없습니다.'); };
  reader.readAsArrayBuffer(file);
}

function parseXmlSpreadsheet(xmlText) {
  // CDATA 언래핑
  const unwrap = (s) => s.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1').trim();

  // Row 블록 추출
  const rowMatches = xmlText.match(/<(?:[a-z]+:)?Row[^>]*>[\s\S]*?<\/(?:[a-z]+:)?Row>/gi);
  if (!rowMatches || !rowMatches.length) throw new Error('Row 데이터를 찾을 수 없습니다.');

  // Cell 추출 함수
  const parseCells = (rowHtml) => {
    const cells = [];
    const cellRe = /<(?:[a-z]+:)?Cell[^>]*>([\s\S]*?)<\/(?:[a-z]+:)?Cell>/gi;
    let cm;
    while ((cm = cellRe.exec(rowHtml)) !== null) {
      const inner = cm[1];
      const dataMatch = inner.match(/<(?:[a-z]+:)?Data[^>]*>([\s\S]*?)<\/(?:[a-z]+:)?Data>/i);
      cells.push(dataMatch ? unwrap(dataMatch[1]) : '');
    }
    return cells;
  };

  // 헤더
  const headers = parseCells(rowMatches[0]);
  if (!headers.length) throw new Error('헤더를 찾을 수 없습니다.');

  // 데이터
  const json = [];
  for (let i = 1; i < rowMatches.length; i++) {
    const vals = parseCells(rowMatches[i]);
    const obj = {};
    headers.forEach((h, idx) => {
      if (h) obj[h] = vals[idx] !== undefined ? vals[idx] : '';
    });
    if (Object.values(obj).some(v => v !== '')) json.push(obj);
  }
  return json;
}

function buildRows(json) {
  const colMap = detectColumns(Object.keys(json[0]));
  rows = json.map((r, i) => ({
    idx: i + 1,
    title: str(r[colMap.title]),
    doi: str(r[colMap.doi]).replace(/^https?:\/\/doi\.org\//i, '').trim(),
    date: formatDate(r[colMap.date]),
    journal: str(r[colMap.journal]),
    role: str(r[colMap.role]),
    corrNames: str(r[colMap.corrNames]),
    corrCount: parseInt(str(r[colMap.corrCount])) || 0,
    authorCount: parseInt(str(r[colMap.authorCount])) || 0,
    researcherName: str(r[colMap.name]),
    status: 'pending',
    doiDate: null, doiJournal: null, doiTitle: null, doiAuthorCount: null, doiAuthorCount: null, doiAuthorCount: null, doiAllDates: [], periodResult: null,
    inferredRole: null, roleSource: null,
    issues: [], roleIssue: false, doiNote: ''
  }));

  // 검증 대상자 설정된 경우 모든 행에 적용
  const targetName = getTargetName();
  if (targetName) {
    rows.forEach(r => { r.researcherName = targetName; });
  }
  showResults();
  updateStats();
  renderTable();
}

function str(v) { return (v || '').toString().trim(); }

function detectColumns(keys) {
  // 정확히 일치하는 컬럼 우선, 없으면 포함 매칭
  const findExact = (exact) => keys.find(k => k === exact) || '';
  const find = (...patterns) =>
    keys.find(k => patterns.some(p => k.toLowerCase().includes(p))) || '';
  return {
    title:     findExact('논문제목') || find('논문제목', 'title', '제목'),
    doi:       findExact('doi(논문아이디)') || find('doi', '아이디'),
    date:      findExact('발표일') || find('발표일'),
    journal:   findExact('학술지명') || find('학술지명', 'journal'),
    role:      findExact('참여형태') || find('참여형태', '역할'),
    corrNames:   findExact('교신저자명') || find('교신저자명'),
    corrCount:   findExact('교신저자수') || find('교신저자수'),
    authorCount: findExact('참여자수') || find('참여자수'),
    name:        findExact('성명') || find('성명', 'name'),
  };
}

function formatDate(v) {
  if (!v) return '';
  const s = v.toString().trim().replace(/\./g, '-');
  if (/^\d{8}$/.test(s)) return s.slice(0,4)+'-'+s.slice(4,6)+'-'+s.slice(6,8);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return s;
}

function showResults() {
  document.getElementById('uploadSection').style.display = 'none';
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('statsBar').style.display = 'flex';
  document.getElementById('ctrlBar').style.display = 'flex';
  document.getElementById('tableSection').style.display = 'block';
  document.getElementById('btnVerifyAll').disabled = false;
  document.getElementById('btnExport').disabled = false;
}

function resetFile() {
  rows = [];
  verifying = false;
  document.getElementById('uploadSection').style.display = 'block';
  document.getElementById('emptyState').style.display = 'flex';
  document.getElementById('statsBar').style.display = 'none';
  document.getElementById('ctrlBar').style.display = 'none';
  document.getElementById('tableSection').style.display = 'none';
  document.getElementById('btnVerifyAll').disabled = true;
  document.getElementById('btnExport').disabled = true;
  document.getElementById('fileInput').value = '';
}

/* ── 전체 검증 ── */
async function verifyAll() {
  if (verifying) return;
  verifying = true;
  const btn = document.getElementById('btnVerifyAll');
  btn.disabled = true;
  btn.innerHTML = '<span class="btn-icon">⏳</span> 검증 중...';

  const progressWrap = document.getElementById('progressWrap');
  const progressBar  = document.getElementById('progressBar');
  const progressLabel = document.getElementById('progressLabel');
  progressWrap.style.display = 'flex';

  const toVerify = rows.filter(r => r.doi);
  for (let i = 0; i < toVerify.length; i++) {
    const r = toVerify[i];
    r.status = 'checking';
    r.inferredRole = 'checking';
    renderTable();
    await verifyRow(r);
    const pct = Math.round((i + 1) / toVerify.length * 100);
    progressBar.style.width = pct + '%';
    progressLabel.textContent = `${i+1} / ${toVerify.length}`;
    updateStats();
    renderTable();
    await sleep(250); // rate-limit 방지
  }

  // DOI 없는 항목 → 미검증
  rows.filter(r => !r.doi).forEach(r => {
    r.status = 'warn';
    r.doiNote = 'DOI 없음';
  });

  verifying = false;
  progressWrap.style.display = 'none';
  btn.disabled = false;
  btn.innerHTML = '<span class="btn-icon">▶</span> 재검증';
  updateStats();
  renderTable();
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

/* ── 단건 검증 ── */
async function verifyRow(row) {
  try {
    // 0. 영문명 조회 — 설정의 검증대상자 이름 또는 엑셀 성명 컬럼 사용
    const targetName = getTargetName();
    const lookupName = targetName || row.researcherName;
    const researcher = findResearcher(lookupName);
    if (researcher && researcher.familyName) {
      row._engName = {
        family: researcher.familyName,
        given:  researcher.givenName || ''
      };
    } else if (!researcher && lookupName) {
      // researchers.json에 없으면 엑셀 성명을 그대로 researcherName으로 사용
      row._lookupName = lookupName;
    }

    // 1. CrossRef
    const crossrefData = await fetchCrossRef(row.doi);
    if (crossrefData) {
      applyCrossRef(row, crossrefData);
    }

    // 2. PubMed (교신저자 판별 강화)
    const pubmedData = await fetchPubMed(row.doi);
    if (pubmedData) {
      applyPubMed(row, pubmedData);
    }

    // 3. 교신저자명 컬럼으로 추가 판별
    inferRoleFromCorrNames(row);

    // 4. 판별 결과 확정
    if (!row.inferredRole || row.inferredRole === 'checking') {
      if (row._firstAuthor) {
        // CrossRef에서 제1저자 확인된 경우만 확정
        row.inferredRole = '제1저자';
        row.roleSource = 'CrossRef (저자순서)';
      } else if (row._authorIdx > 0) {
        // 저자 목록에는 있지만 교신저자 정보 없음 → 판별불가
        row.inferredRole = null;
        row.roleSource = null;
      } else {
        // CrossRef에서 저자 매칭 자체 실패 → 판별불가
        row.inferredRole = null;
        row.roleSource = null;
      }
    }

    // 5. 최종 이슈 계산
    computeIssues(row);

  } catch (e) {
    row.status = 'warn';
    row.doiNote = '검증 오류: ' + e.message;
  }
}

/* ── CrossRef API ── */
async function fetchCrossRef(doi) {
  try {
    const res = await fetch(`https://api.crossref.org/works/${encodeURIComponent(doi)}`, {
      headers: { 'User-Agent': 'DOI-Verifier/1.0 (mailto:research@university.ac.kr)' }
    });
    if (!res.ok) return null;
    const d = await res.json();
    return d.message;
  } catch { return null; }
}

function applyCrossRef(row, msg) {
  // 발표일 — published-print 우선, 없으면 published-online
  // (early access / online first는 평가기준 날짜로 인정하지 않음)
  const toDateStr = (dp) => {
    if (!dp || !dp['date-parts'] || !dp['date-parts'][0]) return null;
    const [y, m, d] = dp['date-parts'][0];
    if (!y) return null;
    return [y, m ? String(m).padStart(2,'0') : null, d ? String(d).padStart(2,'0') : null].filter(Boolean).join('-');
  };
  const printDate  = toDateStr(msg['published-print']);
  const onlineDate = toDateStr(msg['published-online']);
  const pubDate    = toDateStr(msg['published']);

  // 평가기준: print > online > published 순서 우선
  row.doiDate = printDate || onlineDate || pubDate || null;
  row.doiPrintDate  = printDate;
  row.doiOnlineDate = onlineDate;
  // 표시용 전체 날짜
  row.doiAllDates = [printDate, onlineDate].filter(Boolean);

  // 저자수
  const authors = msg.author || [];
  row.doiAuthorCount = authors.length;

  // 학술지명
  if (msg['container-title'] && msg['container-title'].length) {
    row.doiJournal = msg['container-title'][0];
  }

  // 논문제목
  if (msg.title && msg.title.length) {
    row.doiTitle = msg.title[0];
  }

  // 저자 분석 (CrossRef 기반 초벌 판별)
  row._crossrefAuthors = authors;

  if (!authors.length) return;

  const name = row.researcherName.toLowerCase();
  const authorIdx = findAuthorIndex(authors, name, row._engName);

  if (authorIdx === -1) {
    // 이름 매칭 실패 → 보수적으로 유지
    row._authorIdx = -1;
    row._totalAuthors = authors.length;
    return;
  }

  row._authorIdx = authorIdx;
  row._totalAuthors = authors.length;
  row._firstAuthor = (authorIdx === 0);

  // CrossRef sequence 기반 초벌 추론
  // (PubMed 또는 교신저자명으로 덮어씌워짐)
  if (authorIdx === 0) {
    row.inferredRole = '제1저자';
    row.roleSource = 'CrossRef (저자순서)';
  } else {
    row.inferredRole = '참여자';
    row.roleSource = 'CrossRef (저자순서)';
  }
}

function findAuthorIndex(authors, nameLower, engName) {
  if (!authors || !authors.length) return -1;

  // 영문명이 있으면 정확히 매칭 (Scopus 기반)
  if (engName && engName.family) {
    const fam = engName.family.toLowerCase();
    const giv = (engName.given || '').toLowerCase();
    for (let i = 0; i < authors.length; i++) {
      const a = authors[i];
      const af = (a.family || '').toLowerCase();
      const ag = (a.given || '').toLowerCase();
      // family 일치 + given 앞글자 일치
      if (af === fam && (ag.startsWith(giv[0] || '') || giv.startsWith(ag[0] || ''))) return i;
      // family만 일치
      if (af === fam) return i;
    }
  }

  // 한글명 fallback (부분 매칭)
  if (nameLower) {
    for (let i = 0; i < authors.length; i++) {
      const a = authors[i];
      const full = ((a.given || '') + (a.family || '')).toLowerCase().replace(/\s+/g, '');
      const fullRev = ((a.family || '') + (a.given || '')).toLowerCase().replace(/\s+/g, '');
      const name = nameLower.replace(/\s+/g, '');
      if (full === name || fullRev === name) return i;
    }
  }
  return -1;
}

/* ── PubMed API ── */
async function fetchPubMed(doi) {
  // 5초 타임아웃 (Scopus 논문 등 PubMed에 없는 경우 무한 대기 방지)
  const timeout = (ms) => new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), ms));
  const fetchWithTimeout = (url) => Promise.race([fetch(url), timeout(5000)]);

  try {
    // Step 1: DOI → PMID
    const searchRes = await fetchWithTimeout(
      `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term=${encodeURIComponent(doi)}[DOI]&retmode=json`
    );
    if (!searchRes.ok) return null;
    const searchData = await searchRes.json();
    const ids = searchData.esearchresult?.idlist || [];
    if (!ids.length) return null;

    // Step 2: PMID → XML (저자 메타데이터)
    const fetchRes = await fetchWithTimeout(
      `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=${ids[0]}&retmode=xml`
    );
    if (!fetchRes.ok) return null;
    const xml = await fetchRes.text();
    return parsePubMedXML(xml);
  } catch { return null; }
}

function parsePubMedXML(xml) {
  const result = { authors: [], hasCorrespondingInfo: false };

  // AuthorList 파싱
  const authorMatches = xml.matchAll(/<Author[^>]*ValidYN="Y"[^>]*>([\s\S]*?)<\/Author>/g);
  for (const m of authorMatches) {
    const block = m[0];
    const lastName  = (block.match(/<LastName>(.*?)<\/LastName>/)  || [])[1] || '';
    const foreName  = (block.match(/<ForeName>(.*?)<\/ForeName>/)  || [])[1] || '';
    const initials  = (block.match(/<Initials>(.*?)<\/Initials>/)  || [])[1] || '';
    const isCorresp = block.includes('EqualContrib="Y"') || /corresp[^"]*"Y"/i.test(block);
    const equalContrib = block.includes('EqualContrib="Y"');
    result.authors.push({ lastName, foreName, initials, isCorresp, equalContrib });
    if (isCorresp || equalContrib) result.hasCorrespondingInfo = true;
  }

  // AffiliationInfo에서 교신저자 표시 추출 (일부 저널)
  // "Corresponding author" 패턴 탐색
  const corrPattern = /[Cc]orrespond\w+\s+author[\s\S]{0,200}?([A-Z][a-z]+[\s\S]{0,50}?[A-Z][a-z]+)/;
  const corrMatch = xml.match(corrPattern);
  if (corrMatch) {
    result.corrAuthorHint = corrMatch[0].slice(0, 100);
    result.hasCorrespondingInfo = true;
  }

  return result;
}

function applyPubMed(row, pmData) {
  if (!pmData || !pmData.authors.length) return;

  const name = row.researcherName.toLowerCase().replace(/\s/g, '');
  const authors = pmData.authors;
  const engName = row._engName;

  // 해당 연구자 찾기 (Scopus 영문명 우선)
  let myIdx = -1;
  for (let i = 0; i < authors.length; i++) {
    const a = authors[i];
    if (engName && engName.family) {
      if (a.lastName.toLowerCase() === engName.family.toLowerCase()) { myIdx = i; break; }
    }
    const full = (a.lastName + a.foreName).toLowerCase().replace(/\s/g,'');
    const fullRev = (a.foreName + a.lastName).toLowerCase().replace(/\s/g,'');
    if (full === name || fullRev === name || name.includes(a.lastName.toLowerCase())) {
      myIdx = i; break;
    }
  }
  if (myIdx === -1) return;

  const me = authors[myIdx];
  const corrCount = authors.filter(a => a.isCorresp || a.equalContrib).length;
  const isFirst = myIdx === 0;
  const isCorr = me.isCorresp || me.equalContrib;

  row.roleSource = 'PubMed';
  row._pmAuthorIdx = myIdx;
  row._pmCorrCount = corrCount;

  if (isFirst && isCorr) {
    row.inferredRole = '교신저자&제1저자';
  } else if (isFirst) {
    row.inferredRole = '제1저자';
  } else if (isCorr && corrCount > 1) {
    row.inferredRole = '공동교신저자';
  } else if (isCorr) {
    row.inferredRole = '교신저자';
  } else {
    row.inferredRole = '참여자';
  }
}

/* ── 엑셀 내부 데이터 기반 참여형태 판별 (핵심 로직) ── */
function inferRoleFromCorrNames(row) {
  const targetName = getTargetName();
  const myName = (targetName || row.researcherName).replace(/\s/g, '');
  if (!myName) return;

  // 교신저자명 파싱: "윤종승, 선양국" 형태
  const corrList = row.corrNames
    ? row.corrNames.split(/[,，、]+/).map(s => s.trim()).filter(Boolean)
    : [];

  const nameMatch = (a, b) => {
    const na = a.replace(/\s/g, '');
    const nb = b.replace(/\s/g, '');
    return na === nb || na.includes(nb) || nb.includes(na);
  };

  const iAmInCorrList = corrList.some(c => nameMatch(c, myName));

  // 교신저자수 컬럼 우선, 없으면 교신저자명 목록 수로 대체
  const corrCount = row.corrCount > 0 ? row.corrCount : corrList.length;
  const isFirst = row._firstAuthor || row._pmAuthorIdx === 0;

  if (iAmInCorrList) {
    // 교신저자명에 본인 포함 → 교신저자 계열 판별
    if (corrCount >= 2) {
      row.inferredRole = '공동교신저자';
    } else if (isFirst) {
      row.inferredRole = '교신저자&제1저자';
    } else {
      row.inferredRole = '교신저자';
    }
    row.roleSource = '교신저자명 컬럼';

  } else if (corrList.length > 0) {
    // 교신저자명이 있는데 본인이 없음 → 제1저자 또는 참여자
    if (isFirst) {
      row.inferredRole = '제1저자';
      row.roleSource = 'CrossRef (저자순서)';
    } else {
      row.inferredRole = '참여자';
      row.roleSource = 'CrossRef (저자순서)';
    }

  } else {
    // 교신저자명 컬럼이 비어있음
    // → 원본 참여형태(엑셀 기재값)를 그대로 신뢰하고 불일치 판정 안 함
    row.inferredRole = row.role || null;
    row.roleSource = '원본신뢰 (교신자명 미기재)';
  }
}

/* ── 이슈 계산 ── */
function computeIssues(row) {
  row.issues = [];
  row.roleIssue = false;

  // 논문제목 비교
  if (row.doiTitle && row.title) {
    if (!titleMatch(row.title, row.doiTitle)) {
      row.issues.push('논문제목');
    }
  }

  // 업적평가 기간 검증 (published-print 기준, 없으면 online)
  const evalDate = row.doiPrintDate || row.doiOnlineDate || row.doiDate;
  if (evalDate) {
    const pr = checkPeriod(evalDate);
    row.periodResult = pr;
    if (pr && !pr.inPeriod) {
      row.issues.push('평가기간외');
    }
  }

  // 저자수 비교
  if (row.doiAuthorCount !== null && row.authorCount > 0) {
    if (row.doiAuthorCount !== row.authorCount) {
      row.issues.push('저자수');
    }
  }

  // 발표일 비교는 제거 (업적평가 기간으로 대체)

  // 학술지명 비교
  if (row.doiJournal && row.journal) {
    if (!journalMatch(row.journal, row.doiJournal)) {
      row.issues.push('학술지명');
    }
  }

  // 참여형태 비교
  if (row.inferredRole && row.inferredRole !== 'checking') {
    if (!roleMatch(row.role, row.inferredRole)) {
      row.roleIssue = true;
    }
  }

  // 최종 상태 결정
  if (row.issues.length > 0) {
    row.status = 'fail';
  } else if (row.roleIssue) {
    row.status = 'role_only';
  } else if (!row.doiDate && !row.doiJournal) {
    row.status = 'warn';
    if (!row.doiNote) row.doiNote = 'DOI 조회 실패';
  } else {
    row.status = 'ok';
  }
}

function titleMatch(a, b) {
  const n = s => s.toLowerCase().replace(/[^a-z0-9가-힣]/g, '');
  const na = n(a), nb = n(b);
  if (na === nb) return true;
  // 70% 이상 토큰 일치
  return tokenSim(na, nb) > 0.7;
}

function journalMatch(a, b) {
  const n = s => s.toLowerCase().replace(/[^a-z0-9가-힣]/g, '');
  const na = n(a), nb = n(b);
  if (na === nb) return true;
  if (na.includes(nb) || nb.includes(na)) return true;
  // 약어 매칭 (J ALLOY COMPD ↔ Journal of Alloys and Compounds)
  return tokenSim(na, nb) > 0.6;
}

function tokenSim(a, b) {
  const ta = new Set(a.match(/[a-z0-9가-힣]{2,}/g) || []);
  const tb = new Set(b.match(/[a-z0-9가-힣]{2,}/g) || []);
  let inter = 0;
  ta.forEach(t => { if (tb.has(t)) inter++; });
  return inter / Math.max(ta.size, tb.size, 1);
}

function roleMatch(recorded, inferred) {
  if (!recorded || !inferred) return true; // 판별 불가 → 불일치 미표시
  // 원본 신뢰 케이스 → 불일치 판정 안 함
  if (inferred === recorded) return true;
  if (inferred.includes('불명')) return true;
  const normalize = s => s.toLowerCase()
    .replace(/\s+/g,'')
    .replace('&', '')
    .replace('제1저자', '1저자')
    .replace('제일저자', '1저자')
    .replace('1저자', 'first');
  const a = normalize(recorded);
  const b = normalize(inferred);
  if (a === b) return true;
  if (a.includes('교신') && b.includes('교신')) return true;
  if (a.includes('first') && b.includes('first')) return true;
  if (a === '참여자' && b.includes('참여자')) return true;
  return false;
}

/* ── 통계 갱신 ── */
function updateStats() {
  const total   = rows.length;
  const ok      = rows.filter(r => r.status === 'ok').length;
  const fail    = rows.filter(r => r.status === 'fail').length;
  const warn    = rows.filter(r => ['warn','pending','checking'].includes(r.status)).length;
  const role    = rows.filter(r => r.roleIssue).length;
  const period  = rows.filter(r => r.periodResult && !r.periodResult.inPeriod).length;
  const authors = rows.filter(r => r.doiAuthorCount !== null && r.authorCount > 0 && r.doiAuthorCount !== r.authorCount).length;
  document.getElementById('s-total').textContent = total;
  document.getElementById('s-ok').textContent = ok;
  document.getElementById('s-fail').textContent = fail;
  document.getElementById('s-warn').textContent = warn;
  document.getElementById('s-role').textContent = role;
  const ps = document.getElementById('s-period');
  const as = document.getElementById('s-authors');
  if (ps) ps.textContent = period;
  if (as) as.textContent = authors;
}

/* ── 테이블 렌더링 ── */
function renderTable() {
  const filter  = document.getElementById('filterStatus').value;
  const search  = document.getElementById('searchInput').value.toLowerCase();

  let filtered = rows.filter(r => {
    if (filter === 'role_only' && !r.roleIssue) return false;
    if (filter === 'period_out' && !(r.periodResult && !r.periodResult.inPeriod)) return false;
    if (filter === 'author_mismatch' && !(r.doiAuthorCount !== null && r.authorCount > 0 && r.doiAuthorCount !== r.authorCount)) return false;
    if (filter && !['role_only','period_out','author_mismatch'].includes(filter) && r.status !== filter) return false;
    if (search && !r.title.toLowerCase().includes(search) && !r.doi.toLowerCase().includes(search)) return false;
    return true;
  });

  document.getElementById('rowCount').textContent = `${filtered.length}개`;

  if (!filtered.length) {
    document.getElementById('tbody').innerHTML =
      `<tr><td colspan="9" style="text-align:center;padding:3rem;color:var(--text3);font-size:13px">조건에 맞는 항목이 없습니다</td></tr>`;
    return;
  }

  document.getElementById('tbody').innerHTML = filtered.map(row => renderRow(row)).join('');
}

function renderRow(r) {
  const rowClass = r.status === 'fail' ? 'row-fail' : r.roleIssue ? 'row-role' : '';

  // 평가기간 셀
  let periodCell = '—';
  if (r.periodResult) {
    const pr = r.periodResult;
    const cls = pr.inPeriod ? 'ok' : 'fail';
    const icon = pr.inPeriod ? '✓' : '✗';
    periodCell = `<div style="font-size:11px">
      <div class="period-badge ${cls}">${icon} ${pr.inPeriod ? '기간 내' : '기간 외'}</div>
      <div style="color:var(--text3);margin-top:2px">${pr.date}</div>
      <div style="color:var(--text3)">${pr.period}</div>
    </div>`;
  } else if (r.doiDate) {
    periodCell = `<span class="period-badge none">미설정</span>`;
  }

  // 저자수 셀
  const authorMismatch = r.issues.includes('저자수');
  const authorCell = `<div class="cell-compare">
    <div class="cell-orig">${r.authorCount || '—'}</div>
    <div class="cell-doi-val ${authorMismatch ? 'mismatch' : r.doiAuthorCount ? 'match' : ''}">
      ${r.doiAuthorCount !== null ? '▸ ' + r.doiAuthorCount : r.status === 'checking' ? '…' : '—'}
    </div>
    ${authorMismatch ? '<span class="mismatch-label">불일치</span>' : ''}
  </div>`;

  // 발표일
  const dateMismatch = r.issues.includes('발표일');
  const dateCell = `
    <div class="cell-compare">
      <div class="cell-orig">${r.date || '—'}</div>
      <div class="cell-doi-val ${dateMismatch ? 'mismatch' : r.doiDate ? 'match' : ''}">
        ${r.doiDate ? '▸ ' + r.doiDate : r.status === 'checking' ? '…' : '—'}
      </div>
      ${dateMismatch ? '<span class="mismatch-label">불일치</span>' : ''}
    </div>`;

  // 학술지명
  const journalMismatch = r.issues.includes('학술지명');
  const journalCell = `
    <div class="cell-compare">
      <div class="cell-orig" title="${esc(r.journal)}">${trunc(r.journal, 22)}</div>
      <div class="cell-doi-val ${journalMismatch ? 'mismatch' : r.doiJournal ? 'match' : ''}" title="${esc(r.doiJournal||'')}">
        ${r.doiJournal ? '▸ ' + trunc(r.doiJournal, 22) : r.status === 'checking' ? '…' : '—'}
      </div>
      ${journalMismatch ? '<span class="mismatch-label">불일치</span>' : ''}
    </div>`;

  // 참여형태
  const roleCell = renderRoleCell(r);

  // 상태 배지
  const statusLabels = {
    ok: '✓ 일치', fail: '✗ 불일치', warn: '미검증',
    role_only: '⚑ 형태불일치', pending: '대기', checking: '조회중'
  };
  const statusBadge = `<span class="status-badge ${r.status}">${statusLabels[r.status] || r.status}</span>`;

  // 이슈 목록
  let issuesHtml = '<div class="issues-list">';
  r.issues.forEach(i => issuesHtml += `<div class="issue-item">✗ ${i}</div>`);
  if (r.roleIssue) issuesHtml += `<div class="issue-role">⚑ 참여형태</div>`;
  if (!r.issues.length && !r.roleIssue && r.doiNote)
    issuesHtml += `<div class="issue-note">${r.doiNote}</div>`;
  if (!r.issues.length && !r.roleIssue && !r.doiNote && r.status === 'ok')
    issuesHtml += `<div style="color:var(--ok);font-size:11px">이상 없음</div>`;
  issuesHtml += '</div>';

  // 논문제목 (특수문자 안전 처리 — 템플릿 리터럴 밖에서 모두 escape)
  const titleMismatch = r.issues.includes('논문제목');
  const t_orig = esc(trunc(r.title, 40));
  const t_orig_full = esc(r.title);
  const t_doi = r.doiTitle ? esc(trunc(r.doiTitle, 40)) : '';
  const t_doi_full = esc(r.doiTitle || '');
  const t_status = r.status === 'checking' ? '…' : '';
  const titleCell = '<div class="cell-compare">'
    + '<div class="paper-title ' + (titleMismatch ? 'mismatch' : '') + '" title="' + t_orig_full + '">' + t_orig + '</div>'
    + '<div class="cell-doi-val ' + (titleMismatch ? 'mismatch' : t_doi ? 'match' : '') + '" title="' + t_doi_full + '">'
    + (t_doi ? '&#9658; ' + t_doi : t_status)
    + '</div>'
    + (titleMismatch ? '<span class="mismatch-label">불일치</span>' : '')
    + '</div>';

  return `<tr class="${rowClass}">
    <td class="col-no"><span class="cell-no">${r.idx}</span></td>
    <td class="col-title">${titleCell}</td>
    <td class="col-doi">
      ${r.doi ? `<a class="doi-link" href="https://doi.org/${r.doi}" target="_blank" rel="noopener">${r.doi}</a>` : '—'}
    </td>
    <td class="col-date">${dateCell}</td>
    <td class="col-period">${periodCell}</td>
    <td class="col-authors">${authorCell}</td>
    <td class="col-journal">${journalCell}</td>
    <td class="col-role">${roleCell}</td>
    <td class="col-corr"><div style="font-size:11px;color:var(--text3)">${r.corrNames || '—'}</div></td>
    <td class="col-status">${statusBadge}</td>
    <td class="col-issues">${issuesHtml}</td>
  </tr>`;
}

function renderRoleCell(r) {
  const orig = r.role || '—';
  const inferred = r.inferredRole;
  const source = r.roleSource || '';
  const mismatch = r.roleIssue;

  const roleClass = getRoleBadgeClass(inferred);
  const origClass = getRoleBadgeClass(orig);

  const origBadge  = `<span class="role-badge ${origClass}">${orig}</span>`;
  const isUnknown = inferred && inferred.includes('불명');
  const isTrustedRole = source && source.includes('원본신뢰');
  const inferBadge = inferred && inferred !== 'checking'
    ? `<span class="role-badge ${isTrustedRole ? origClass : (isUnknown ? 'unknown' : roleClass)}" style="opacity:${isTrustedRole ? '0.75' : '1'}">${inferred} ${isTrustedRole ? '<span style="font-size:9px;opacity:0.7">(원본)</span>' : ''}</span>`
    : inferred === 'checking'
    ? `<span class="role-badge checking">조회중…</span>`
    : `<span style="font-size:11px;color:var(--text3);background:var(--bg3);padding:2px 8px;border-radius:20px;">판별불가</span>`;

  const isTrusted = source && source.includes('원본신뢰');
  const sourceLabel = source
    ? `<span style="font-size:10px;color:${isTrusted ? 'var(--warn)' : 'var(--text3)'}">${source}</span>`
    : `<span style="font-size:10px;color:var(--text3)">교신저자명 없음</span>`;

  const mismatchFlag = mismatch
    ? `<span class="role-mismatch-flag">⚑ 불일치</span>` : '';

  return `
    <div class="role-cell">
      <div>${origBadge}</div>
      <div class="role-inferred">
        <span class="role-arrow">▸</span>
        ${inferBadge}
      </div>
      ${sourceLabel}
      ${mismatchFlag}
    </div>`;
}

function getRoleBadgeClass(role) {
  if (!role) return 'unknown';
  const r = role.toLowerCase().replace(/\s/g,'');
  if (r.includes('교신저자&제1') || r.includes('교신저자&1') || r.includes('1저자&교신')) return 'first-corr';
  if (r.includes('공동교신')) return 'co-corr';
  if (r.includes('교신')) return 'corr';
  if (r.includes('1저자') || r.includes('제1')) return 'first';
  if (r.includes('참여')) return 'participant';
  return 'unknown';
}

function trunc(s, n) {
  if (!s) return '—';
  return s.length > n ? s.slice(0, n) + '…' : s;
}

function esc(s) {
  return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ── CSV 내보내기 ── */
function exportCSV() {
  const headers = ['#','논문제목','DOI','최초출판일(DOI)','평가기간포함여부','저자수(원본)','저자수(DOI)','학술지명(원본)','학술지명(DOI조회)',
    '참여형태(원본)','참여형태(판별)','판별근거','교신저자명(원본)','검증결과','불일치항목'];
  const lines = [headers.join(',')];
  rows.forEach(r => {
    const issues = [...r.issues, ...(r.roleIssue ? ['참여형태'] : [])].join('/');
    const statusMap = {ok:'일치',fail:'불일치',warn:'미검증',role_only:'참여형태불일치',pending:'대기',checking:'조회중'};
    lines.push([
      r.idx,
      `"${r.title.replace(/"/g,'""')}"`,
      r.doi,
      r.doiDate || '',
      r.periodResult ? (r.periodResult.inPeriod ? '기간내' : '기간외') : '',
      r.authorCount || '',
      r.doiAuthorCount !== null ? r.doiAuthorCount : '',
      `"${r.journal.replace(/"/g,'""')}"`,
      `"${(r.doiJournal||'').replace(/"/g,'""')}"`,
      r.role,
      r.inferredRole || '',
      r.roleSource || '',
      `"${r.corrNames.replace(/"/g,'""')}"`,
      statusMap[r.status] || r.status,
      issues
    ].join(','));
  });

  const bom = '\uFEFF';
  const blob = new Blob([bom + lines.join('\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `DOI_검증결과_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

/* ── 헬프 패널 ── */
function toggleHelp() {
  const panel = document.getElementById('helpPanel');
  panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}
