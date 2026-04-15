/* ════════════════════════════════════════════
   연구실적 DOI 검증 시스템 — app.js
   ════════════════════════════════════════════ */

let rows = [];
let verifying = false;
let verifyQueue = [];

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
    researcherName: str(r[colMap.name]),
    status: 'pending',
    doiDate: null, doiJournal: null,
    inferredRole: null, roleSource: null,
    issues: [], roleIssue: false, doiNote: ''
  }));

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
    corrNames: findExact('교신저자명') || find('교신저자명'),
    name:      findExact('성명') || find('성명', 'name'),
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

    // 4. 최종 이슈 계산
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
  // 발표일
  const dp = msg['published-print'] || msg['published-online'] || msg['published'] || msg['created'];
  if (dp && dp['date-parts'] && dp['date-parts'][0]) {
    const [y, m, d] = dp['date-parts'][0];
    row.doiDate = [y, m ? String(m).padStart(2,'0') : null, d ? String(d).padStart(2,'0') : null]
      .filter(Boolean).join('-');
  }

  // 학술지명
  if (msg['container-title'] && msg['container-title'].length) {
    row.doiJournal = msg['container-title'][0];
  }

  // 저자 분석 (CrossRef 기반 초벌 판별)
  const authors = msg.author || [];
  row._crossrefAuthors = authors;

  if (!authors.length) return;

  const name = row.researcherName.toLowerCase();
  const authorIdx = findAuthorIndex(authors, name);

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

function findAuthorIndex(authors, nameLower) {
  if (!nameLower) return -1;
  const parts = nameLower.replace(/\s+/g, '').split('');
  for (let i = 0; i < authors.length; i++) {
    const a = authors[i];
    const full = ((a.given || '') + (a.family || '')).toLowerCase().replace(/\s+/g, '');
    const fullRev = ((a.family || '') + (a.given || '')).toLowerCase().replace(/\s+/g, '');
    if (full === parts.join('') || fullRev === parts.join('')) return i;
    // 한국어 이름: family가 성씨, given이 이름인 경우
    const korean = (a.family || '').replace(/\s/g,'') + (a.given || '').replace(/\s/g,'');
    if (korean.toLowerCase() === nameLower.replace(/\s/g,'')) return i;
    // 부분 매칭 (성씨만)
    if (a.family && nameLower.includes(a.family.toLowerCase())) {
      return i; // 소극적 매칭
    }
  }
  return -1;
}

/* ── PubMed API ── */
async function fetchPubMed(doi) {
  try {
    // Step 1: DOI → PMID
    const searchRes = await fetch(
      `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term=${encodeURIComponent(doi)}[DOI]&retmode=json`
    );
    if (!searchRes.ok) return null;
    const searchData = await searchRes.json();
    const ids = searchData.esearchresult?.idlist || [];
    if (!ids.length) return null;

    // Step 2: PMID → XML (저자 메타데이터)
    const fetchRes = await fetch(
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

  // 해당 연구자 찾기
  let myIdx = -1;
  for (let i = 0; i < authors.length; i++) {
    const a = authors[i];
    const full = (a.lastName + a.foreName).toLowerCase().replace(/\s/g,'');
    const fullRev = (a.foreName + a.lastName).toLowerCase().replace(/\s/g,'');
    if (full === name || fullRev === name || name.includes(a.lastName.toLowerCase())) {
      myIdx = i;
      break;
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

/* ── 교신저자명 컬럼 기반 판별 (엑셀 내부 정보 활용) ── */
function inferRoleFromCorrNames(row) {
  if (!row.corrNames) return;

  const myName = row.researcherName.toLowerCase().replace(/\s/g,'');
  if (!myName) return;

  // 교신저자명 파싱: "윤종승, 선양국" 형태
  const corrList = row.corrNames
    .split(/[,，、\s]+/)
    .map(s => s.trim().replace(/\s/g,'').toLowerCase())
    .filter(Boolean);

  if (!corrList.length) return;

  const iAmCorr = corrList.some(c => c.includes(myName) || myName.includes(c));
  if (!iAmCorr) return;

  const isFirst = row._firstAuthor || row._pmAuthorIdx === 0;
  const isMultiCorr = corrList.length > 1;

  let inferred;
  if (isFirst && !isMultiCorr) {
    inferred = '교신저자&제1저자';
  } else if (isMultiCorr) {
    inferred = '공동교신저자';
  } else {
    inferred = '교신저자';
  }

  // 교신저자명 컬럼은 신뢰도 높음 → 덮어씌우기
  row.inferredRole = inferred;
  row.roleSource = '교신저자명 컬럼';
}

/* ── 이슈 계산 ── */
function computeIssues(row) {
  row.issues = [];
  row.roleIssue = false;

  // 발표일 비교 (연월 기준)
  if (row.doiDate && row.date) {
    const a = row.date.replace(/-/g,'').slice(0, 6);
    const b = row.doiDate.replace(/-/g,'').slice(0, 6);
    if (a !== b) row.issues.push('발표일');
  }

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
  if (!recorded || !inferred) return true; // 판별 불가 → 패스
  const normalize = s => s.toLowerCase()
    .replace(/\s+/g,'')
    .replace('&', '')
    .replace('제1저자', '1저자')
    .replace('제일저자', '1저자');
  const a = normalize(recorded);
  const b = normalize(inferred);
  if (a === b) return true;
  // 부분 허용
  if (a.includes('교신') && b.includes('교신')) return true;
  if (a.includes('1저자') && b.includes('1저자')) return true;
  if (a === '참여자' && b === '참여자') return true;
  return false;
}

/* ── 통계 갱신 ── */
function updateStats() {
  const total = rows.length;
  const ok    = rows.filter(r => r.status === 'ok').length;
  const fail  = rows.filter(r => r.status === 'fail').length;
  const warn  = rows.filter(r => ['warn','pending','checking'].includes(r.status)).length;
  const role  = rows.filter(r => r.roleIssue).length;
  document.getElementById('s-total').textContent = total;
  document.getElementById('s-ok').textContent = ok;
  document.getElementById('s-fail').textContent = fail;
  document.getElementById('s-warn').textContent = warn;
  document.getElementById('s-role').textContent = role;
}

/* ── 테이블 렌더링 ── */
function renderTable() {
  const filter  = document.getElementById('filterStatus').value;
  const search  = document.getElementById('searchInput').value.toLowerCase();

  let filtered = rows.filter(r => {
    if (filter === 'role_only' && !r.roleIssue) return false;
    if (filter && filter !== 'role_only' && r.status !== filter) return false;
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

  return `<tr class="${rowClass}">
    <td class="col-no"><span class="cell-no">${r.idx}</span></td>
    <td class="col-title"><div class="paper-title" title="${esc(r.title)}">${esc(r.title) || '—'}</div></td>
    <td class="col-doi">
      ${r.doi ? `<a class="doi-link" href="https://doi.org/${r.doi}" target="_blank" rel="noopener">${r.doi}</a>` : '—'}
    </td>
    <td class="col-date">${dateCell}</td>
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
  const inferBadge = inferred && inferred !== 'checking'
    ? `<span class="role-badge ${roleClass}">${inferred}</span>`
    : inferred === 'checking'
    ? `<span class="role-badge checking">조회중…</span>`
    : `<span style="font-size:11px;color:var(--text3)">—</span>`;

  const sourceLabel = source
    ? `<span style="font-size:10px;color:var(--text3)">${source}</span>` : '';

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
  const headers = ['#','논문제목','DOI','발표일(원본)','발표일(DOI조회)','학술지명(원본)','학술지명(DOI조회)',
    '참여형태(원본)','참여형태(판별)','판별근거','교신저자명(원본)','검증결과','불일치항목'];
  const lines = [headers.join(',')];
  rows.forEach(r => {
    const issues = [...r.issues, ...(r.roleIssue ? ['참여형태'] : [])].join('/');
    const statusMap = {ok:'일치',fail:'불일치',warn:'미검증',role_only:'참여형태불일치',pending:'대기',checking:'조회중'};
    lines.push([
      r.idx,
      `"${r.title.replace(/"/g,'""')}"`,
      r.doi,
      r.date,
      r.doiDate || '',
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
