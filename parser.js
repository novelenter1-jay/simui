// ===== 공용 파서 (admin.html에 동일 로직 탑재) =====
function excelSerialToDate(n){ // Excel 1900 system
  const ms = Math.round((n - 25569) * 86400 * 1000);
  return new Date(ms);
}
function toNum(v){ // 숫자 또는 날짜로 잘못 저장된 숫자 복원
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  if (v instanceof Date) { // cellDates:true일 때 대비
    return Math.round(v.getTime()/86400000 + 25569);
  }
  const s = String(v).trim();
  if (/^\d+(\.0+)?$/.test(s)) return parseInt(s,10);
  return null;
}
function fmtDate(serial){
  const d = excelSerialToDate(serial);
  const y=d.getUTCFullYear(), m=d.getUTCMonth()+1, dd=d.getUTCDate();
  return `${y}-${String(m).padStart(2,'0')}-${String(dd).padStart(2,'0')}`;
}
function fmtTime(v){
  if (v == null || v === '') return '';
  if (typeof v === 'number'){ // time fraction or datetime serial
    let f = v % 1; if (v < 1) f = v;
    const mins = Math.round(f*24*60);
    return `${String(Math.floor(mins/60)).padStart(2,'0')}:${String(mins%60).padStart(2,'0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/(\d{1,2})\s*[:시]\s*(\d{0,2})/);
  if (m) return `${m[1].padStart(2,'0')}:${(m[2]||'00').padStart(2,'0')}`;
  return s;
}
function roomOf(simui){ // 심의번호에서 (늘)/(누)
  const s = String(simui||'');
  if (s.includes('늘')) return '늘품실';
  if (s.includes('누')) return '누리실';
  return '';
}

// ===== 월별 일정 시트 파싱 =====
function parseScheduleSheet(rows, sheetName){
  // 헤더 행 찾기
  let hi = -1, col = {};
  for (let i=0;i<Math.min(rows.length,8);i++){
    const r = rows[i].map(c=>String(c||'').replace(/\s/g,''));
    if (r.includes('심의번호')){
      hi = i;
      r.forEach((h,j)=>{
        if (h==='심의날짜') col.date=j;
        else if (h==='요일') col.dow=j;
        else if (h==='소위원회') col.sub=j;
        else if (h==='간사') col.gansa=j;
        else if (h==='회차') col.round=j;
        else if (h==='심의번호') col.simui=j;
        else if (h==='시간') col.time=j;
        else if (h==='학교') col.school=j;
        else if (h==='급') col.level=j;
      });
      break;
    }
  }
  if (hi < 0) return [];
  const out = [];
  let lastDate=null, lastDow='';
  for (let i=hi+1;i<rows.length;i++){
    const r = rows[i];
    const get = k => col[k]!=null ? r[col[k]] : '';
    const dser = toNum(get('date'));
    if (dser && dser > 40000) lastDate = dser;             // 병합셀 대비 carry-down
    const dow = String(get('dow')||'').trim();
    if (dow) lastDow = dow;
    const simui = String(get('simui')||'').trim();
    if (!simui || !/\d{4}\s*-\s*\d+/.test(simui)) continue; // 심의번호 있는 행만
    let school = String(get('school')||'').replace(/\s+/g,' ').trim();
    school = school.replace(/0\d{1,2}[-.\s]?\d{3,4}[-.\s]?\d{4}/g,'').replace(/\s{2,}/g,' ').trim(); // 전화번호 자동 제거
    if (!school) continue;
    const sub = toNum(get('sub'));
    const round = toNum(get('round'));
    const id = simui.match(/(\d{4})\s*-\s*(\d+)/);
    out.push({
      id: `${id[1]}-${id[2]}`,                 // Firestore 문서 ID: 2026-426
      date: lastDate ? fmtDate(lastDate) : '',
      dow: lastDow,
      sub: sub || 0,
      gansa: String(get('gansa')||'').replace(/\s+/g,' ').trim(),
      round: round || null,
      simui: simui,
      room: roomOf(simui),
      time: fmtTime(get('time')),
      school: school,
      level: String(get('level')||'').trim(),
      cancelled: /취소/.test(school),
      month: sheetName.replace(/\s/g,''),
    });
  }
  return out;
}

// ===== 소위별 명단 시트 파싱 (성명·구분만, 연락처/이메일 제외) =====
function parseRosterSheet(rows){
  const subs = {}; // {1: {sub:1, gansa, dow, members:[{name, role, chair}]}}
  let cur = null;
  for (const r of rows){
    const cells = r.map(c=>String(c||'').trim());
    const joined = cells.join(' ');
    const m = joined.match(/제\s*(\d+)\s*소위원회/);
    let subNo = m ? parseInt(m[1],10) : (cur ? cur.sub : null);
    // 성명 위치 추정: 헤더에서 열 찾기보다 고정 열(구성 파일 기준) 사용 + 검증
    // 열: 0구분 1소위 2위원구분 3위원구분2 4소속 5성별 6성명 7연락처 8이메일 9간사 10요일
    if (!m && !cells[6]) continue;
    if (m){
      if (!subs[subNo]) subs[subNo] = { sub: subNo, gansa:'', dow:'', members:[] };
      cur = subs[subNo];
    }
    if (!cur) continue;
    const name = cells[6];
    if (!name || name==='성명') continue;
    const role2 = cells[3] || cells[2];
    cur.members.push({ name, role: cells[2]||'', chair: /소위원장/.test(role2) });
    if (cells[9] && cells[9]!=='간사') cur.gansa = cells[9];
    if (cells[10] && cells[10]!=='요일') cur.dow = cells[10];
  }
  return Object.values(subs).sort((a,b)=>a.sub-b.sub);
}

const _api = { parseScheduleSheet, parseRosterSheet, toNum, fmtDate, fmtTime };
if (typeof module !== 'undefined') module.exports = _api;
if (typeof window !== 'undefined') window.svparser = _api;
