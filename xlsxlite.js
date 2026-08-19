// xlsxlite: 값만 읽는 초경량 xlsx 리더 (한컴 hs: 태그·비표준 styles 무시)
// 브라우저(JSZip)와 Node에서 동일 동작. XML은 정규식 기반(값 추출 목적에 충분).
(function(root){
function unesc(s){
  return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"')
          .replace(/&apos;/g,"'").replace(/&#x([0-9a-fA-F]+);/g,(m,h)=>String.fromCodePoint(parseInt(h,16)))
          .replace(/&#(\d+);/g,(m,d)=>String.fromCodePoint(parseInt(d,10))).replace(/&amp;/g,'&');
}
function parseSharedStrings(xml){
  const out = [];
  if (!xml) return out;
  const siRe = /<(?:\w+:)?si[\s>][\s\S]*?<\/(?:\w+:)?si>|<(?:\w+:)?si\/>/g;
  let m;
  while ((m = siRe.exec(xml))){
    const si = m[0];
    let text = '';
    const tRe = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g;
    let t;
    while ((t = tRe.exec(si))) text += unesc(t[1]);
    out.push(text);
  }
  return out;
}
function colToIdx(ref){ // "AB12" -> col index 0-based
  let c = 0;
  for (const ch of ref){
    if (ch >= 'A' && ch <= 'Z') c = c*26 + (ch.charCodeAt(0)-64);
    else break;
  }
  return c-1;
}
function parseSheet(xml, shared){
  const rows = [];
  const rowRe = /<(?:\w+:)?row\b[^>]*>([\s\S]*?)<\/(?:\w+:)?row>/g;
  let rm;
  while ((rm = rowRe.exec(xml))){
    const rowXml = rm[0];
    const rNum = (rowXml.match(/^<(?:\w+:)?row\b[^>]*\br="(\d+)"/)||[])[1];
    const idx = rNum ? parseInt(rNum,10)-1 : rows.length;
    const cells = [];
    const cRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let cm;
    while ((cm = cRe.exec(rm[1]))){
      const attrs = cm[1], inner = cm[2] || '';
      const ref = (attrs.match(/\br="([A-Z]+)\d+"/)||[])[1];
      const type = (attrs.match(/\bt="([^"]+)"/)||[])[1] || 'n';
      const ci = ref ? colToIdx(ref) : cells.length;
      let val = '';
      if (type === 's'){
        const v = (inner.match(/<(?:\w+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?v>/)||[])[1];
        val = v != null ? (shared[parseInt(v,10)] ?? '') : '';
      } else if (type === 'inlineStr'){
        let text=''; const tRe=/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g; let t;
        while ((t=tRe.exec(inner))) text += unesc(t[1]);
        val = text;
      } else { // n, str, b, e
        const v = (inner.match(/<(?:\w+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?v>/)||[])[1];
        if (v == null) val = '';
        else if (type === 'n') { const num = Number(v); val = isNaN(num) ? unesc(v) : num; }
        else val = unesc(v);
      }
      cells[ci] = val;
    }
    if (cells.length) rows[idx] = cells;
  }
  // 빈 행 채우기
  for (let i=0;i<rows.length;i++) if (!rows[i]) rows[i] = [];
  return rows;
}
function parseWorkbookMap(wbXml, relsXml){
  const rels = {};
  const relRe = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>|<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"[^>]*\/>/g;
  let rm;
  while ((rm = relRe.exec(relsXml))){
    const id = rm[1] || rm[4], target = rm[2] || rm[3];
    rels[id] = target.replace(/^\//,'').replace(/^xl\//,'');
  }
  const sheets = [];
  const shRe = /<(?:\w+:)?sheet\b([^>]*)\/>/g;
  let sm;
  while ((sm = shRe.exec(wbXml))){
    const a = sm[1];
    const name = unesc((a.match(/\bname="([^"]*)"/)||[])[1]||'');
    const rid = (a.match(/\br:id="([^"]*)"/)||[])[1] || (a.match(/\bid="([^"]*)"/)||[])[1];
    sheets.push({ name, path: 'xl/' + (rels[rid] || '') });
  }
  return sheets;
}
// entries: {경로: xml문자열} — zip 해제는 호출측(JSZip 등)에서
function readWorkbook(entries){
  const shared = parseSharedStrings(entries['xl/sharedStrings.xml'] || '');
  const sheets = parseWorkbookMap(entries['xl/workbook.xml'] || '', entries['xl/_rels/workbook.xml.rels'] || '');
  return {
    sheetNames: sheets.map(s=>s.name),
    getRows(name){
      const s = sheets.find(x=>x.name===name);
      if (!s || !entries[s.path]) return [];
      return parseSheet(entries[s.path], shared);
    }
  };
}
const api = { readWorkbook, parseSharedStrings, parseSheet };
if (typeof module !== 'undefined') module.exports = api;
else root.xlsxlite = api;
})(typeof self !== 'undefined' ? self : globalThis);
