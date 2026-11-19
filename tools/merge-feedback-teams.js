// 팀피드백.html 의 TEAMS 배열을 '코드로 검증된' 내용으로 바꿉니다.
// 사용: node merge_teams.js <팀피드백.html> <fb_verify 폴더>
// - id · members · tags · links 는 기존 값을 유지, 나머지(one · summary · features · risks · strengths · improves · mission · questions · scores)는 검증본으로 교체
// - team5.fixed.json 이 있으면 새 팀으로 추가 (표지 소개 · 태그 · 링크 비움)
const fs = require('fs'), path = require('path');
const [html, dir] = process.argv.slice(2);
let s = fs.readFileSync(html, 'utf8');
const a = s.indexOf('var TEAMS = ['), b = s.indexOf('var CRIT');
if (a < 0 || b < 0) throw new Error('TEAMS 위치를 찾지 못함');
const old = new Function(s.slice(a, b) + '; return TEAMS;')();
const KEYS = ['one', 'summary', 'features', 'risks', 'strengths', 'improves', 'mission', 'questions', 'scores'];
const out = old.map(t => {
  const f = path.join(dir, t.id + '.fixed.json');
  if (!fs.existsSync(f)) { console.log('검증본 없음, 그대로:', t.id); return t; }
  const v = JSON.parse(fs.readFileSync(f, 'utf8'));
  const n = { id: t.id, name: t.name, members: t.members };
  KEYS.forEach(k => { n[k] = v[k] !== undefined ? v[k] : t[k]; });
  n.tags = t.tags; n.links = t.links;
  console.log('교체:', t.id, '잘한 점', n.strengths.length, '· 더 좋아지려면', n.improves.length);
  return n;
});
const t5 = path.join(dir, 'team5.fixed.json');
if (fs.existsSync(t5) && !out.some(t => t.id === 'team5')) {
  const v = JSON.parse(fs.readFileSync(t5, 'utf8'));
  const n = { id: 'team5', name: v.name || '팀 (이름 미정)', members: ['고준형', '김리온', '이호연'] };
  KEYS.forEach(k => { n[k] = v[k]; });
  n.one = ''; n.tags = []; n.links = [];
  out.push(n);
  console.log('추가: team5');
}
const header = '/* =====================================================================\n' +
  '   팀 데이터 — 2026-11-18 제출본, 원본 코드와 문장마다 대조해 확인한 내용만\n' +
  '   (확인하지 않은 추론은 넣지 않았습니다. 행 번호는 근거 위치)\n' +
  '   ===================================================================== */\n';
const start = s.lastIndexOf('/* ====', a);
const cut = start >= 0 && a - start < 400 ? start : a;
s = s.slice(0, cut) + header + 'var TEAMS = ' + JSON.stringify(out, null, 1) + ';\n' + s.slice(b);
fs.writeFileSync(html, s);
console.log('저장 · 팀', out.length);
