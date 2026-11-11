/* Excel Helper — 오프라인 규칙 기반 검색엔진. 사용자 입력을 외부로 전송하지 않습니다. */
(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.ExcelSearch = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';

  const STOP = new Set(['이거','이걸','그거','이것','저것','이런','그런','어떻게','어케','어떤','무엇','뭐','좀','그냥','방법','알려줘','알려주세요','알려','할까','하나요','하려면','하면','할수','하는','하기','해요','해주세요','싶어','되나요','되게','엑셀','excel','셀에서','제가','시트에서','데이터에서','것을','있는','이랑','저거','거를','거랑','이거랑','그걸','할때','하는법','하는방법','거','좀요','무슨','싶은데','가능','화면','표를','값을','거야','해','법']);
  const GROUPS = [
    ['합계','합산','총합','총계','덧셈','더하기','더하','더해','플러스','합치기','sum'],
    ['빼기','빼','빼면','차액','차감','뺄셈','마이너스'],
    ['곱하기','곱셈','곱해','곱하','곱','乘','multiply'],
    ['나누기','나눗셈','나누','인당','일인당'],
    ['평균','평균값','산술평균','average'],
    ['중복','중복값','겹치는','같은값','겹친'],
    ['제거','삭제','지우','없애','없애기'],
    ['찾기','찾아','조회','검색','찾아오기','가져오기'],
    ['표시','나타내','보이','보여','보이게','표현'],
    ['조건','경우','맞으면','해당하는'],
    ['서식','꾸미','형식','모양'],
    ['정렬','순서','오름차순','내림차순','가나다순'],
    ['인쇄','프린트','출력','뽑기'],
    ['합치기','붙이기','이어붙이기','연결','병합'],
    ['오류','에러','에러가','문제','잘못'],
    ['숫자','수치','금액'],
    ['명단','이름','사람'],
    ['날짜','일자','일시'],
    ['색','색깔','컬러','빨간색','강조'],
    ['만료','기한','마감'],
    ['건수','횟수','개수','몇개','몇명','세기','갯수'],
    ['줄바꿈','줄 바꿈','개행','두줄','줄넘김'],
    ['드롭다운','선택목록','유효성검사','목록선택'],
    ['필터','거르기','조건만보기'],
    ['백분율','퍼센트','%','비율','달성률'],
    ['차트','그래프','막대그래프','꺾은선'],
    ['피벗','피벗테이블','집계표'],
    ['열너비','칸크기','칸 너비','열 넓이'],
    ['시트','탭','워크시트']
  ];
  const synonyms = new Map();
  for (const grp of GROUPS) for (const x of grp) synonyms.set(x,grp[0]);

  function normalize(s) {
    return String(s || '').normalize('NFKC').toLowerCase()
      .replace(/[\u200b-\u200d\ufeff]/g, '').replace(/[^가-힣a-z0-9_.$%#?!:+*\/-]+/g, ' ').trim().replace(/\s+/g, ' ');
  }
  function squash(s) { return normalize(s).replace(/\s/g, ''); }
  function tokens(s) {
    const n = normalize(s);
    const words = n.split(/\s+/).filter(Boolean);
    const out = [];
    for (let w of words) {
      if (STOP.has(w) || /^([a-z]{1,3}\$?\d+)$/.test(w)) continue;
      let stem=w;
      if (w.length >= 4 && /[가-힣]/.test(w)) {
        stem = w.replace(/(으로는|으로|에서|에게|까지|처럼|마다|하고|해서|하면|하려면|하는|보다|으로|하면|하는거|했어|이에요|해줘|주세요|있어|있음|하는법|하는데|하려고|일때|을|를|의|에|로|은|는|이|가|과|와|만)$/u, '');
      }
      const canonical = synonyms.get(w) || synonyms.get(stem) || stem;
      if (canonical.length > 1 && !STOP.has(canonical)) out.push(canonical);
      if (stem !== canonical && stem.length > 1) out.push(stem);
    }
    return [...new Set(out)];
  }

  const REF_PATTERN = /\$?[a-z]{1,3}\$?\d{1,7}/gi;
  function refContexts(question) {
    const q=String(question||'');
    const extracted=q.match(REF_PATTERN)||[];
    return [...new Set(extracted.map(x=>x.toUpperCase()))].slice(0,6);
  }
  function classify(question) {
    const s = normalize(question), sq=squash(question);
    const has = (...parts) => parts.some(p=>sq.includes(squash(p)));
    const refs=refContexts(question);
    const prompts=[];
    const explicit=(id)=>prompts.push(id);
    // 구분해야 하는 반대 의도 우선: 문자 연결, 셀 병합, 숫자 계산
    if (has('글자 합치','글자두개합','문자 합치','문자열 합치','이름 합치','텍스트 합치','글자 붙이','문자 붙이','글자를붙','문자를붙')) explicit('g051');
    else if (has('셀 병합','셀 합치','셀두개합','칸 합치','하나의 큰 셀','병합하고')) explicit('g096');
    else if (has('더하','더해','더하기','합계','총합','덧셈','플러스','합산')) {
      if (has('조건','특정','별','경우','항목별') && !has('무조건')) explicit('g009');
      else if (has('두칸','두셀','두개','둘','이거랑이거랑','와','하고','a1','b2','각각','따로') && !has('전부','전체','모두','여러','범위','1월부터')) explicit('g001');
      else explicit('g002');
    }
    if (has('중복')) {
      if (has('지우','삭제','제거','없애')) explicit('g080');
      else if (has('찾아','찾기','찾는','표시','색','겹치','확인','구분')) explicit('g079');
    }
    if (has('몇퍼센트','백분율','비율','퍼센트','%')) explicit('g007');
    if (has('빼기','뺄셈','차액','마이너스') && !has('날짜')) explicit('g003');
    if (has('곱하기','곱셈','곱해')) explicit('g004');
    if (has('나누기','나눗셈','나눠','인당')) explicit('g005');
    if (has('평균')) explicit('g006');
    if (has('날짜','월','일자') && has('형식','표시','나타내','바꾸')) explicit('g077');
    if (has('형식','서식','나타내','표시') && has('원으로','원처럼','원표시','원붙','명으로','명처럼','명표시','개로','개처럼','개표시','단위')) explicit('g101');
    if (has('천단위','쉼표','콤마') && !has('분리')) explicit('g100');
    if (has('계산안','수식안','글자로','계산안되','수식그대로')) explicit('g122');
    if (has('출퇴근','근무시간','휴게시간')) explicit('g076');
    if (has('근무일수','평일수','공휴일제외')) explicit('g072');
    if (has('인쇄','출력','프린트') && has('한장','한페이지','잘림','잘려')) explicit('g109');
    // 사용자의 지시어만 있는 질문에서는 근거 없는 답변을 제시하지 않음
    const vague = (sq.length<17 && /^(이거|이걸|저거|이런|그런|이것)/.test(sq) && !has('더하','더해','합','찾','빼','곱','나누','평균','색','날짜','글자','숫자','병합','정렬','인쇄','원','중복'))
      || /^(이걸|이거|이런거|저걸)(이런식으로|어떻게)(나타내|나타낼|표시|해|할)/.test(sq);
    return {hints:[...new Set(prompts)],refs,vague};
  }

  function scoreEntry(item, query, qtokens, hints, version) {
    const nq=squash(query); if (!nq) return 0;
    const title=squash(item.title);
    const name=item.name&&squash(item.name);
    const keyList=(item.keys||[]).map(squash).filter(Boolean);
    let score=0;
    if (name && name === nq) score+=260;
    if (name && new RegExp('(^|[^a-z])'+name.replace(/\./g,'\\.')+'($|[^a-z])','i').test(normalize(query))) score+=205;
    if (name && nq===name+'함수') score+=260;
    if (title===nq) score+=190;
    if (nq.length>=2 && title.includes(nq)) score+=94;
    if (nq.length>=3 && nq.includes(title)) score+=105;
    for (const k of keyList) {
      if (k === nq) score+=110;
      else if (k.length>=2 && nq.includes(k)) score+=Math.min(70,35+k.length*3);
      else if (nq.length>=3 && k.includes(nq)) score+=Math.min(60,20+nq.length*3);
    }
    const titleTokens=tokens(item.title);
    const keyTokens=tokens((item.keys||[]).join(' '));
    const descTokens=tokens(item.desc||'');
    let matched=0;
    for (const t of qtokens) {
      let gain=0;
      if (t.length<2)continue;
      if (titleTokens.includes(t)) gain=24;
      else if (keyTokens.includes(t))gain=19;
      else if (titleTokens.some(w=>w.includes(t)&&t.length>=2))gain=14;
      else if (keyTokens.some(w=>w.includes(t)&&t.length>=2))gain=12;
      else if (descTokens.some(w=>w.includes(t)&&t.length>=3))gain=5;
      if (gain) matched++;
      score+=gain;
    }
    if (qtokens.length>=2 && matched>=2) score+= 10*matched/qtokens.length;
    if (hints.includes(item.id)) score+=330;
    score+= (item.priority||0)*0.8;
    if (item.kind==='function' && !item.verified) score-=14;
    if (item.kind==='function' && !name?.includes(nq)) score-=10;
    // 선택된 구버전에서 최신 버전 함수를 최상위로 권장하지 않음 (단 직접 검색은 남김)
    if (version && version!=='365' && isIncompatible(item.version,version)) score-=30;
    return Math.max(0,Math.round(score));
  }

  function isIncompatible(version, selected) {
    if (!version || version==='all' || version==='unknown'||selected==='365') return false;
    const year=Number(selected);
    if (version==='365/2024')return year<2024;
    const match=String(version).match(/^(\d{4})\+/);
    return !!match && year<Number(match[1]);
  }

  function search(data, query, options={}) {
    const category=options.category||'전체';
    const version=options.version||'365';
    const limit=options.limit||12;
    const filters = category==='전체' ? [...data.guides,...data.functions] :
      category==='함수 사전'?data.functions:data.guides.filter(x=>x.category===category);
    const q=normalize(query);
    const intent=classify(q);
    if (!q) {
      // 최초 진입시 인기 순/기능 소개. 함수 사전은 사전순.
      const items=category==='함수 사전'?filters:filters.filter(x=>x.kind==='guide').sort((a,b)=>(b.priority||0)-(a.priority||0));
      return {results:items.slice(0,limit).map(x=>({item:x,score:0})),count:items.length,intent,confidence:'none'};
    }
    if (intent.vague) return {results:[],count:0,intent,confidence:'vague'};
    const qtokens=tokens(q);
    const ranked=filters.map(item=>({item,score:scoreEntry(item,q,qtokens,intent.hints,version)}))
      .filter(x=>x.score>=18).sort((a,b)=>b.score-a.score||String(a.item.title).localeCompare(String(b.item.title),'ko'));
    // 동일 목적 도움말은 중복이 있을 수 있으므로 맨 앞 정확한 응답을 우선
    const max=ranked[0]?.score||0;
    let confidence=max>=95?'high':max>=45?'medium':max>0?'low':'none';
    // 검색어가 두 단어 이상의 낯선 문장이면, 단어 하나만 매칭한 낮은 검색 점수를 낮은 확신도로 제한
    if (qtokens.length>=3 && max<110 && !intent.hints.length) confidence='low';
    return {results:ranked.slice(0,limit),count:ranked.length,intent,confidence};
  }
  function suggestedFormula(question,item) {
    if (!item || item.id!=='g001')return null;
    const refs=refContexts(question);
    if (refs.length>=2 && /더|합|플러스|덧셈/.test(question)) return '='+refs[0]+'+'+refs[1];
    return null;
  }
  return {search,normalize,tokens,classify,scoreEntry,refContexts,suggestedFormula,isIncompatible};
});
