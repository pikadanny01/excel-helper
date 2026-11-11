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
  // 의도별 우선 규칙은 서로 다른 엑셀 작업(덧셈/문자 결합/셀 병합)을 혼동하지 않도록 분리한다.
  function classify(question) {
    const s=normalize(question), sq=squash(question);
    const has=(...parts)=>parts.some(p=>sq.includes(squash(p)));
    const refs=refContexts(question);
    const hints=[];
    const add=id=>{if(!hints.includes(id))hints.push(id);};
    const wantMerge=has('합치','합쳐','붙이','붙여','이어','연결','결합','하나로','한칸으로','합쳐서');
    const textWords=has('글자','문자','문자열','텍스트','이름','성명','성 과 이름','성하고 이름','단어','전화번호','단어들');
    const cellMerge=has('셀병합','셀 병합','칸 병합','병합하고','셀을 병합','병합해서','병합하고 가운데','셀두개합','셀 두 개 합','두셀합','셀 합치','칸 합치');
    const numbers=has('더하기','더하','더해','덧셈','플러스','합산','합계','총합');
    const range=/\$?[A-Z]{1,3}\$?\d+\s*(?:[:~]|부터|에서)\s*\$?[A-Z]{1,3}\$?\d+/i.test(question)||has('까지','범위','전체','모든','전부','여러','총액');
    const vague=(sq.length<17 && /^(이거|이걸|저거|이런|그런|이것)/.test(sq) && !has('더하','더해','합','찾','빼','곱','나누','평균','색','날짜','글자','숫자','병합','정렬','인쇄','원','중복'))
       || /^(이걸|이거|이런거|저걸)(이런식으로|어떻게)(나타내|나타낼|표시|해|할)/.test(sq);
    let ambiguous=false;
    // 지정된 작업이 아니라 '합쳐'만 있으면 엑셀에서는 3가지 뜻이 있어 추정 금지.
    if(wantMerge && !numbers && !cellMerge && !textWords && has('열','셀','칸','a1','b1','값')) {
      ambiguous=true;
    } else if(cellMerge) add('g096');
    else if(textWords && wantMerge) add('g051');
    else if(numbers){
      if(has('조건','특정','항목별','직원별','사람별','그룹별') && !has('무조건')) add('g009');
      else if(!range && (refs.length===2 || has('두칸','두셀','두개','둘','이거랑이거랑','각각','따로')))add('g001');
      else add('g002');
    }
    if(has('중복')) {
      if(has('지우','삭제','제거','없애','하나만 남'))add('g080');
      else if(has('찾아','찾기','찾는','표시','색','겹치','확인','구분'))add('g079');
    }
    if(has('년생','출생연도','출생년도','태어난년도') && has('참여','참석','참가','출석','미참여','불참','결석')) {
      if(has('몇명','인원수','사람수','사람 수','몇사람','몇건','몇개','세기','개수','몇 명','몇 명이','인원','명수'))add('g015');
      else add('g143');
    }
    if((/(?:^|[^가-힣])(남자|여자|남성|여성)(?=$|[가-힣\s])/u.test(s) || has('성별')) && has('참여','참석','미참여','불참','결석'))add('g144');
    if(has('나이','연령','살','세이상','세이하') && has('참여','참석','미참여','불참') && /\d+\s*(세|살|살이상|세이상)/.test(sq))add('g150');
    if(has('점수','성적','시험') && has('이상','초과','미만','이하') && has('사람','이름','명단','추출','찾아'))add('g145');
    if(has('참여횟수','참석횟수','이용횟수','출석횟수','이용 횟수','방문횟수') && has('이상','넘는','이하','초과','미만')){
      if(/(?:19|20)\d{2}년/.test(sq) && has('참여','참석','출석','이용'))add('g148');
      else add('g146');
    }
    if(has('a열','b열','두열','서로 다른 열') && has('둘다','모두','양쪽','공통','겹치','중복','양쪽에','두곳','두 군데') && has('이름','명단','사람'))add('g149');
    if(has('오타','맞춤법','비슷한단어','비슷한말'))add('g147');
    // 구체적인 오류 코드는 이름이 같은 함수보다 오류 진단을 우선한다.
    const errors=[['#div/0!','g116'],['#n/a','g117'],['#value!','g118'],['#ref!','g119'],['#name?','g120'],['#spill!','g121']];
    if(has('오류','에러','에러나','뜨는','떠','왜') || errors.some(([code])=>s.includes(code))) {
      for(const [code,id] of errors) if(s.includes(code))add(id);
    }
    if(has('조건부서식','조건부 서식') || (has('색','빨강','초록','하이라이트') && has('이상','이하','미만','초과','조건','참석률','출석률','참여율','퍼센트','%'))) {
      if(!has('중복'))add('g102');
    }
    if(has('몇퍼센트','백분율','비율','퍼센트','%') && !has('색','빨강','초록','조건부서식','조건부 서식')) add('g007');
    if(has('빼기','뺄셈','차액','마이너스','빼고','잔액') && !has('날짜')) {
      if(has('예산','집행','지출','사용한금액','사용금액','남은돈'))add('g129');
      else add('g003');
    }
    if(has('곱하기','곱셈','곱해'))add('g004');
    if(has('나누기','나눗셈','나눠','인당'))add('g005');
    if(has('평균'))add('g006');
    if(has('날짜','월','일자') && has('형식','표시','나타내','바꾸'))add('g077');
    if(has('형식','서식','나타내','표시') && has('원으로','원처럼','원표시','원붙','명으로','명처럼','명표시','개로','개처럼','개표시','단위'))add('g101');
    if(has('천단위','쉼표','콤마') && !has('분리'))add('g100');
    if(has('계산안','수식안','글자로','계산안되','수식그대로'))add('g122');
    if(has('출퇴근','근무시간','휴게시간'))add('g076');
    if(has('근무일수','평일수','공휴일제외'))add('g072');
    if(has('인쇄','출력','프린트') && has('한장','한페이지','잘림','잘려'))add('g109');
    if(has('시트합치','시트 병합','여러 시트의 데이터 합'))add('g087');
    if(has('빈 행','빈행','빈줄','빈 줄') && has('지우','삭제','제거','정리'))add('g086');
    if(has('고정','스크롤') && has('첫째행','첫행','첫 행','제목행','행고정','틀고정'))add('g090');
    if(has('가로로','세로로') && has('바꾸','뒤집','전환','변경'))add('g048');
    if(has('같은 사람','이름별','사람별') && has('지출','금액','매출') && has('합계','총합'))add('g130');
    return {hints,refs,vague,ambiguous};
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
    if (hints.includes(item.id)) score+=550;
    if (hints.length && !hints.includes(item.id)) score-=65;
    // 직접 오류 코드를 입력했으면 일반 함수명/오류가 아닌 다른 도움말을 앞세우지 않는다.
    if (/#[a-z0-9\/]+[!?]/i.test(query) && item.kind==='function') score-=200;
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
    if (intent.ambiguous) {
      const choices=['g051','g001','g096'].map(id=>filters.find(x=>x.id===id)).filter(Boolean);
      return {results:choices.map(item=>({item,score:0})),count:choices.length,intent,confidence:'ambiguous'};
    }
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
  function interpretConditions(question){
    const q=String(question||'').normalize('NFKC');
    const yearMatch=q.match(/(?:19|20)\d{2}(?=년생|\s*년\s*출생|\s*출생)/) || q.match(/(?:19|20)\d{2}(?=년생)/);
    const bornYear=yearMatch ? Number(yearMatch[0]) : null;
    // '미참여'에는 '참여'가 들어 있으므로 부정 상태를 우선 판별한다.
    const status= /미참여|불참|결석|미출석/.test(q) ? '미참여' : /참석|참여|출석|참가/.test(q) ? '참여' : null;
    const eventYearMatch=q.match(/((?:19|20)\d{2})\s*년에?(?:\s|$|[^가-힣])/);
    const eventYear=eventYearMatch?Number(eventYearMatch[1]):null;
    const gender=/(?:^|[^가-힣])(여자|여성)/u.test(q)?'여':/(?:^|[^가-힣])(남자|남성)/u.test(q)?'남':null;
    const threshold=q.match(/(\d+(?:\.\d+)?)\s*(?:점|회|번|세|살|%)?\s*(이상|초과|이하|미만)/);
    return {bornYear,eventYear,status,gender,threshold:threshold?Number(threshold[1]):null,operator:threshold?({'이상':'>=','초과':'>','이하':'<=','미만':'<'}[threshold[2]]):null};
  }
  function suggestedFormula(question,item){
    if(!item)return null;
    const refs=refContexts(question);
    if(item.id==='g001'&&refs.length>=2&&/더|합|플러스|덧셈/.test(question))return '='+refs[0]+'+'+refs[1];
    if(item.id==='g002'){
      const m=String(question).match(/(\$?[A-Z]{1,3}\$?\d+)\s*(?:[:~]|부터|에서)\s*(\$?[A-Z]{1,3}\$?\d+)/i);
      if(m)return '=SUM('+m[1].toUpperCase()+':'+m[2].toUpperCase()+')';
    }
    if(item.id==='g051'&&refs.length>=2) return '='+refs[0]+'&'+refs[1];
    const c=interpretConditions(question);
    if(item.id==='g149')return '=FILTER(A2:A100,COUNTIF(B2:B100,A2:A100)>0,"해당 없음")';
    if(item.id==='g150' && c.status && c.threshold!==null)
      return '=FILTER(A2:C100,(B2:B100'+c.operator+c.threshold+')*(C2:C100="'+c.status+'"),"해당 없음")';
    if(item.id==='g143' && c.bornYear && c.status)
      return '=FILTER(A2:C100,(B2:B100='+c.bornYear+')*(C2:C100="'+c.status+'"),"해당 없음")';
    if(item.id==='g015' && c.bornYear && c.status)
      return '=COUNTIFS(B2:B100,'+c.bornYear+',C2:C100,"'+c.status+'")';
    if(item.id==='g144' && c.gender && c.status)
      return '=FILTER(A2:C100,(B2:B100="'+c.gender+'")*(C2:C100="'+c.status+'"),"해당 없음")';
    if(item.id==='g145' && c.threshold!==null)
      return '=FILTER(A2:B100,B2:B100'+c.operator+c.threshold+',"해당 없음")';
    if(item.id==='g146' && c.threshold!==null)
      return '=FILTER(A2:B100,B2:B100'+c.operator+c.threshold+',"해당 없음")';
    if(item.id==='g148' && c.threshold!==null && c.eventYear && c.status)
      return '=FILTER(A2:D100,(B2:B100='+c.eventYear+')*(C2:C100'+c.operator+c.threshold+')*(D2:D100="'+c.status+'"),"해당 없음")';
    return null;
  }
  function dynamicExample(question,item){
    if(!item)return null;
    const c=interpretConditions(question);
    if(item.id==='g051') {
      const refs=refContexts(question), f=refs.length>=2?'='+refs[0]+'&'+refs[1]:'=A2&B2';
      return {headers:['성','이름','전체 이름'],rows:[['김','민수','김민수'],['박','지연','박지연']],formula:f,result:'김민수',explain:'문자열을 이어 붙일 때는 &를 사용합니다. 공백이 필요하면 =A2&" "&B2처럼 입력하세요. 셀 병합과는 다른 기능입니다.'};
    }
    if(item.id==='g149')return {headers:['A명단','B명단','검색 결과'],rows:[['김철수','이영희',''],['이영희','최지은','이영희'],['박민수','박민수','박민수']],formula:'=FILTER(A2:A4,COUNTIF(B2:B4,A2:A4)>0,"해당 없음")',result:'이영희, 박민수',explain:'A열 명단에서 B열에도 등장하는 이름을 찾습니다. 동일 이름이 서로 다른 사람일 수 있으니 실제 업무에서는 고유번호를 기준으로 대조하세요.'};
    if(item.id==='g150' && c.status && c.threshold!==null){
      const op=x=>({'=':x===c.threshold,'>=':x>=c.threshold,'>':x>c.threshold,'<=':x<=c.threshold,'<':x<c.threshold}[c.operator]);
      const rows=[['김철수',String(c.threshold),c.status],['이영희',String(c.threshold-1),c.status],['박민수',String(c.threshold+2),'미참여']];
      const names=rows.filter(r=>op(Number(r[1]))&&r[2]===c.status).map(r=>r[0]);
      return {headers:['이름','나이(숫자)','참여여부'],rows,formula:'=FILTER(A2:C4,(B2:B4'+c.operator+c.threshold+')*(C2:C4="'+c.status+'"),"해당 없음")',result:names.join(', ')||'해당 없음',explain:'나이가 이미 숫자로 입력된 B열을 가정합니다. 생년월일만 있는 경우에는 별도 나이 계산 과정이 필요합니다.'};
    }
    if(item.id==='g143' && c.bornYear && c.status){
      const other=c.bornYear===2005?2004:2005;
      const rows=[['김철수',String(c.bornYear),c.status],['이영희',String(other),'참여'],['박민수',String(c.bornYear),c.status==='참여'?'미참여':'참여'],['최지은',String(c.bornYear),c.status]];
      return {headers:['이름','출생연도','참여여부'],rows,formula:'=FILTER(A2:C5,(B2:B5='+c.bornYear+')*(C2:C5="'+c.status+'"),"해당 없음")',result:'김철수, 최지은',explain:'예시 데이터에서 '+c.bornYear+'년생이면서 '+c.status+' 상태인 사람만 찾습니다. B열이 숫자 출생연도이고 C열에 실제로 "'+c.status+'"라고 적혀 있다고 가정합니다.'};
    }
    if(item.id==='g015' && c.bornYear && c.status){
      const base=dynamicExample(question,{id:'g143'});
      return {...base,formula:'=COUNTIFS(B2:B5,'+c.bornYear+',C2:C5,"'+c.status+'")',result:'2명',explain:'조건에 맞는 사람의 수를 세는 예시입니다. 명단이 필요하면 FILTER를 사용하세요.'};
    }
    if(item.id==='g144' && c.gender && c.status){
      return {headers:['이름','성별','참여여부'],rows:[['김철수','남',c.status],['이영희','여','참여'],['박민수',c.gender,c.status],['최지은','여','미참여']],formula:'=FILTER(A2:C5,(B2:B5="'+c.gender+'")*(C2:C5="'+c.status+'"),"해당 없음")',result:c.gender==='여' && c.status==='참여'?'이영희, 박민수':c.gender==='여'?'박민수, 최지은':'김철수, 박민수',explain:'예시는 성별 열에 남/여, 참여여부 열에 참여/미참여라고 적힌 경우에 한합니다.'};
    }
    if(item.id==='g148' && c.threshold!==null && c.eventYear && c.status){
      const other=c.eventYear===2025?2024:2025;
      const rows=[['김철수',String(c.eventYear),String(c.threshold),c.status],['이영희',String(other),String(c.threshold+5),c.status],['박민수',String(c.eventYear),String(Math.max(0,c.threshold-1)),c.status],['최지은',String(c.eventYear),String(c.threshold+4),c.status]];
      const op=x=>({'=':x===c.threshold,'>=':x>=c.threshold,'>':x>c.threshold,'<=':x<=c.threshold,'<':x<c.threshold}[c.operator]);
      const names=rows.filter(r=>Number(r[1])===c.eventYear&&op(Number(r[2]))&&r[3]===c.status).map(r=>r[0]);
      return {headers:['이름','참여연도','이용횟수','참여여부'],rows,formula:'=FILTER(A2:D5,(B2:B5='+c.eventYear+')*(C2:C5'+c.operator+c.threshold+')*(D2:D5="'+c.status+'"),"해당 없음")',result:names.join(', ')||'해당 없음',explain:'B열은 참여연도(숫자), C열은 이용횟수(숫자), D열은 참여여부(참여/미참여)인 예시입니다. 실제 파일의 열 배치가 다르면 주소를 바꿔야 합니다.'};
    }
    if(['g145','g146'].includes(item.id) && c.threshold!==null){
      const freq=item.id==='g146',unit=freq?'회':'점';
      let a=c.threshold + (c.operator==='>'?1:0),b=c.threshold-1,cc=c.threshold+10;
      if(c.operator==='<='||c.operator==='<') {a=c.threshold-(c.operator==='<'?1:0);b=c.threshold+1;cc=c.threshold-3;}
      const match=x=>({'=':x===c.threshold,'>=':x>=c.threshold,'>':x>c.threshold,'<=':x<=c.threshold,'<':x<c.threshold}[c.operator]);
      const rows=[['김철수',String(a)],['이영희',String(b)],['최지은',String(cc)]];
      const out=rows.filter(r=>match(Number(r[1]))).map(r=>r[0]).join(', ')||'해당 없음';
      return {headers:['이름',freq?'참여횟수':'점수'],rows,formula:'=FILTER(A2:B4,B2:B4'+c.operator+c.threshold+',"해당 없음")',result:out,explain:'수치가 숫자로 입력되어 있다고 가정합니다. '+unit+' 기준 '+c.operator+c.threshold+' 조건을 적용한 가상 예시입니다.'};
    }
    return null;
  }
  return {search,normalize,tokens,classify,scoreEntry,refContexts,suggestedFormula,interpretConditions,dynamicExample,isIncompatible};
});
