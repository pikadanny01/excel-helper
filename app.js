(function() {
  'use strict';
  const DB=window.EXCEL_HELPER_DATA;
  if(!DB || !Array.isArray(DB.guides) || !Array.isArray(DB.functions)) { document.body.textContent='도움말 파일을 불러올 수 없습니다. 파일 구성을 확인해 주세요.'; return; }
  const customStore='exh-custom-v12', feedbackStore='exh-feedback-v12';
  // Added entries live in this browser only until exported and deployed by an administrator.
  try {
    const extras=JSON.parse(localStorage.getItem(customStore)||'[]');
    if(Array.isArray(extras))for(const x of extras.slice(0,100))if(x&&/^local-[a-z0-9-]+$/.test(x.id)&&typeof x.title==='string'&&typeof x.desc==='string'&&x.title.length<=120&&x.desc.length<=2000){DB.guides.push(x);}
  }catch{}
  const engine=window.ExcelSearch;
  if (!DB||!engine) {document.body.textContent='도움말 파일을 불러올 수 없습니다. 모든 파일을 같은 폴더에 두세요.';return;}
  const $=id=>document.getElementById(id);
  const el={query:$('query'),form:$('searchForm'),list:$('resultList'),detail:$('detail'),related:$('related'),title:$('sectionTitle'),subtitle:$('resultSubtitle'),cat:$('categories'),notice:$('notice'),more:$('moreBtn'),version:$('excelVersion'),toast:$('toast'),tool:$('toolPanel')};
  const CATS={전체:'▦','기본 계산':'∑','조건·조회':'⌕','문자·텍스트':'Aa','날짜·시간':'◷','데이터 정리':'≡','서식·표시':'◈','인쇄·차트':'▥','문제 해결':'!','업무 예시':'▤','함수 사전':'ƒx'};
  const EXAMPLE=['두 셀의 숫자 더하기','중복값 색깔로 찾기','한 장에 맞춰 인쇄하기','표에서 값 찾기 (VLOOKUP)','조건에 맞는 합계 (SUMIF)','날짜 입력과 표시 형식'];
  function localRead(key,fallback){try{return JSON.parse(localStorage.getItem(key))??fallback;}catch{return fallback;}}
  function localWrite(key,v){try{localStorage.setItem(key,JSON.stringify(v));}catch{}}
  const state={category:'전체',query:'',version:localRead('exh-version','365'),saved:new Set(localRead('exh-saved',[])),savedOnly:false,selected:null,limit:12,last:null,theme:localRead('exh-theme','light')};
  const ALL=[...DB.guides,...DB.functions];
  const byId=new Map(ALL.map(x=>[x.id,x]));
  const escapeHTML=(s)=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const safeAttr=escapeHTML;
  const sourceURL=(url)=>/^https:\/\/support\.microsoft\.com\//.test(url||'')?url:DB.sources.help;
  const versionLabel=(v)=>({'all':'대부분 버전','2013+':'Excel 2013 이상','2016+':'Excel 2016 이상','2019+':'Excel 2019 이상','2021+':'Excel 2021 이상','365/2024':'Microsoft 365 / Excel 2024','unknown':'버전 확인 필요'})[v]||'버전 확인 필요';
  function toast(s){el.toast.textContent=s;el.toast.classList.add('visible');clearTimeout(toast.timer);toast.timer=setTimeout(()=>el.toast.classList.remove('visible'),2200);}
  function theme(){document.documentElement.dataset.theme=state.theme;$('themeBtn').textContent=state.theme==='dark'?'☀':'☾';$('themeBtn').setAttribute('aria-label',state.theme==='dark'?'밝은 테마로 전환':'어두운 테마로 전환');}
  function renderCats(){
    el.cat.innerHTML=DB.categories.map(c=>{
      const count=c==='전체'?DB.guides.length+DB.functions.length:c==='함수 사전'?DB.functions.length:DB.guides.filter(x=>x.category===c).length;
      return `<button class="cat-btn ${state.category===c&&!state.savedOnly?'active':''}" type="button" data-cat="${safeAttr(c)}"><span class="cat-symbol">${CATS[c]}</span><span>${escapeHTML(c)}</span><span class="cat-count">${count}</span></button>`;
    }).join('');
    $('showSaved').classList.toggle('active',state.savedOnly);
    $('savedCount').textContent=state.saved.size;
  }
  function resultCard(x,rank){
    const a=x.item;
    return `<button class="result-row" data-open="${safeAttr(a.id)}" type="button"><span class="item-emblem">${escapeHTML(a.kind==='function'?'ƒx':(CATS[a.category]||'✓'))}</span><span class="result-body"><span class="result-name">${escapeHTML(a.title)}</span><span class="result-desc">${escapeHTML(a.desc)}</span><span class="result-tags">${escapeHTML(a.kind==='function'?'함수 사전':a.category)}${a.kind==='function'&&!a.verified?' · 공식 문서 참조':''}</span></span><span class="result-arrow">↗</span></button>`;
  }
  function colName(i){let out='';for(let n=i+1;n>0;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;}
  function questionContext(a){
    const q=state.query.trim();
    if(!q)return '';
    const c=engine.interpretConditions(q);
    const labels=[];
    if(['g143','g015','g154'].includes(a.id) && c.bornYear && c.status){
      labels.push('출생연도 '+c.bornYear+'년','참여 여부 '+c.status);if(a.id==='g154'&&c.threshold!==null)labels.push('횟수 '+c.operator+' '+c.threshold);
    } else if(['g144','g151'].includes(a.id) && c.gender && c.status){
      labels.push('성별 '+(c.gender==='남'?'남성':'여성'),'참여 여부 '+c.status);if(a.id==='g151'&&c.threshold!==null)labels.push('횟수 '+c.operator+' '+c.threshold);
    } else if(a.id==='g152' && c.status){const m=q.match(/((?:19|20)\d{2})\s*년\s*(0?[1-9]|1[0-2])\s*월/);if(m)labels.push(m[1]+'년 '+Number(m[2])+'월',c.status);
    } else if(['g145','g146'].includes(a.id) && c.threshold!==null){
      labels.push((a.id==='g145'?'점수 ':'이용횟수 ')+c.operator+' '+c.threshold);
    } else if(a.id==='g148' && c.eventYear && c.status && c.threshold!==null){
      labels.push('참여연도 '+c.eventYear+'년','이용횟수 '+c.operator+' '+c.threshold,'참여 여부 '+c.status);
    }
    return labels.join(' · ');
  }
  function visualExampleHTML(a){
    const custom=engine.dynamicExample(state.query,a);
    const e=custom||window.EXCEL_HELPER_VISUALS?.[a.id];
    if(!e)return '';
    const headers=e.headers||[],rows=e.rows||[];
    const count=Math.max(headers.length,...rows.map(r=>r.length),1);
    const letterRow=Array.from({length:count},(_,i)=>`<th scope="col" class="excel-column">${colName(i)}</th>`).join('');
    const headRow=Array.from({length:count},(_,i)=>`<th scope="col" class="excel-header">${escapeHTML(String(headers[i]||'').replace(/^[A-Z]{1,2}\s*\((.*?)\)$/, '$1')||'')}</th>`).join('');
    const body=rows.map((r,i)=>`<tr><th scope="row" class="excel-row-index">${i+2}</th>${Array.from({length:count},(_,c)=>`<td>${escapeHTML(r[c]??'')}</td>`).join('')}</tr>`).join('');
    const formula=e.formula?`<div class="visual-formula"><span>예시 수식</span><code>${escapeHTML(e.formula)}</code><button type="button" class="copy-btn" data-copy="${safeAttr(e.formula)}">복사 ⧉</button></div>`:'';
    const result=e.result?`<div class="visual-outcome"><strong>예상 결과</strong><span>${escapeHTML(e.result)}</span></div>`:'';
    const context=custom&&questionContext(a)?`<div class="question-context"><strong>✓ 현재 질문 조건 적용</strong><span>${escapeHTML(questionContext(a))}</span></div>`:'';
    const label=custom?'질문에 적은 조건을 반영한 가상 예시입니다.':'미리 등록된 일반 예시입니다. 질문의 숫자·조건이 자동 반영된 예시는 아닙니다.';
    return `<section class="answer-section visual-section"><h4>엑셀 예시: 열 문자와 행 번호를 확인하세요</h4>${context}<p class="visual-lead">${label} 1행은 제목이고 2행부터 데이터입니다. 실제 표의 열 배치가 다르면 수식 주소도 바꿔야 합니다.</p><div class="visual-scroll"><table class="example-table excel-grid"><thead><tr><th class="excel-corner" aria-label="셀 위치"></th>${letterRow}</tr><tr><th scope="row" class="excel-row-index">1</th>${headRow}</tr></thead><tbody>${body}</tbody></table></div>${formula}${result}${e.explain?`<p class="visual-explain">${escapeHTML(e.explain)}</p>`:''}</section>`;
  }
  function detailHTML(a){
    let formula=engine.suggestedFormula(state.query,a);
    const bad=engine.isIncompatible(a.version,state.version);
    const source=sourceURL(a.source);
    const examples=formula?[formula]:[...new Set(a.examples||[])];
    const verified=a.kind==='function'?a.verified!==false:true;
    const c=engine.interpretConditions(state.query);
    const steps=[...(a.steps||[])];
    // 고정 예시의 과거 숫자·조건이 맞춤형 질문 설명에 섞이지 않도록 분리
    const generated=engine.dynamicExample(state.query,a);
    if (generated && ['g143','g015','g144','g145','g146','g148','g150','g151','g152','g154'].includes(a.id)) {
      steps.length=0;steps.push('아래 가상 예시 표의 열 위치(A·B·C·D)와 실제 파일의 열을 비교하세요.','질문에 입력된 조건을 적용한 예시 수식을 빈 셀에 입력합니다.','결과가 다르면 수식의 열 범위, 숫자/텍스트 구분, Excel 버전을 확인하세요.');
    }
    if(a.id==='g143' && c.bornYear && c.status){
      steps.length=0;
      steps.push('A열: 이름, B열: 출생연도(숫자), C열: 참여 여부로 표를 구성합니다.',
        'B열은 '+c.bornYear+'년, C열은 '+c.status+'인 행을 찾습니다.',
        'Microsoft 365 또는 Excel 2021 이상이면 아래 FILTER 수식을 사용하고, 구버전에서는 데이터 → 필터에서 두 열의 조건을 각각 선택하세요.');
    }
    const stepHTML=steps.length?`<section class="answer-section"><h4>따라 하는 방법</h4><ol class="steps">${steps.map(s=>`<li>${escapeHTML(s)}</li>`).join('')}</ol></section>`:'';
    const examplesHTML=examples.length?`<section class="answer-section"><h4>${a.kind==='function'?'사용 예시':'예시 수식·입력값'}</h4>${examples.map(x=>`<div class="formula-row"><code>${escapeHTML(x)}</code><button type="button" class="copy-btn" data-copy="${safeAttr(x)}" aria-label="예시 복사">복사 ⧉</button></div>`).join('')}<p class="micro-note">예시의 셀 주소는 가상의 주소입니다. 실제 파일에 맞게 변경하세요.</p></section>`:'';
    return `<div class="answer-top"><div class="answer-icon">${escapeHTML(a.kind==='function'?'ƒx':CATS[a.category]||'✓')}</div><div class="answer-meta"><div class="answer-kicker">${escapeHTML(a.kind==='function'?'함수 사전':a.category)} <span>· ${escapeHTML(versionLabel(a.version))}</span></div><h3>${escapeHTML(a.title)}</h3></div><button class="bookmark ${state.saved.has(a.id)?'marked':''}" type="button" data-save="${safeAttr(a.id)}" title="즐겨찾기 저장" aria-label="${state.saved.has(a.id)?'저장 해제':'저장하기'}">${state.saved.has(a.id)?'★':'☆'}</button></div>
      ${bad?'<div class="warning">선택한 Excel 버전에서는 이 기능이 지원되지 않을 수 있습니다. 버전을 확인하세요.</div>':''}
      ${verified?'':'<div class="warning subdued">이 함수는 이름만 색인에 등록돼 있습니다. 구문과 설명은 Microsoft 공식 문서에서 확인해 주세요.</div>'}
      <p class="answer-intro">${escapeHTML(a.desc)}</p>${stepHTML}${visualExampleHTML(a)}${examplesHTML}
      ${(a.tips||[]).length?`<section class="answer-section tip-section"><h4>알아두세요</h4>${a.tips.map(t=>`<p>• ${escapeHTML(t)}</p>`).join('')}</section>`:''}
      <div class="answer-actions"><a href="${safeAttr(source)}" target="_blank" rel="noopener noreferrer" class="source-link">Microsoft 공식 도움말 ↗</a><button type="button" class="plain-action" data-copy-answer="${safeAttr(a.id)}">설명 복사</button><button type="button" class="plain-action" data-print="1">인쇄</button><button type="button" class="plain-action" data-feedback="${safeAttr(a.id)}" data-val="yes">도움됐어요</button><button type="button" class="plain-action" data-feedback="${safeAttr(a.id)}" data-val="no">부족해요</button></div>`;
  }
  function notice(text,type='info') {el.notice.textContent=text;el.notice.hidden=!text;el.notice.className='notice '+type;}
  function render(){
    el.tool.hidden=true;
    renderCats();
    el.more.hidden=true;
    const isQuery=!!state.query.trim();
    const results=engine.search(DB,state.query,{category:state.category,version:state.version,limit:Math.max(100,ALL.length)});
    if (state.savedOnly) {
      results.results=ALL.filter(a=>state.saved.has(a.id)).map(item=>({item,score:0}));
      results.count=results.results.length;
    }
    state.last=results;
    if (state.savedOnly) {el.title.textContent='저장한 도움말';el.subtitle.textContent='이 컴퓨터에 저장한 즐겨찾기입니다.';}
    else if(state.category==='함수 사전'&&!isQuery){el.title.textContent='엑셀 함수 사전';el.subtitle.textContent=`${DB.functions.length}개 함수 이름을 둘러보세요. 설명이 없는 항목은 공식 문서를 확인하세요.`;}
    else if(isQuery){el.title.textContent='검색 결과';el.subtitle.textContent=`“${state.query.slice(0,50)}${state.query.length>50?'…':''}”에 대한 관련 항목 ${results.count}개` ;}
    else if(state.category!=='전체'){el.title.textContent=state.category;el.subtitle.textContent='실무에 자주 사용하는 해결 방법을 모았습니다.';}
    else {el.title.textContent='자주 찾는 해결 방법';el.subtitle.textContent='원하는 항목을 누르거나 위 검색창에 질문해 보세요.';}

    if (isQuery && results.confidence==='unsupported'){
      notice('질문에서 여러 조건이 발견됐지만, 현재 자동 수식 생성기가 모든 조건을 함께 처리하지 못합니다. 일부 조건을 누락한 수식을 제시하지 않도록 결과를 제한했습니다. 예: 조건을 나누어 필터로 적용하거나 원하는 열 배치를 구체적으로 질문해 주세요.','warning');
      el.list.innerHTML=exampleLinks();el.detail.hidden=true;el.related.innerHTML='';return;
    }
    if (isQuery && results.confidence==='vague'){
      notice('원하는 결과 모양이나 셀 주소가 빠져 있어 정확한 방법을 고르기 어렵습니다. 예: “A1과 B1을 더하고 싶어”, “날짜를 2026-11-11 형태로 보이고 싶어”','warning');
      el.list.innerHTML=exampleLinks();el.detail.hidden=true;el.related.innerHTML='';return;
    }
    if (isQuery && results.confidence==='ambiguous') notice('「합치기」는 숫자 더하기, 글자 이어 붙이기, 셀 병합 중 어떤 작업인지에 따라 방법이 달라집니다. 아래에서 원하는 작업을 선택해 주세요.','warning');
    else if (isQuery && results.confidence==='low') notice('질문에 정확히 맞는 답을 확신하기 어려워 연관된 항목을 보여드립니다. 제목과 예시를 확인해 주세요.','warning');
    else if (isQuery&&!results.count)notice('아직 등록된 해결법이 없습니다. 다른 표현으로 다시 검색하거나 Microsoft 공식 도움말을 확인해 주세요.','warning');
    else notice('');

    if (!results.count){ el.list.innerHTML=exampleLinks();el.detail.hidden=true;el.related.innerHTML='';return;}
    if (!isQuery && state.category==='전체'&&!state.savedOnly&&!state.selected){
      const items=EXAMPLE.map(title=>DB.guides.find(x=>x.title===title)).filter(Boolean);
      el.list.innerHTML='<div class="popular-grid">'+items.map(a=>`<button type="button" class="popular-card" data-open="${a.id}"><span class="popular-glyph">${CATS[a.category]}</span><span class="popular-name">${escapeHTML(a.title)}</span><span class="popular-desc">${escapeHTML(a.desc)}</span><span class="popular-go">자세히 보기 ↗</span></button>`).join('')+'</div>';
      el.detail.hidden=true;el.related.innerHTML='';return;
    }
    const limit=state.limit;
    const items=results.results.slice(0,limit);
    let selected=state.selected&&byId.has(state.selected)?byId.get(state.selected):null;
    if (!selected && isQuery && !['low','ambiguous'].includes(results.confidence)) selected=items[0]?.item;
    if (!selected && !isQuery) selected=null;
    if (selected){
      state.selected=selected.id;
      el.detail.hidden=false; el.detail.innerHTML=detailHTML(selected);
      const others=items.filter(x=>x.item.id!==selected.id).slice(0,7);
      el.list.innerHTML='';
      el.related.innerHTML=others.length?`<h3 class="related-title">함께 보면 좋은 답변</h3><div class="result-list">${others.map(resultCard).join('')}</div>`:'';
    } else {
      el.detail.hidden=true; el.detail.innerHTML=''; el.related.innerHTML='';
      el.list.innerHTML='<div class="result-list">'+items.map(resultCard).join('')+'</div>';
    }
    el.more.hidden=results.count<=limit;
    el.more.textContent=`더 많은 결과 보기 (${Math.min(results.count-limit,24)}개)`;
  }
  function exampleLinks(){return `<div class="empty"><div class="empty-symbol">⌕</div><h3>조금 더 구체적으로 물어보세요</h3><p>어떤 데이터를 어떤 모습으로 만들고 싶은지 알려주면 더 잘 찾습니다.</p><div class="empty-examples">${['A1이랑 B1을 더하려면?','글자 두 개 붙이는 방법은?','날짜를 연-월-일로 표시하는 법','직원별 지출 금액 합치기'].map(q=>`<button type="button" data-query="${safeAttr(q)}">${escapeHTML(q)}</button>`).join('')}</div></div>`;}
  function showTool(title,body){
    el.tool.hidden=false;el.tool.innerHTML='<h3>'+escapeHTML(title)+'</h3>'+body;
    el.detail.hidden=true;el.list.innerHTML='';el.related.innerHTML='';el.more.hidden=true;
    el.title.textContent=title;el.subtitle.textContent='이 도구는 AI 없이 현재 브라우저에서 작동합니다.';
    el.tool.scrollIntoView({behavior:'smooth',block:'start'});
  }
  function openBuilder(){
    const cols=Array.from({length:12},(_,i)=>colName(i));
    const opts=cols.map(c=>`<option value="${c}">${c}열</option>`).join('');
    showTool('조건 수식 자동 만들기',`<p class="tool-intro">실제 엑셀 열 위치를 선택하면 두 조건을 만족하는 명단 조회 수식을 만들어 드립니다. 데이터를 서버로 보내지 않습니다.</p>
      <div class="tool-fields"><label>데이터 마지막 행 <input id="fbEnd" type="number" min="2" max="1000000" value="100"></label>
      <label>첫 번째 조건 열 <select id="fbCol1">${opts}</select></label><label>첫 번째 조건 값 <input id="fbVal1" maxlength="60" value="2005"></label>
      <label>두 번째 조건 열 <select id="fbCol2">${opts}</select></label><label>두 번째 조건 값 <input id="fbVal2" maxlength="60" value="참여"></label>
      <label>찾을 결과 <select id="fbType"><option value="list">조건에 맞는 명단</option><option value="count">조건에 맞는 사람 수</option></select></label></div>
      <button type="button" class="action-primary" id="fbCreate">수식 만들기</button><div id="fbResult" class="tool-result"></div>
      <p class="micro-note">가정: 1행은 제목, 데이터는 2행부터, 전체 표는 A열부터 D열까지입니다. 숫자는 숫자 조건으로 처리합니다. Microsoft 365 또는 Excel 2021 이상에서 FILTER를 사용할 수 있습니다.</p>`);
    $('fbCol1').value='B';$('fbCol2').value='C';
    $('fbCreate').addEventListener('click',()=>{
      const end=Math.max(2,Math.min(1000000,Number($('fbEnd').value)||100));
      const c1=$('fbCol1').value,c2=$('fbCol2').value;
      const v1=$('fbVal1').value.trim(),v2=$('fbVal2').value.trim();
      if(!v1||!v2){$('fbResult').textContent='두 조건 값을 입력하세요.';return;}
      const cond=v=>/^[-+]?\d+(?:\.\d+)?$/.test(v)?v:'"'+v.replace(/"/g,'""')+'"';
      const kind=$('fbType').value;
      const formula=kind==='count'?`=COUNTIFS(${c1}2:${c1}${end},${cond(v1)},${c2}2:${c2}${end},${cond(v2)})`:`=FILTER(A2:D${end},(${c1}2:${c1}${end}=${cond(v1)})*(${c2}2:${c2}${end}=${cond(v2)}),"해당 없음")`;
      $('fbResult').innerHTML=`<div class="formula-row"><code>${escapeHTML(formula)}</code><button type="button" class="copy-btn" data-copy="${safeAttr(formula)}">복사 ⧉</button></div><p class="micro-note">A:D 표 범위와 셀 주소가 실제 파일과 맞는지 확인하세요.</p>`;
    });
    $('fbCreate').click();
  }
  function parseDelimited(text){
    const rows=[];let row=[],cell='',quote=false;const s=String(text).replace(/^\uFEFF/,'');
    const delimiter=(s.split(/\r?\n/)[0]||'').includes('\t')?'\t':',';
    for(let i=0;i<s.length&&rows.length<120;i++){
      const ch=s[i];
      if(ch==='"'){if(quote&&s[i+1]==='"'){cell+='"';i++;}else quote=!quote;}
      else if(ch===delimiter&&!quote){row.push(cell);cell='';}
      else if((ch==='\n'||ch==='\r')&&!quote){if(ch==='\r'&&s[i+1]==='\n')i++;row.push(cell);rows.push(row);row=[];cell='';}
      else cell+=ch;
    }
    if(cell||row.length){row.push(cell);rows.push(row);}
    return rows;
  }
  function openCSV(){
    showTool('CSV·TSV 데이터 미리보기',`<p class="tool-intro">CSV 또는 탭으로 구분한 텍스트 파일을 이 브라우저 안에서만 읽어 미리 봅니다. Excel .xlsx 원본 파일은 지원하지 않습니다.</p><label class="file-input">CSV·TSV 파일 선택 <input id="localCSV" type="file" accept=".csv,.tsv,.txt,text/csv,text/plain"></label><div id="csvPreview"></div><p class="micro-note">민감한 개인정보가 포함된 업무 파일은 이 사이트에서 취급하지 않는 것을 권장합니다. 파일 내용은 전송되지 않지만 PC의 다른 프로그램·확장 프로그램 보안은 별도로 관리해야 합니다.</p>`);
    $('localCSV').addEventListener('change',async()=>{
      const file=$('localCSV').files?.[0];if(!file)return;
      if(file.size>2*1024*1024){$('csvPreview').textContent='2MB 이하의 CSV 파일만 미리 볼 수 있습니다.';return;}
      let text;try{text=await file.text();}catch{$('csvPreview').textContent='파일을 읽을 수 없습니다.';return;}
      let rows=parseDelimited(text).slice(0,25);if(!rows.length){$('csvPreview').textContent='표시할 데이터가 없습니다.';return;}
      const columns=Math.min(15,Math.max(...rows.map(x=>x.length)));
      const headers=Array.from({length:columns},(_,i)=>`<th scope="col" class="excel-column">${colName(i)}</th>`).join('');
      const lines=rows.map((r,i)=>`<tr><th scope="row" class="excel-row-index">${i+1}</th>${Array.from({length:columns},(_,j)=>`<td>${escapeHTML(String(r[j]??'').slice(0,180))}</td>`).join('')}</tr>`).join('');
      $('csvPreview').innerHTML=`<p class="tool-intro">미리보기: 최대 25행 · 15열 / 파일 크기 ${Math.round(file.size/1024)}KB</p><div class="visual-scroll"><table class="example-table excel-grid"><thead><tr><th class="excel-corner"></th>${headers}</tr></thead><tbody>${lines}</tbody></table></div><p class="micro-note">이 표에서 열 위치를 확인한 다음 '조건 수식 만들기'를 이용하세요. CSV 데이터는 실제 Excel에 적용·수정되지 않습니다.</p>`;
    });
  }
  function downloadJSON(name,object){const blob=new Blob([JSON.stringify(object,null,2)],{type:'application/json;charset=utf-8'});const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),2000);}
  function openCustom(){
    showTool('나만의 도움말 관리',`<p class="tool-intro">등록한 도움말은 <strong>현재 브라우저에만 저장</strong>됩니다. 다른 직원에게 공유하려면 내보낸 파일을 검토 후 GitHub에 별도로 반영해야 합니다.</p>
      <div class="tool-fields"><label>도움말 제목 <input id="customTitle" maxlength="120" placeholder="예: 프로그램 참석자 명단 만들기"></label>
      <label>관련 검색어(쉼표 구분) <input id="customKeys" maxlength="400" placeholder="참석자, 참여자, 명단"></label>
      <label class="wide">답변 내용 <textarea id="customDesc" maxlength="2000" rows="4" placeholder="구체적인 해결 방법을 작성하세요."></textarea></label>
      <label class="wide">예시 수식 (선택) <input id="customExample" maxlength="400" placeholder="=COUNTIF(B2:B100,\"참여\")"></label></div>
      <div class="tool-buttons"><button id="customSave" type="button" class="action-primary">내 브라우저에 추가</button><button id="customExport" type="button" class="action-secondary">도움말 데이터 내보내기</button><button id="feedbackExport" type="button" class="action-secondary">평가 내역 내보내기</button></div><div id="customMsg" class="micro-note"></div>`);
    $('customSave').addEventListener('click',()=>{
      const title=$('customTitle').value.trim(),desc=$('customDesc').value.trim(),example=$('customExample').value.trim();
      if(title.length<2||desc.length<5){$('customMsg').textContent='제목 2글자 이상, 답변 5글자 이상 입력하세요.';return;}
      const id='local-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,6);
      const guide={id,kind:'guide',category:'업무 예시',title,keys:$('customKeys').value.split(',').map(x=>x.trim()).filter(Boolean).slice(0,15),desc,steps:[],examples:example?[example]:[],tips:['이 브라우저에서 직접 작성한 도움말입니다. 내용을 확인하세요.'],source:'https://support.microsoft.com/ko-kr/excel/',version:'unknown',priority:1};
      let existing=localRead(customStore,[]);existing=Array.isArray(existing)?existing.slice(-99):[];existing.push(guide);localWrite(customStore,existing);DB.guides.push(guide);ALL.push(guide);byId.set(id,guide);
      $('customMsg').textContent='현재 브라우저에 등록했습니다. 검색할 수 있습니다. 다른 직원에게는 자동으로 공유되지 않습니다.';
    });
    $('customExport').addEventListener('click',()=>downloadJSON('excel-helper-custom-guides.json',localRead(customStore,[])));
    $('feedbackExport').addEventListener('click',()=>downloadJSON('excel-helper-feedback.json',localRead(feedbackStore,[])));
  }
  function applyQuery(value){state.category='전체';state.query=String(value||'').slice(0,220);el.query.value=state.query;state.limit=12;state.selected=null;state.savedOnly=false;render();}
  async function copyText(str){
    try{if(navigator.clipboard&&window.isSecureContext){await navigator.clipboard.writeText(str);}else{
      const x=document.createElement('textarea');x.value=str;x.style.cssText='position:fixed;left:-10000px;top:0';document.body.appendChild(x);x.select();if(!document.execCommand('copy'))throw new Error('copy');x.remove();
    }toast('클립보드에 복사했습니다.');} catch{toast('복사할 수 없습니다. 텍스트를 직접 선택해 주세요.');}
  }
  function openItem(id){if(!byId.has(id))return;state.selected=id; const a=byId.get(id);if(state.category!=='전체'&&state.category!==a.category&&!state.savedOnly)state.category='전체';render(); $('results-area').scrollIntoView({behavior:'smooth',block:'start'});}
  document.addEventListener('click',function(e){
    const btn=e.target.closest('button');if(!btn)return;
    if(btn.dataset.query!==undefined){applyQuery(btn.dataset.query);el.query.focus();return;}
    if(btn.dataset.cat!==undefined){state.category=btn.dataset.cat;state.savedOnly=false;state.selected=null;state.limit=12;render();return;}
    if(btn.dataset.open!==undefined){openItem(btn.dataset.open);return;}
    if(btn.dataset.copy!==undefined){copyText(btn.dataset.copy);return;}
    if(btn.dataset.copyAnswer!==undefined){const a=byId.get(btn.dataset.copyAnswer);if(a)copyText([a.title,a.desc,...(a.steps||[]),...(a.examples||[]),...(a.tips||[])].join('\n'));return;}
    if(btn.dataset.save!==undefined){const id=btn.dataset.save;state.saved.has(id)?state.saved.delete(id):state.saved.add(id);localWrite('exh-saved',[...state.saved]);render();toast(state.saved.has(id)?'저장한 도움말에 추가했습니다.':'저장한 도움말에서 제거했습니다.');return;}
    if(btn.dataset.print!==undefined){window.print();return;}
    if(btn.dataset.feedback!==undefined){
      const records=localRead(feedbackStore,[]);
      const list=Array.isArray(records)?records.slice(-199):[];
      list.push({helpId:btn.dataset.feedback,useful:btn.dataset.val==='yes',date:new Date().toISOString().slice(0,10)});
      localWrite(feedbackStore,list);toast('평가가 이 브라우저에 저장되었습니다. 관리자에게 자동 전송되지는 않습니다.');return;
    }
  });
  el.form.addEventListener('submit',function(e){e.preventDefault();applyQuery(el.query.value);$('results-area').scrollIntoView({behavior:'smooth',block:'start'});});
  let timer;
  el.query.addEventListener('input',function(){clearTimeout(timer);timer=setTimeout(()=>applyQuery(el.query.value),200);});
  el.version.addEventListener('change',function(){state.version=el.version.value;localWrite('exh-version',state.version);render();});
  $('themeBtn').addEventListener('click',function(){state.theme=state.theme==='dark'?'light':'dark';localWrite('exh-theme',state.theme);theme();});
  $('showSaved').addEventListener('click',function(){state.savedOnly=true;state.category='전체';state.selected=null;state.query='';el.query.value='';state.limit=12;render();});
  $('openFormulaBuilder').addEventListener('click',openBuilder);
  $('openCSV').addEventListener('click',openCSV);
  $('openCustom').addEventListener('click',openCustom);
  $('openDictionary').addEventListener('click',function(){state.category='함수 사전';state.savedOnly=false;state.selected=null;state.query='';el.query.value='';state.limit=12;render();$('results-area').scrollIntoView({behavior:'smooth',block:'start'});});
  el.more.addEventListener('click',function(){state.limit+=24;render();});
  document.addEventListener('keydown',function(e){if(e.key==='/'&&!['INPUT','TEXTAREA'].includes(document.activeElement.tagName)&&!e.metaKey&&!e.ctrlKey){e.preventDefault();el.query.focus();}});
  // query & category are intentionally not encoded in URL to avoid sharing private question text.
  $('guideCount').textContent=DB.guides.length;
  $('fnCount').textContent=DB.functions.length;
  $('allCount').textContent=DB.functions.length;
  $('verifiedCount').textContent=DB.functions.filter(f=>f.verified).length;
  $('appVersion').textContent=DB.version;
  el.version.value=state.version;
  theme();render();
})();
