(function() {
  'use strict';
  const DB=window.EXCEL_HELPER_DATA;
  const engine=window.ExcelSearch;
  if (!DB||!engine) {document.body.textContent='도움말 파일을 불러올 수 없습니다. 모든 파일을 같은 폴더에 두세요.';return;}
  const $=id=>document.getElementById(id);
  const el={query:$('query'),form:$('searchForm'),list:$('resultList'),detail:$('detail'),related:$('related'),title:$('sectionTitle'),subtitle:$('resultSubtitle'),cat:$('categories'),notice:$('notice'),more:$('moreBtn'),version:$('excelVersion'),toast:$('toast')};
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
  function visualExampleHTML(a){
    const e=window.EXCEL_HELPER_VISUALS?.[a.id];
    if(!e)return '';
    const header=(e.headers||[]).map(x=>`<th scope="col">${escapeHTML(x)}</th>`).join('');
    const rows=(e.rows||[]).map(r=>`<tr>${r.map(v=>`<td>${escapeHTML(v)}</td>`).join('')}</tr>`).join('');
    const formula=e.formula?`<div class="visual-formula"><span>적용 수식</span><code>${escapeHTML(e.formula)}</code><button type="button" class="copy-btn" data-copy="${safeAttr(e.formula)}">복사 ⧉</button></div>`:'';
    const output=e.result?`<div class="visual-outcome"><strong>예상 결과</strong><span>${escapeHTML(e.result)}</span></div>`:'';
    return `<section class="answer-section visual-section"><h4>실제로 적용하면 이렇게 됩니다</h4><p class="visual-lead">이해를 돕기 위한 가상 데이터 예시입니다.</p><div class="visual-scroll"><table class="example-table"><thead><tr>${header}</tr></thead><tbody>${rows}</tbody></table></div>${formula}${output}<p class="visual-explain">${escapeHTML(e.explain)}</p></section>`;
  }
  function detailHTML(a){
    let formula=engine.suggestedFormula(state.query,a);
    const bad=engine.isIncompatible(a.version,state.version);
    const source=sourceURL(a.source);
    const examples=[...new Set([...(formula?[formula]:[]),...(a.examples||[])])];
    const verified=a.kind==='function'?a.verified!==false:true;
    const stepHTML=(a.steps||[]).length?`<section class="answer-section"><h4>따라 하는 방법</h4><ol class="steps">${a.steps.map(s=>`<li>${escapeHTML(s)}</li>`).join('')}</ol></section>`:'';
    const examplesHTML=examples.length?`<section class="answer-section"><h4>${a.kind==='function'?'사용 예시':'예시 수식·입력값'}</h4>${examples.map(x=>`<div class="formula-row"><code>${escapeHTML(x)}</code><button type="button" class="copy-btn" data-copy="${safeAttr(x)}" aria-label="예시 복사">복사 ⧉</button></div>`).join('')}<p class="micro-note">예시의 셀 주소는 가상의 주소입니다. 실제 파일에 맞게 변경하세요.</p></section>`:'';
    return `<div class="answer-top"><div class="answer-icon">${escapeHTML(a.kind==='function'?'ƒx':CATS[a.category]||'✓')}</div><div class="answer-meta"><div class="answer-kicker">${escapeHTML(a.kind==='function'?'함수 사전':a.category)} <span>· ${escapeHTML(versionLabel(a.version))}</span></div><h3>${escapeHTML(a.title)}</h3></div><button class="bookmark ${state.saved.has(a.id)?'marked':''}" type="button" data-save="${safeAttr(a.id)}" title="즐겨찾기 저장" aria-label="${state.saved.has(a.id)?'저장 해제':'저장하기'}">${state.saved.has(a.id)?'★':'☆'}</button></div>
      ${bad?'<div class="warning">선택한 Excel 버전에서는 이 기능이 지원되지 않을 수 있습니다. 버전을 확인하세요.</div>':''}
      ${verified?'':'<div class="warning subdued">이 함수는 이름만 색인에 등록돼 있습니다. 구문과 설명은 Microsoft 공식 문서에서 확인해 주세요.</div>'}
      <p class="answer-intro">${escapeHTML(a.desc)}</p>${stepHTML}${visualExampleHTML(a)}${examplesHTML}
      ${(a.tips||[]).length?`<section class="answer-section tip-section"><h4>알아두세요</h4>${a.tips.map(t=>`<p>• ${escapeHTML(t)}</p>`).join('')}</section>`:''}
      <div class="answer-actions"><a href="${safeAttr(source)}" target="_blank" rel="noopener noreferrer" class="source-link">Microsoft 공식 도움말 ↗</a><button type="button" class="plain-action" data-copy-answer="${safeAttr(a.id)}">설명 복사</button><button type="button" class="plain-action" data-print="1">인쇄</button></div>`;
  }
  function notice(text,type='info') {el.notice.textContent=text;el.notice.hidden=!text;el.notice.className='notice '+type;}
  function render(){
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

    if (isQuery && results.confidence==='vague'){
      notice('원하는 결과 모양이나 셀 주소가 빠져 있어 정확한 방법을 고르기 어렵습니다. 예: “A1과 B1을 더하고 싶어”, “날짜를 2026-11-11 형태로 보이고 싶어”','warning');
      el.list.innerHTML=exampleLinks();el.detail.hidden=true;el.related.innerHTML='';return;
    }
    if (isQuery && results.confidence==='low') notice('질문에 정확히 맞는 답을 확신하기 어려워 연관된 항목을 보여드립니다. 제목과 예시를 확인해 주세요.','warning');
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
    if (!selected && isQuery && results.confidence!=='low') selected=items[0]?.item;
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
  function applyQuery(value){state.query=String(value||'').slice(0,220);el.query.value=state.query;state.limit=12;state.selected=null;state.savedOnly=false;render();}
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
  });
  el.form.addEventListener('submit',function(e){e.preventDefault();applyQuery(el.query.value);$('results-area').scrollIntoView({behavior:'smooth',block:'start'});});
  let timer;
  el.query.addEventListener('input',function(){clearTimeout(timer);timer=setTimeout(()=>applyQuery(el.query.value),200);});
  el.version.addEventListener('change',function(){state.version=el.version.value;localWrite('exh-version',state.version);render();});
  $('themeBtn').addEventListener('click',function(){state.theme=state.theme==='dark'?'light':'dark';localWrite('exh-theme',state.theme);theme();});
  $('showSaved').addEventListener('click',function(){state.savedOnly=true;state.category='전체';state.selected=null;state.query='';el.query.value='';state.limit=12;render();});
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
