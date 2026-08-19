/**
 * SIMS Keyword Explorer v0.1.4
 * P1 prototype: Internal Discovery from SIMS Site Collector Evidence.
 *
 * Scope:
 * - Select/import one site's Collector Evidence ZIP
 * - Analyze page/query evidence
 * - Optional Article Master matching
 * - Cluster query intent before Candidate Registry
 * - Search Persona Profile from aggregated GSC query behavior
 * - OWNED_QUERY exclusion + novelty gate
 * - Candidate Gate: max 10 internal Blue Ocean candidates
 * - Build Candidate Registry
 * - Generate Doctor referral ZIPs with user-recognizable names
 *
 * Not included in P1:
 * - External Web Discovery
 * - SERP Gap automation
 * - Automatic Creator execution
 */

const SKE_VERSION = '0.1.4';
const SKE_PRODUCT_NAME = 'SIMS Keyword Explorer';
const SKE_CONFIG = {
  sheets: {
    home: 'Home',
    candidates: 'キーワード候補',
    settings: '_SKE_SETTINGS',
    pageSummary: '_SKE_EVIDENCE_PAGE_SUMMARY',
    pageWeekly: '_SKE_EVIDENCE_PAGE_WEEKLY',
    pageQuery: '_SKE_EVIDENCE_PAGE_QUERY',
    querySummary: '_SKE_EVIDENCE_QUERY_SUMMARY',
    articleMaster: '_SKE_ARTICLE_MASTER',
    persona: '検索ペルソナ'
  },
  candidateHeaders: [
    '選択','Candidate ID','SiteID','ブログ','Primary Query','Discovery Type',
    'P1 Score','需要成熟度','記事寿命','既存記事判定','関連ArticleID','関連URL',
    'Engine判定','状態','表示回数','クリック','平均順位','URL数','発見理由',
    'Doctor判定','Doctor確信度','次回確認日','公開ArticleID','公開URL','更新日時'
  ]
};

function onOpen() {
  // Menus must be added before any setup work. If setup fails on a fresh sheet,
  // the user must still have a visible recovery path.
  skeBuildMenu_();
}

function onInstall(e) {
  onOpen(e);
}

function skeBuildMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SIMS Keyword Explorer')
    .addItem('1. 初期設定 / 対象ブログを準備', 'skeInitialSetup')
    .addItem('2. Evidenceを読み込む', 'skeImportEvidencePrompt')
    .addItem('3. 新しいキーワード候補を探す', 'skeRunInternalDiscovery')
    .addItem('4. 検索ペルソナを確認する', 'skeOpenPersonaProfile')
    .addItem('5. 候補を確認する', 'skeOpenCandidates')
    .addItem('6. 処置を進める', 'skeContinueWorkflow')
    .addSeparator()
    .addSubMenu(ui.createMenu('追加の操作')
      .addItem('Article Masterの使い方', 'skeArticleMasterHelp')
      .addItem('選択候補のDoctor用ZIPを作る', 'skeGenerateDoctorPackageForSelected')
      .addItem('Homeを更新', 'skeRenderHome'))
    .addToUi();
}

function skeInitialSetup() {
  skeSetup_();
  SpreadsheetApp.getUi().alert(
    'SIMS Keyword Explorerの初期設定が完了しました。\n\n' +
    '次の操作：「2. Evidenceを読み込む」を実行してください。'
  );
}

function skeSetup_() {
  const ss = SpreadsheetApp.getActive();
  const ensure = (name, headers, hidden) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (headers && headers.length) {
      if (sh.getLastRow() === 0 || String(sh.getRange(1,1).getValue()||'') !== headers[0]) {
        sh.clear();
        sh.getRange(1,1,1,headers.length).setValues([headers]);
      }
      sh.setFrozenRows(1);
    }
    if (hidden && !sh.isSheetHidden()) sh.hideSheet();
    return sh;
  };
  ensure(SKE_CONFIG.sheets.home, null, false);
  ensure(SKE_CONFIG.sheets.candidates, SKE_CONFIG.candidateHeaders, false);
  ensure(SKE_CONFIG.sheets.settings, ['Key','Value'], true);
  ensure(SKE_CONFIG.sheets.pageSummary, null, true);
  ensure(SKE_CONFIG.sheets.pageWeekly, null, true);
  ensure(SKE_CONFIG.sheets.pageQuery, null, true);
  ensure(SKE_CONFIG.sheets.querySummary, null, true);
  ensure(SKE_CONFIG.sheets.articleMaster, ['ArticleID','記事タイトル','記事URL','メインクエリ','SearchIntent','状態'], true);
  ensure(SKE_CONFIG.sheets.persona, ['順位','検索ペルソナ','主な対象','主な困りごと','代表クエリ','表示回数','クリック','構成比','外部探索テーマ'], false);
  skeSetSetting_('version', SKE_VERSION);
  skeRenderHome();
}

function skeRenderHome() {
  skeSetupLight_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SKE_CONFIG.sheets.home);
  sh.clear();
  const siteName = skeGetSetting_('siteName') || '未選択';
  const siteUrl = skeGetSetting_('siteUrl') || '';
  const fileName = skeGetSetting_('evidenceFileName') || '未読込';
  const cand = skeReadObjects_(SKE_CONFIG.sheets.candidates);
  const count = k => cand.filter(r => String(r['状態']||'') === k).length;
  const rows = [
    ['SIMS Keyword Explorer', `v${SKE_VERSION}`],
    ['現在の対象ブログ', siteName],
    ['サイトURL', siteUrl],
    ['Evidence', fileName],
    ['', ''],
    ['新規候補', count('DISCOVERED')],
    ['Doctor診断候補', cand.filter(r=>String(r['Engine判定']||'')==='DOCTOR_REVIEW').length],
    ['再確認待ち', count('EARLY_OPPORTUNITY')],
    ['既存記事改善候補', count('WRITER_REDIRECT')],
    ['公開済み', count('PUBLISHED')],
    ['', ''],
    ['次の操作', siteName==='未選択' ? '2. Evidenceを読み込む' : '3. 新しいキーワード候補を探す']
  ];
  sh.getRange(1,1,rows.length,2).setValues(rows);
  sh.getRange('A1:B1').setFontWeight('bold').setFontSize(16);
  sh.getRange('A2:A12').setFontWeight('bold');
  sh.setColumnWidth(1,180); sh.setColumnWidth(2,520);
}

function skeSetupLight_() {
  const ss=SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SKE_CONFIG.sheets.settings)) {
    let sh=ss.insertSheet(SKE_CONFIG.sheets.settings); sh.getRange(1,1,1,2).setValues([['Key','Value']]); sh.hideSheet();
  }
}

function skeImportEvidencePrompt() {
  skeSetup_();
  const root=DriveApp.getRootFolder();
  const html=HtmlService.createHtmlOutput(skeEvidencePickerHtml_({folderId:root.getId(),folderName:'マイドライブ'}))
    .setWidth(700).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html,'Evidence Packageを選ぶ');
}

function skeListEvidencePickerFolder(folderId){
  let folder;
  try{folder=folderId?DriveApp.getFolderById(folderId):DriveApp.getRootFolder();}
  catch(e){folder=DriveApp.getRootFolder();}
  const folders=[],files=[];
  let it=folder.getFolders(),n=0;
  while(it.hasNext()&&n<150){const f=it.next();folders.push({id:f.getId(),name:f.getName()});n++;}
  let fit=folder.getFiles(),m=0;
  while(fit.hasNext()&&m<300){
    const f=fit.next(),name=f.getName();
    if(/\.zip$/i.test(name)&&(/SIMS/i.test(name)||/Evidence/i.test(name)))files.push({id:f.getId(),name:name,updated:f.getLastUpdated().toISOString()});
    m++;
  }
  folders.sort((a,b)=>a.name.localeCompare(b.name,'ja'));
  files.sort((a,b)=>String(b.updated).localeCompare(String(a.updated)));
  let parent=null;
  try{const ps=folder.getParents();if(ps.hasNext()){const p=ps.next();parent={id:p.getId(),name:p.getName()||'マイドライブ'};}}catch(e){}
  return {id:folder.getId(),name:folder.getName()||'マイドライブ',parent:parent,folders:folders,files:files};
}

function skeInspectEvidenceFile(fileId){
  const file=DriveApp.getFileById(fileId);
  if(!/\.zip$/i.test(file.getName()))throw new Error('ZIPファイルではありません。');
  const blobs=Utilities.unzip(file.getBlob());
  let manifest=null;
  blobs.forEach(b=>{if(String(b.getName()||'').split('/').pop()==='manifest.json'){try{manifest=JSON.parse(b.getDataAsString('UTF-8'));}catch(e){}}});
  const site=manifest&&manifest.site?manifest.site:{};
  const period=manifest&&manifest.period?manifest.period:{};
  return {
    fileId:file.getId(),fileName:file.getName(),
    siteName:String(site.siteName||site.site_name||''),
    siteUrl:String(site.siteUrl||site.site_url||site.searchConsoleProperty||''),
    generatedAt:String((manifest&&((manifest.generatedAt)||(manifest.generated_at)))||''),
    periodLabel:period.days?String(period.days)+'日':(period.start&&period.end?period.start+' ～ '+period.end:''),
    format:String((manifest&&manifest.format)||'')
  };
}

function skeImportSelectedEvidence(payload){
  const fileId=String(payload&&payload.fileId||'');
  if(!fileId)throw new Error('Evidence Packageが選択されていません。');
  const r=skeImportEvidenceById_(fileId);
  return {ok:true,siteName:r.siteName,fileName:r.fileName,queryRows:r.queryRows,next:'次は「3. 新しいキーワード候補を探す」を実行してください。'};
}

function skeEvidencePickerHtml_(o){
  const data=JSON.stringify(o||{}).replace(/</g,'\\u003c');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:20px}
  .hero{background:#185abc;color:#fff;padding:16px 18px;border-radius:10px}.hero h2{margin:0 0 5px;font-size:20px}.hero p{margin:0;font-size:13px}
  .card{background:#fff;border:1px solid #dadce0;border-radius:10px;margin-top:14px;padding:14px}.bar{display:flex;gap:8px;align-items:center}.where{flex:1;font-weight:bold;color:#174ea6}
  button{border:1px solid #dadce0;background:#fff;border-radius:6px;padding:8px 12px;cursor:pointer}button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}
  .list{height:235px;overflow:auto;border:1px solid #e0e0e0;border-radius:7px;margin-top:10px}.row{padding:9px 11px;border-bottom:1px solid #f1f3f4;cursor:pointer}.row:hover{background:#f8f9fa}.selected{background:#e8f0fe!important}
  .meta{margin-top:12px;background:#f8fafd;border-radius:7px;padding:10px;line-height:1.7;font-size:13px}.hint{color:#5f6368;font-size:12px}.err{color:#b3261e;margin-top:8px}.actions{text-align:right;margin-top:12px}
  </style></head><body><div class="wrap"><div class="hero"><h2>Evidence Packageを読み込む</h2><p>Collectorで作成したZIPを、Google Drive内のフォルダーを移動して選択します。</p></div>
  <div class="card"><div class="bar"><button id="up">↑ 上へ</button><div id="where" class="where"></div></div><div id="list" class="list"></div><div class="hint">📁 フォルダーをクリックして移動し、📦 Evidence ZIPを選択してください。</div></div>
  <div id="meta" class="meta">Evidence Packageを選択すると、サイト名・URL・作成日時・収集期間を確認できます。</div><div id="err" class="err"></div>
  <div class="actions"><button onclick="google.script.host.close()">キャンセル</button> <button id="import" class="primary" disabled>このEvidenceを読み込む</button></div>
  </div><script>
  const init=${data};let current=null,selected=null;
  const esc=s=>String(s||'').replace(/[&<>\"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]));
  function fail(e){document.getElementById('err').textContent=e.message||e;}
  function load(id){document.getElementById('list').innerHTML='<div class="row">読み込み中...</div>';google.script.run.withSuccessHandler(render).withFailureHandler(fail).skeListEvidencePickerFolder(id);}
  function render(d){current=d;document.getElementById('where').textContent=d.name;document.getElementById('up').disabled=!d.parent;const box=document.getElementById('list');box.innerHTML='';
    d.folders.forEach(f=>{const x=document.createElement('div');x.className='row';x.textContent='📁 '+f.name;x.onclick=()=>load(f.id);box.appendChild(x);});
    d.files.forEach(f=>{const x=document.createElement('div');x.className='row';x.textContent='📦 '+f.name;x.onclick=()=>choose(f,x);box.appendChild(x);});
    if(!d.folders.length&&!d.files.length)box.innerHTML='<div class="row">このフォルダーにEvidence ZIPはありません。</div>';
  }
  function choose(f,el){selected=f;document.querySelectorAll('.selected').forEach(x=>x.classList.remove('selected'));el.classList.add('selected');document.getElementById('import').disabled=true;
    document.getElementById('meta').textContent='内容を確認しています...';google.script.run.withSuccessHandler(m=>{document.getElementById('meta').innerHTML='<b>'+esc(m.fileName)+'</b><br>サイト名：'+esc(m.siteName||'不明')+'<br>サイトURL：'+esc(m.siteUrl||'不明')+'<br>作成日時：'+esc(m.generatedAt||'不明')+'<br>収集期間：'+esc(m.periodLabel||'不明');document.getElementById('import').disabled=false;}).withFailureHandler(fail).skeInspectEvidenceFile(f.id);}
  document.getElementById('up').onclick=()=>{if(current&&current.parent)load(current.parent.id)};
  document.getElementById('import').onclick=()=>{if(!selected)return;const b=document.getElementById('import');b.disabled=true;b.textContent='読み込み中...';google.script.run.withSuccessHandler(r=>{document.getElementById('meta').innerHTML='<b style="color:#137333">読み込み完了</b><br>サイト：'+esc(r.siteName)+'<br>Query行：'+esc(r.queryRows)+'<br>'+esc(r.next);b.textContent='閉じる';b.disabled=false;b.onclick=()=>google.script.host.close();}).withFailureHandler(e=>{b.disabled=false;b.textContent='このEvidenceを読み込む';fail(e)}).skeImportSelectedEvidence({fileId:selected.id});};
  load(init.folderId);
  </script></body></html>`;
}

function skeImportEvidenceById_(fileId) {
  const file=DriveApp.getFileById(fileId);
  const name=file.getName();
  if(!/\.zip$/i.test(name)) throw new Error(`ZIPファイルではありません: ${name}`);
  const blobs=Utilities.unzip(file.getBlob()), map={};
  blobs.forEach(b=>{const n=String(b.getName()||'').split('/').pop(); if(n) map[n]=b;});
  const required=[
    ['page_summary.csv',SKE_CONFIG.sheets.pageSummary],
    ['page_weekly.csv',SKE_CONFIG.sheets.pageWeekly],
    ['page_query_top.csv',SKE_CONFIG.sheets.pageQuery],
    ['query_summary.csv',SKE_CONFIG.sheets.querySummary]
  ];
  const missing=required.filter(x=>!map[x[0]]).map(x=>x[0]);
  if(missing.length) throw new Error(`Evidence ZIPに必要ファイルがありません: ${missing.join(', ')}`);
  let manifest={};
  if(map['manifest.json']) try{manifest=JSON.parse(map['manifest.json'].getDataAsString('UTF-8'));}catch(e){}
  const site=manifest.site||{};
  const siteName=String(site.siteName||site.site_name||skeSiteNameFromUrl_(site.siteUrl||site.site_url||'')||'Unknown Site');
  const siteUrl=String(site.siteUrl||site.site_url||site.searchConsoleProperty||'');
  const siteId=skeSiteIdFromUrl_(siteUrl||siteName);
  required.forEach(x=>skeImportCsvBlobToSheet_(map[x[0]],x[1]));
  skeSetSetting_('siteName',siteName); skeSetSetting_('siteUrl',siteUrl); skeSetSetting_('siteId',siteId);
  skeSetSetting_('evidenceFileId',fileId); skeSetSetting_('evidenceFileName',name);
  skeSetSetting_('evidenceGeneratedAt',String(manifest.generatedAt||manifest.generated_at||''));
  skeRenderHome();
  return {siteName,siteUrl,siteId,fileName:name,queryRows:Math.max(0,skeSheet_(SKE_CONFIG.sheets.pageQuery).getLastRow()-1)};
}


function skeOpenPersonaProfile(){
  const sh=skeSheet_(SKE_CONFIG.sheets.persona);
  SpreadsheetApp.getActive().setActiveSheet(sh);
}

function skeBuildPersonaProfile_(qRows){
  const defs=[
    {name:'Windows困りごと層', target:'Windows / PC', problem:'設定・エラー・できない・戻らない', re:/\bwindows\b|windows ?1[01]|win ?1[01]|pc\b|パソコン|defender|office|edge/i,
     intent:/できない|エラー|error|消え|戻ら|反応しない|開かない|解除|設定|不具合|トラブル/i,
     explore:'Windows Update、新機能、廃止機能、仕様変更、不具合、エラーメッセージ'},
    {name:'Apple設定・トラブル層', target:'iPhone / Apple Watch / Mac', problem:'同期・設定・使い方・不具合', re:/iphone|ipad|apple ?watch|mac(book|os)?|ios|watchos|icloud|airpods/i,
     intent:/できない|同期|設定|使い方|方法|消え|反応|エラー|解除|変更/i,
     explore:'iOS/macOS/watchOS新機能、設定変更、同期問題、仕様変更'},
    {name:'Androidトラブル層', target:'Android / Pixel / Galaxy', problem:'アップデート後・表示・設定・不具合', re:/android|pixel|galaxy|quick ?share|google ?play/i,
     intent:/できない|消え|表示されない|追加できない|反応しない|設定|アップデート|不具合|エラー/i,
     explore:'Android/Pixel新機能、OS更新後の不具合、Google公式修正情報'},
    {name:'生成AIトラブル層', target:'ChatGPT / Gemini / Claude等', problem:'生成失敗・エラー・使い方', re:/chat ?gpt|chatgpt|gemini|claude|copilot|ai\b|openai/i,
     intent:/error|エラー|something went wrong|できない|生成|使い方|意味|直し方|反応/i,
     explore:'生成AIの新エラー、仕様変更、モデル変更、新機能、障害'},
    {name:'動画・SNS設定層', target:'YouTube / LINE / Instagram等', problem:'再生・通知・表示・設定', re:/youtube|line|instagram|インスタ|tiktok|x twitter|twitter|facebook/i,
     intent:/できない|設定|解除|オフ|消え|通知|再生|開かない|反応/i,
     explore:'SNS/動画サービスのUI変更、新設定、新不具合、機能廃止'},
    {name:'PC・IT用語調査層', target:'PC / IT一般', problem:'意味・違い・仕組みを知りたい', re:/とは|意味|違い|仕組み|種類|役割|読み方/i,
     intent:/とは|意味|違い|仕組み|種類|役割|読み方/i,
     explore:'新しいPC/IT用語、新規格、新サービス、新機能の初心者向け解説'}
  ];

  const stats={};
  defs.forEach(d=>stats[d.name]={def:d,impressions:0,clicks:0,queries:{}});
  let totalImp=0;

  qRows.forEach(r=>{
    const q=String(skeObj_(r,['query','クエリ'])||'').trim();
    if(!q)return;
    const imp=Number(skeObj_(r,['impressions','表示回数'])||0);
    const clk=Number(skeObj_(r,['clicks','クリック数'])||0);
    totalImp+=imp;

    let best=null,score=-1;
    defs.forEach(d=>{
      let s=0;
      if(d.re.test(q))s+=2;
      if(d.intent.test(q))s+=1;
      if(s>score){score=s;best=d;}
    });
    if(!best || score<=0)return;
    const st=stats[best.name];
    st.impressions+=imp;st.clicks+=clk;
    const nq=skeNormalizeQuery_(q);
    const x=st.queries[nq]||(st.queries[nq]={q:q,imp:0});
    x.imp+=imp;
  });

  const rows=Object.values(stats)
    .filter(x=>x.impressions>0)
    .sort((a,b)=>b.impressions-a.impressions)
    .slice(0,8)
    .map((x,i)=>{
      const reps=Object.values(x.queries).sort((a,b)=>b.imp-a.imp).slice(0,5).map(y=>y.q).join(' / ');
      return [i+1,x.def.name,x.def.target,x.def.problem,reps,x.impressions,x.clicks,totalImp?x.impressions/totalImp:0,x.def.explore];
    });

  const sh=skeSheet_(SKE_CONFIG.sheets.persona);
  sh.clearContents();
  sh.getRange(1,1,1,9).setValues([['順位','検索ペルソナ','主な対象','主な困りごと','代表クエリ','表示回数','クリック','構成比','外部探索テーマ']]);
  if(rows.length)sh.getRange(2,1,rows.length,9).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,9).setFontWeight('bold').setWrap(true);
  if(rows.length){
    sh.getRange(2,8,rows.length,1).setNumberFormat('0.0%');
    sh.getRange(2,1,rows.length,9).setVerticalAlignment('top');
    [2,3,4,5,9].forEach(c=>sh.getRange(2,c,rows.length,1).setWrap(true));
  }
  sh.setColumnWidth(1,55);sh.setColumnWidth(2,180);sh.setColumnWidth(3,190);sh.setColumnWidth(4,220);
  sh.setColumnWidth(5,430);sh.setColumnWidth(6,100);sh.setColumnWidth(7,90);sh.setColumnWidth(8,90);sh.setColumnWidth(9,420);
  return rows;
}

function skeArticleMasterRequired_(){
  const rows=skeReadObjects_(SKE_CONFIG.sheets.articleMaster);
  if(!rows.length){
    throw new Error('v0.1.4では新記事候補のカニバリ防止のためArticle Masterが必須です。追加の操作 → Article Masterの使い方 から記事一覧を登録してください。');
  }
  return rows;
}

function skeOwnedQueryAssessment_(query, urls, articleRows){
  let best=null;
  articleRows.forEach(r=>{
    const url=skeNormalizeUrl_(skeObj_(r,['記事URL','URL','url'])||'');
    if(!url)return;
    const title=String(skeObj_(r,['記事タイトル','タイトル','title'])||'');
    const mq=String(skeObj_(r,['メインクエリ','Main Query','main_query'])||'');
    const intent=String(skeObj_(r,['SearchIntent','検索意図'])||'');
    const aid=String(skeObj_(r,['ArticleID','記事ID','article_id'])||'');
    let sim=Math.max(skeQuerySimilarity_(query,mq),skeTitleQueryCoverage_(title,query),skeQuerySimilarity_(query,intent));
    const observed=(urls||[]).some(u=>skeNormalizeUrl_(u.url)===url);
    if(observed)sim=Math.max(sim,.58);
    if(!best||sim>best.score)best={score:sim,url:url,articleId:aid,title:title,mainQuery:mq,observed:observed};
  });
  return best;
}

function skeNoveltySignal_(query, impressions, urlCount, owned){
  const q=skeNormalizeQuery_(query);
  const modifiers=/できない|消えた|消え|なくなった|追加できない|反応しない|エラー|error|変更|終了|廃止|アップデート後|新機能|代わり|直し方|未対応|対応機種|原因|なぜ/i;
  const hasModifier=modifiers.test(q);
  const unowned=!owned || owned.score<.42;
  const weakOwnership=owned && owned.score>=.42 && owned.score<.62;
  let type='NONE',score=0;
  if(unowned&&hasModifier){type='UNOWNED_INTENT_NEW_MODIFIER';score=2;}
  else if(unowned){type='UNOWNED_INTENT';score=1;}
  else if(weakOwnership&&hasModifier){type='INTENT_DRIFT';score=1;}
  return {type:type,score:score};
}

function skeRunInternalDiscovery() {
  skeSetup_();
  const siteId=skeGetSetting_('siteId');
  if(!siteId) throw new Error('先に「2. Evidenceを読み込む」を実行してください。');

  const qRows=skeReadObjects_(SKE_CONFIG.sheets.pageQuery);
  if(!qRows.length) throw new Error('page_query_top Evidenceがありません。');

  const articleRows=skeArticleMasterRequired_();
  const personaRows=skeBuildPersonaProfile_(qRows);

  const grouped={};
  qRows.forEach(r=>{
    const q=String(skeObj_(r,['query','クエリ'])||'').trim();
    const nq=skeNormalizeQuery_(q);
    const url=skeNormalizeUrl_(skeObj_(r,['page','url','URL'])||'');
    if(!q||!nq||!url)return;
    const g=grouped[nq]||(grouped[nq]={query:q,urls:{},clicks:0,impressions:0,posNum:0,posDen:0});
    const imp=Number(skeObj_(r,['impressions','表示回数'])||0);
    const clk=Number(skeObj_(r,['clicks','クリック数'])||0);
    const pos=Number(skeObj_(r,['position','掲載順位','平均掲載順位'])||0);
    g.clicks+=clk;g.impressions+=imp;
    if(pos>0&&imp>0){g.posNum+=pos*imp;g.posDen+=imp;}
    const u=g.urls[url]||(g.urls[url]={url:url,impressions:0,clicks:0,posNum:0,posDen:0});
    u.impressions+=imp;u.clicks+=clk;
    if(pos>0&&imp>0){u.posNum+=pos*imp;u.posDen+=imp;}
  });

  const all=Object.values(grouped).filter(g=>g.impressions>=3).sort((a,b)=>b.impressions-a.impressions);
  const maxImp=Math.max(1,...all.map(x=>x.impressions));
  const candidates=[];
  let ownedExcluded=0, noNoveltyExcluded=0;

  all.forEach(g=>{
    const urls=Object.values(g.urls).sort((a,b)=>b.impressions-a.impressions);
    const owned=skeOwnedQueryAssessment_(g.query,urls,articleRows);

    // OWNED_QUERY Gate: 既存記事が強く担当しているQueryはSKE新記事候補から除外
    if(owned && owned.score>=.62){
      ownedExcluded++;
      return;
    }

    const novelty=skeNoveltySignal_(g.query,g.impressions,urls.length,owned);
    if(novelty.score<=0){
      noNoveltyExcluded++;
      return;
    }

    const pos=g.posDen?g.posNum/g.posDen:0;
    const demand=Math.min(25,Math.round(25*Math.log1p(g.impressions)/Math.log1p(maxImp)));
    const noveltyScore=novelty.score===2?25:16;
    const ownershipGap=!owned||owned.score<.42?25:14;
    const posScore=pos>0&&pos<=20?12:pos<=40?8:5;
    const score=Math.min(100,demand+noveltyScore+ownershipGap+posScore+8);

    const relatedId=owned?owned.articleId||'':'';
    const existing=!owned||owned.score<.42?'VERIFIED_NO_STRONG_OWNER':'POSSIBLE_OVERLAP';
    const reason=
      `新規性=${novelty.type}。`+
      (owned?`最も近い既存記事との類似度=${Math.round(owned.score*100)}%。`:'近い既存記事なし。')+
      ` 既存記事が強く所有するQueryは除外済み。外部Web/SERPで需要と情報ギャップをDoctorが確認する。`;

    candidates.push({
      query:g.query,score:score,maturity:g.impressions>=50?'OBSERVED':g.impressions>=10?'EMERGING':'PREDICTED',
      existing:existing,relatedId:relatedId,urls:urls,impressions:g.impressions,clicks:g.clicks,pos:pos,
      urlCount:urls.length,reason:reason,novelty:novelty.type
    });
  });

  candidates.sort((a,b)=>b.score-a.score||b.impressions-a.impressions);
  const picked=candidates.slice(0,10);

  skeRemoveRegeneratableCandidates_(siteId);
  const out=picked.map(e=>[
    false,skeCandidateId_(siteId,e.query),siteId,skeGetSetting_('siteName'),e.query,'INTERNAL_PERSONA_GSC',
    e.score,e.maturity,'UNKNOWN',e.existing,e.relatedId,e.urls.slice(0,3).map(x=>x.url).join('\n'),
    'DOCTOR_REVIEW','DISCOVERED',e.impressions,e.clicks,e.pos,e.urlCount,e.reason,'','','','','',new Date()
  ]);

  const sh=skeSheet_(SKE_CONFIG.sheets.candidates);
  if(out.length){
    sh.getRange(sh.getLastRow()+1,1,out.length,SKE_CONFIG.candidateHeaders.length).setValues(out);
    sh.getRange(2,1,sh.getLastRow()-1,1).insertCheckboxes();
  }
  skeFormatCandidates_();
  skeRenderHome();

  SpreadsheetApp.getUi().alert(
    '内部探索が完了しました。\n\n'+
    `検索ペルソナ：${personaRows.length}グループ\n`+
    `OWNED_QUERY除外：${ownedExcluded}件\n`+
    `新規性不足除外：${noNoveltyExcluded}件\n`+
    `内部Blue Ocean候補：${out.length}件（最大10件）\n\n`+
    '次は「4. 検索ペルソナを確認する」で、このブログに来ている検索ニーズの集団を確認してください。'
  );
}

function skeOpenCandidates(){ const sh=skeSheet_(SKE_CONFIG.sheets.candidates); SpreadsheetApp.getActive().setActiveSheet(sh); }

function skeContinueWorkflow(){
  const rows=skeReadObjects_(SKE_CONFIG.sheets.candidates);
  const selected=rows.filter(r=>r['選択']===true || String(r['選択']).toLowerCase()==='true');
  if(selected.length){ skeGenerateDoctorPackageForSelected(); return; }
  const early=rows.filter(r=>String(r['状態']||'')==='EARLY_OPPORTUNITY');
  SpreadsheetApp.getUi().alert(early.length ? `再確認待ち候補が${early.length}件あります。\nキーワード候補シートで確認してください。` : '現在、選択中の候補や再確認待ち候補はありません。');
}

function skeGenerateDoctorPackageForSelected(){
  const sh=skeSheet_(SKE_CONFIG.sheets.candidates), vals=sh.getDataRange().getValues();
  if(vals.length<2) throw new Error('候補がありません。');
  const h=vals[0].map(String), ix={};h.forEach((x,i)=>ix[x]=i);
  const targets=[];
  for(let r=1;r<vals.length;r++) if(vals[r][ix['選択']]===true) targets.push({row:r+1,values:vals[r]});
  if(!targets.length) throw new Error('Doctorへ送る候補の「選択」にチェックを入れてください。');
  if(targets.length>3) throw new Error('1回にDoctorへ送れる候補は最大3件です。');
  const folder=DriveApp.getRootFolder(); const created=[];
  targets.forEach(t=>{
    const row=t.values, get=n=>row[ix[n]];
    const candidate={
      format:'SIMS_KEYWORD_EXPLORER_DOCTOR_REFERRAL_V1', contract_version:'0.1.4',
      identity:{candidate_id:String(get('Candidate ID')),site_id:String(get('SiteID')),site_name:String(get('ブログ'))},
      discovery:{type:String(get('Discovery Type')),primary_query:String(get('Primary Query')),demand_maturity:String(get('需要成熟度')),article_lifespan:String(get('記事寿命')),p1_score:Number(get('P1 Score')||0)},
      existing_article_check:{status:String(get('既存記事判定')),related_article_id:String(get('関連ArticleID')||''),related_urls:String(get('関連URL')||'').split(/\n+/).filter(Boolean)},
      evidence:{impressions:Number(get('表示回数')||0),clicks:Number(get('クリック')||0),average_position:Number(get('平均順位')||0),url_count:Number(get('URL数')||0),reason:String(get('発見理由')||'')},
      requested_decision:['GREEN','YELLOW','BLOCK'],
      instructions:{green:'Creatorへ新記事候補として紹介',yellow:'EARLY_OPPORTUNITYとして再確認条件・再確認日を提示',block:'新記事非推奨。既存記事改善が適切ならWriter振替を提示'},
      note:'P1は内部GSC探索のみ。SERP Gapと外部Web EvidenceはDoctor側で独立確認してください。'
    };
    const evidenceCsv=skeCandidateEvidenceCsv_(String(get('Primary Query')));
    const readme=[
      'SIMS Keyword Explorer Doctor Package',
      '',`Candidate: ${get('Candidate ID')}`,`Site: ${get('ブログ')}`,`Primary Query: ${get('Primary Query')}`,
      '', 'このZIPをSIMS Doctorへ渡してください。',
      'P1では外部Web探索・SERP Gap自動判定は未実装です。Doctorは現在のWeb/SERPを独立して確認してください。'
    ].join('\n');
    const files=[Utilities.newBlob(JSON.stringify(candidate,null,2),'application/json','doctor_referral.json'),Utilities.newBlob(evidenceCsv,'text/csv','candidate_evidence.csv'),Utilities.newBlob(readme,'text/plain','README-FIRST.md')];
    const fn=skePackageFileName_(String(get('ブログ')),'Doctor用',String(get('Candidate ID')));
    const f=folder.createFile(Utilities.zip(files,fn)); created.push({name:fn,url:f.getUrl()});
  });
  const msg=created.map(x=>`ファイル：${x.name}\n用途：SIMS Doctorへ渡してください\n保存先：マイドライブ`).join('\n\n');
  SpreadsheetApp.getUi().alert(`Doctor用Packageを作成しました。\n\n${msg}\n\n次の操作：生成したZIPをSIMS Doctorへ渡してください。`);
}

function skeCandidateEvidenceCsv_(query){
  const nq=skeNormalizeQuery_(query), rows=skeReadObjects_(SKE_CONFIG.sheets.pageQuery).filter(r=>skeNormalizeQuery_(skeObj_(r,['query','クエリ'])||'')===nq);
  const vals=[['page','query','clicks','impressions','ctr','position']];
  rows.forEach(r=>vals.push([skeObj_(r,['page','url'])||'',skeObj_(r,['query'])||'',skeObj_(r,['clicks'])||0,skeObj_(r,['impressions'])||0,skeObj_(r,['ctr'])||0,skeObj_(r,['position'])||0]));
  return vals.map(row=>row.map(skeCsvCell_).join(',')).join('\r\n');
}

function skeArticleMasterHelp(){
  const sh=skeSheet_(SKE_CONFIG.sheets.articleMaster); if(sh.isSheetHidden())sh.showSheet(); SpreadsheetApp.getActive().setActiveSheet(sh);
  SpreadsheetApp.getUi().alert('v0.1.4ではArticle Masterは必須です。既存記事とのカニバリを避けるため、新記事候補探索の前に登録してください。\n\n列：ArticleID / 記事タイトル / 記事URL / メインクエリ / SearchIntent / 状態\n\nSBM等から取得できる記事情報を2行目以降へ貼り付けてください。');
}

function skeBuildArticleMasterMap_(){
  const rows=skeReadObjects_(SKE_CONFIG.sheets.articleMaster), map={};
  rows.forEach(r=>{const url=skeNormalizeUrl_(skeObj_(r,['記事URL','URL','url'])||'');if(!url)return;map[url]={articleId:String(skeObj_(r,['ArticleID','記事ID','article_id'])||''),title:String(skeObj_(r,['記事タイトル','タイトル','title'])||''),mainQuery:String(skeObj_(r,['メインクエリ','Main Query','main_query'])||''),searchIntent:String(skeObj_(r,['SearchIntent','検索意図'])||'')};});
  return map;
}

function skeFormatCandidates_(){
  const sh=skeSheet_(SKE_CONFIG.sheets.candidates);
  if(sh.getLastRow()<1)return;

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,SKE_CONFIG.candidateHeaders.length)
    .setFontWeight('bold').setWrap(true);
  if(sh.getLastRow()>=2){
    sh.getRange(2,1,sh.getLastRow()-1,SKE_CONFIG.candidateHeaders.length)
      .setVerticalAlignment('top');
    [5,12,19].forEach(c=>sh.getRange(2,c,sh.getLastRow()-1,1).setWrap(true));
  }

  sh.autoResizeColumns(1,SKE_CONFIG.candidateHeaders.length);
  sh.setColumnWidth(1,55);
  sh.setColumnWidth(4,120);
  sh.setColumnWidth(5,300);
  sh.setColumnWidth(7,80);
  sh.setColumnWidth(8,105);
  sh.setColumnWidth(10,140);
  sh.setColumnWidth(12,240);
  sh.setColumnWidth(13,135);
  sh.setColumnWidth(14,110);
  sh.setColumnWidth(19,420);

  // 利用者が通常判断に使わない内部列は非表示。データ自体は保持する。
  const hiddenHeaders=['Candidate ID','SiteID','Discovery Type','記事寿命','関連ArticleID','URL数','更新日時'];
  const header=SKE_CONFIG.candidateHeaders;
  hiddenHeaders.forEach(name=>{
    const col=header.indexOf(name)+1;
    if(col>0) sh.hideColumns(col);
  });
}

function skeNormalizeQuery_(q){return String(q||'').toLowerCase().replace(/[　\s]+/g,' ').replace(/[｜|／/・,，。!！?？:：;；()[\]【】「」『』]/g,' ').replace(/\s+/g,' ').trim();}
function skeQueryTokens_(q){return skeNormalizeQuery_(q).split(' ').map(x=>x.trim()).filter(x=>x&&x.length>=2).filter(x=>!/^(202[0-9]|20[0-9]{2}|最新版|最新|完全版|方法|やり方)$/.test(x));}
function skeQuerySimilarity_(a,b){const na=skeNormalizeQuery_(a),nb=skeNormalizeQuery_(b);if(!na||!nb)return 0;if(na===nb)return 1;if(na.indexOf(nb)>=0||nb.indexOf(na)>=0)return .85;const ta=skeQueryTokens_(na),tb=skeQueryTokens_(nb);if(!ta.length||!tb.length)return 0;const sa={};ta.forEach(x=>sa[x]=1);let common=0;tb.forEach(x=>{if(sa[x])common++});const union={};ta.concat(tb).forEach(x=>union[x]=1);return common/Math.max(Object.keys(union).length,1);}
function skeTitleQueryCoverage_(title,q){const nt=skeNormalizeQuery_(title),nq=skeNormalizeQuery_(q);if(!nt||!nq)return 0;if(nt.indexOf(nq)>=0)return 1;const terms=skeQueryTokens_(nq);if(!terms.length)return 0;let hit=0;terms.forEach(t=>{if(nt.indexOf(t)>=0)hit++});return hit/terms.length;}
function skeNormalizeUrl_(u){return String(u||'').trim().replace(/[?#].*$/,'').replace(/\/$/,'');}
function skeSiteNameFromUrl_(u){try{return new URL(String(u)).hostname.replace(/^www\./,'');}catch(e){return '';}}
function skeSiteIdFromUrl_(u){const s=String(u||'').toLowerCase().replace(/^https?:\/\//,'').replace(/^sc-domain:/,'').replace(/^www\./,'').split('/')[0].replace(/[^a-z0-9.-]+/g,'-');return s||'site';}
function skeCandidateId_(siteId,q){const raw=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,siteId+'|'+skeNormalizeQuery_(q));const hex=raw.map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join('').slice(0,8).toUpperCase();return `SKE-${Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Tokyo','yyyyMMdd')}-${hex}`;}
function skePackageFileName_(site,purpose,cid){const safe=String(site||'site').replace(/[\\/:*?"<>|]/g,'-').replace(/\s+/g,'').slice(0,30);const c=String(cid||'').replace(/[^A-Za-z0-9_-]/g,'').slice(-16);const ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Tokyo','yyyyMMdd-HHmm');return `SKE-${safe}-${purpose}-${c}-${ts}.zip`;}
function skeExtractDriveId_(s){const x=String(s||'').trim();const m=x.match(/[-\w]{20,}/);return m?m[0]:'';}
function skeImportCsvBlobToSheet_(blob,name){const text=blob.getDataAsString('UTF-8').replace(/^\uFEFF/,'');const vals=Utilities.parseCsv(text);if(!vals.length)throw new Error(`${name} が空です。`);const sh=skeSheet_(name);sh.clearContents();const width=Math.max.apply(null,vals.map(r=>r.length));const norm=vals.map(r=>{const a=r.slice();while(a.length<width)a.push('');return a;});sh.getRange(1,1,norm.length,width).setValues(norm);}
function skeReadObjects_(name){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const v=sh.getDataRange().getValues(),h=v[0].map(String);return v.slice(1).map(r=>{const o={};h.forEach((x,i)=>o[x]=r[i]);return o;});}
function skeObj_(o,keys){for(let i=0;i<keys.length;i++){const k=keys[i];if(o&&Object.prototype.hasOwnProperty.call(o,k)&&o[k]!==''&&o[k]!=null)return o[k];}return '';}
function skeSheet_(name){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(!sh)throw new Error(`シートがありません: ${name}`);return sh;}
function skeSetSetting_(k,v){const sh=skeSheet_(SKE_CONFIG.sheets.settings),vals=sh.getDataRange().getValues();for(let i=1;i<vals.length;i++)if(String(vals[i][0])===k){sh.getRange(i+1,2).setValue(v);return;}sh.appendRow([k,v]);}
function skeGetSetting_(k){const rows=skeReadObjects_(SKE_CONFIG.sheets.settings);const r=rows.find(x=>String(x['Key'])===k);return r?String(r['Value']||''):'';}
function skeCsvCell_(v){const s=String(v==null?'':v);return /[",\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;}
