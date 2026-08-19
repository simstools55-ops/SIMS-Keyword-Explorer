/**
 * SIMS Keyword Explorer v0.1.0
 * P1 prototype: Internal Discovery from SIMS Site Collector Evidence.
 *
 * Scope:
 * - Select/import one site's Collector Evidence ZIP
 * - Analyze page/query evidence
 * - Optional Article Master matching
 * - Build Candidate Registry
 * - Generate Doctor referral ZIPs with user-recognizable names
 *
 * Not included in P1:
 * - External Web Discovery
 * - SERP Gap automation
 * - Automatic Creator execution
 */

const SKE_VERSION = '0.1.0';
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
    articleMaster: '_SKE_ARTICLE_MASTER'
  },
  candidateHeaders: [
    '選択','Candidate ID','SiteID','ブログ','Primary Query','Discovery Type',
    'P1 Score','需要成熟度','記事寿命','既存記事判定','関連ArticleID','関連URL',
    'Engine判定','状態','表示回数','クリック','平均順位','URL数','発見理由',
    'Doctor判定','Doctor確信度','次回確認日','公開ArticleID','公開URL','更新日時'
  ]
};

function onOpen() {
  skeSetup_();
  SpreadsheetApp.getUi().createMenu('SIMS Keyword Explorer')
    .addItem('1. 対象ブログ / Evidenceを読み込む', 'skeImportEvidencePrompt')
    .addItem('2. 新しいキーワード候補を探す', 'skeRunInternalDiscovery')
    .addItem('3. 候補を確認する', 'skeOpenCandidates')
    .addItem('4. 処置を進める', 'skeContinueWorkflow')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('追加の操作')
      .addItem('Article Masterの使い方', 'skeArticleMasterHelp')
      .addItem('選択候補のDoctor用ZIPを作る', 'skeGenerateDoctorPackageForSelected')
      .addItem('Homeを更新', 'skeRenderHome'));
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
    ['次の操作', siteName==='未選択' ? '1. 対象ブログ / Evidenceを読み込む' : '2. 新しいキーワード候補を探す']
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
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('対象ブログ / Evidenceを読み込む', 'SIMS Site Collectorが生成したEvidence ZIPのGoogle Drive URLまたはファイルIDを貼り付けてください。', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const id = skeExtractDriveId_(res.getResponseText());
  if (!id) throw new Error('Google DriveのファイルIDを判定できませんでした。');
  const r = skeImportEvidenceById_(id);
  ui.alert(`Evidenceを読み込みました。\n\nブログ：${r.siteName}\nファイル：${r.fileName}\nQuery行：${r.queryRows}\n\n次は「2. 新しいキーワード候補を探す」を実行してください。`);
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

function skeRunInternalDiscovery() {
  skeSetup_();
  const siteId=skeGetSetting_('siteId');
  if(!siteId) throw new Error('先に「1. 対象ブログ / Evidenceを読み込む」を実行してください。');
  const qRows=skeReadObjects_(SKE_CONFIG.sheets.pageQuery);
  if(!qRows.length) throw new Error('page_query_top Evidenceがありません。');
  const articleMap=skeBuildArticleMasterMap_();
  const grouped={};
  qRows.forEach(r=>{
    const q=String(skeObj_(r,['query','クエリ'])||'').trim(); const nq=skeNormalizeQuery_(q);
    const url=skeNormalizeUrl_(skeObj_(r,['page','url','URL'])||'');
    if(!q||!nq||!url) return;
    const g=grouped[nq]||(grouped[nq]={query:q,urls:{},clicks:0,impressions:0,posNum:0,posDen:0});
    const imp=Number(skeObj_(r,['impressions','表示回数'])||0), clk=Number(skeObj_(r,['clicks','クリック数'])||0), pos=Number(skeObj_(r,['position','掲載順位','平均掲載順位'])||0);
    g.clicks+=clk; g.impressions+=imp; if(pos>0&&imp>0){g.posNum+=pos*imp;g.posDen+=imp;}
    const u=g.urls[url]||(g.urls[url]={url,impressions:0,clicks:0,position:0}); u.impressions+=imp;u.clicks+=clk; if(pos>0)u.position=pos;
  });
  const all=Object.values(grouped).filter(g=>g.impressions>=3).sort((a,b)=>b.impressions-a.impressions);
  const maxImp=Math.max(1,...all.map(x=>x.impressions));
  const existingIds=new Set(skeReadObjects_(SKE_CONFIG.sheets.candidates).map(r=>String(r['Candidate ID']||'')));
  const out=[];
  all.slice(0,250).forEach(g=>{
    const urls=Object.values(g.urls).sort((a,b)=>b.impressions-a.impressions);
    const primaryUrl=urls[0] ? urls[0].url : '';
    const master=articleMap[primaryUrl]||null;
    const urlCount=urls.length;
    let existing='POSSIBLE_OVERLAP', decision='DOCTOR_REVIEW', status='DISCOVERED', relatedId='', reason='';
    if(urlCount>=2){
      existing='POSSIBLE_OVERLAP'; decision='WRITER_REDIRECT'; status='WRITER_REDIRECT'; reason='同じ検索意図が複数URLに分散しているため、新記事より既存記事整理を優先。';
    } else if(master){
      relatedId=master.articleId||'';
      const sim=Math.max(skeQuerySimilarity_(g.query,master.mainQuery||''),skeTitleQueryCoverage_(master.title||'',g.query));
      if(sim>=0.70){ existing='EXISTING_ARTICLE_FOUND'; decision='WRITER_REDIRECT'; status='WRITER_REDIRECT'; reason='既存記事がこの検索意図を強く担当している可能性が高い。'; }
      else if(sim<0.40){ existing='VERIFIED_NO_CONFLICT'; decision='DOCTOR_REVIEW'; reason='実表示がある一方、既存記事の主題との一致が弱く、専用記事の余地をDoctorで確認する価値がある。'; }
      else { existing='POSSIBLE_OVERLAP'; decision='DOCTOR_REVIEW'; reason='近い既存記事があるため、独立記事化の可否をDoctorで確認。'; }
    } else {
      existing='POSSIBLE_OVERLAP'; decision='DOCTOR_REVIEW'; reason='Article Master未確認のため重複判定は保留。Doctor診断前に既存記事情報の確認を推奨。';
    }
    const maturity=g.impressions>=50?'OBSERVED':g.impressions>=10?'EMERGING':'PREDICTED';
    const demand=Math.min(25,Math.round(25*Math.log1p(g.impressions)/Math.log1p(maxImp)));
    const pos=g.posDen?g.posNum/g.posDen:0;
    const posScore=pos>0&&pos<=20?20:pos<=40?12:6;
    const fit=master?18:12;
    const gapProxy=existing==='VERIFIED_NO_CONFLICT'?18:existing==='POSSIBLE_OVERLAP'?10:4;
    let score=Math.max(0,Math.min(100,demand+posScore+fit+gapProxy+10));
    if(decision==='WRITER_REDIRECT')score=Math.min(score,74);
    if(score<50 && decision==='DOCTOR_REVIEW'){decision='DROP';status='BLOCK';reason+=' P1 Scoreが低いため初版では候補外。';}
    const cid=skeCandidateId_(siteId,g.query);
    if(existingIds.has(cid)) return;
    out.push([false,cid,siteId,skeGetSetting_('siteName'),g.query,'INTERNAL_GSC',score,maturity,'UNKNOWN',existing,relatedId,urls.map(x=>x.url).join('\n'),decision,status,g.impressions,g.clicks,pos,urlCount,reason,'','', '', '', '', new Date()]);
  });
  const sh=skeSheet_(SKE_CONFIG.sheets.candidates);
  if(out.length) {
    sh.getRange(sh.getLastRow()+1,1,out.length,SKE_CONFIG.candidateHeaders.length).setValues(out);
    sh.getRange(2,1,sh.getLastRow()-1,1).insertCheckboxes();
  }
  skeFormatCandidates_(); skeRenderHome();
  SpreadsheetApp.getUi().alert(`内部探索が完了しました。\n新規候補：${out.length}件\n\n「3. 候補を確認する」で内容を確認してください。`);
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
      format:'SIMS_KEYWORD_EXPLORER_DOCTOR_REFERRAL_V1', contract_version:'0.1.0',
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
  SpreadsheetApp.getUi().alert('Article Masterは任意ですが、候補の重複判定精度が大きく上がります。\n\n列：ArticleID / 記事タイトル / 記事URL / メインクエリ / SearchIntent / 状態\n\nSBM等から取得できる記事情報を2行目以降へ貼り付けてください。');
}

function skeBuildArticleMasterMap_(){
  const rows=skeReadObjects_(SKE_CONFIG.sheets.articleMaster), map={};
  rows.forEach(r=>{const url=skeNormalizeUrl_(skeObj_(r,['記事URL','URL','url'])||'');if(!url)return;map[url]={articleId:String(skeObj_(r,['ArticleID','記事ID','article_id'])||''),title:String(skeObj_(r,['記事タイトル','タイトル','title'])||''),mainQuery:String(skeObj_(r,['メインクエリ','Main Query','main_query'])||''),searchIntent:String(skeObj_(r,['SearchIntent','検索意図'])||'')};});
  return map;
}

function skeFormatCandidates_(){
  const sh=skeSheet_(SKE_CONFIG.sheets.candidates); if(sh.getLastRow()<2)return;
  sh.setFrozenRows(1); sh.getRange(1,1,1,SKE_CONFIG.candidateHeaders.length).setFontWeight('bold').setWrap(true);
  sh.getRange(2,1,sh.getLastRow()-1,SKE_CONFIG.candidateHeaders.length).setVerticalAlignment('top');
  [5,12,19].forEach(c=>sh.getRange(2,c,sh.getLastRow()-1,1).setWrap(true));
  sh.autoResizeColumns(1,SKE_CONFIG.candidateHeaders.length); sh.setColumnWidth(5,260);sh.setColumnWidth(12,300);sh.setColumnWidth(19,360);
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
