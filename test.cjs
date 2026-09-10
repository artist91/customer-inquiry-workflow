const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
let status = 200, body, payload, calls = 0;
const ctx = vm.createContext({UrlFetchApp:{fetch(url,options){
  calls++; assert.equal(url,'https://api.openai.com/v1/responses');
  payload = JSON.parse(options.payload);
  return {getResponseCode:()=>status,getContentText:()=>JSON.stringify(body)};
}}});
vm.runInContext(fs.readFileSync(__dirname+'/Code.gs','utf8'),ctx);
const good = {category:'배송',tags:['배송조회'],priority:'보통',draft:'배송 상태 확인이 필요합니다.',review_notes:'실제 주문 조회 필요'};
body = {status:'completed',output:[{type:'message',content:[{type:'output_text',text:JSON.stringify(good)}]}]};
assert.equal(ctx.analyze_('배송 문의','정책',{key:'fake',model:'test'}).category,'배송');
assert.equal(payload.store,false);
assert.equal(payload.text.format.strict,true);
assert.equal(ctx.literal_('=IMPORTXML("https://example.com")').charAt(0),"'");
assert.equal(ctx.literal_('일반 답변'),'일반 답변');
assert.throws(()=>ctx.validate_({...good,category:'잘못된 분류'}),/형식/);
assert.throws(()=>ctx.validate_({...good,draft:''}),/형식/);
assert.throws(()=>ctx.validate_({...good,tags:Array(6).fill('태그')}),/형식/);
for (const state of ['검수 대기','승인','보류','오류']) {
  const row = Array(14).fill(''); row[2]='문의'; row[8]=state;
  assert.equal(Boolean(ctx.eligible_(row)),false);
}
for (const state of ['','접수']) {
  const row = Array(14).fill(''); row[2]='문의'; row[8]=state;
  assert.equal(Boolean(ctx.eligible_(row)),true);
}
body={status:'incomplete'};
assert.throws(()=>ctx.analyze_('문의','정책',{key:'fake'}),/미완료/);
body={status:'completed',output:[{type:'message',content:[{type:'refusal'}]}]};
assert.throws(()=>ctx.analyze_('문의','정책',{key:'fake'}),/거절/);
status=429;
assert.throws(()=>ctx.analyze_('문의','정책',{key:'fake'}),/HTTP 429/);
const before=calls;
assert.throws(()=>ctx.analyze_('x'.repeat(12001),'정책',{key:'fake'}),/12,000/);
assert.equal(calls,before);
console.log('PASS: 정상 응답, 스키마, 상태 필터, 수식 방어, 미완료·거절·HTTP 오류, 입력 길이');

// Exercise the real queue and audit writer with in-memory Sheets/API adapters.
function scenario(mode) {
  const row=['DUMMY-001','date','배송 문의','','','','','','접수','','','','',0];
  const logs=[];
  const sheet={getLastRow:()=>2,getRange(r,c,h=1,w=1){return {
    getValues:()=>[row.slice(c-1,c-1+w)],
    setValue:v=>{row[c-1]=v;},
    setValues:vs=>{vs[0].forEach((v,i)=>row[c-1+i]=v);},
    clearContent:()=>{row[c-1]='';}
  };}};
  const ss={getSheetByName:n=>n==='문의'?sheet:{getDataRange:()=>({getDisplayValues:()=>[['항목','정책'],['배송','확인 필요']]})}};
  ctx.book_=()=>ss;
  ctx.config_=()=>({key:'secret-not-for-log',model:'test'});
  ctx.Utilities={getUuid:()=> 'run-1'};
  ctx.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})};
  ctx.SpreadsheetApp={flush(){}};
  ctx.logSheet_=()=>({appendRow(v){
    if(mode==='log_failure' && logs.length===1) throw Error('disk');
    logs.push(v);
  }});
  ctx.analyze_=(inquiry,policy,cfg,trace)=>{
    Object.assign(trace,{responseId:'resp-test',inputTokens:12,outputTokens:34});
    if(mode==='error') throw Error('OpenAI HTTP 429');
    if(mode==='edited') row[2]='사람이 수정한 문의';
    return good;
  };
  if(mode==='log_failure') assert.throws(()=>ctx.processQueue(),/로그 기록 실패/);
  else ctx.processQueue();
  assert.equal(JSON.stringify(logs).includes('secret-not-for-log'),false);
  return {row,logs};
}
let s=scenario('ok');
assert.equal(s.row[8],'검수 대기');
assert.deepEqual(s.logs.map(r=>r[3]),['AI 요청 시작','AI 응답 수신','분류·태그·초안 저장']);
assert.equal(JSON.parse(s.logs[0][12]).inquiry,'배송 문의');
assert.equal(JSON.parse(s.logs[1][12]).output.draft,good.draft);
assert.equal(s.logs[1][9],'resp-test');
s=scenario('error'); assert.equal(s.row[8],'오류'); assert.equal(s.logs.at(-1)[3],'처리 오류');
s=scenario('edited'); assert.equal(s.row[6],''); assert.equal(s.logs.at(-1)[3],'결과 반영 취소');
s=scenario('log_failure'); assert.equal(s.row[6],''); assert.equal(s.row[8],'접수');
console.log('PASS: 실제 큐 경로의 로그 순서·입출력·오류·동시 편집·로그 실패 중단·키 제외');
