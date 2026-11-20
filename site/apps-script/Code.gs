const GRAPH = {"35":[36,34,22,18,37],"27":[6,29,28,5,4,26],"6":[27,30,11,29,8,5,1,9,10],"40":[42,41,28,39,26],"30":[6,42,43,36,32,11,29,28],"42":[40,30,43,41,36,28],"43":[30,42,36],"41":[40,42,36,39,37],"36":[35,30,42,43,41,34,32,37],"34":[35,36,32,33,16,14,18],"32":[30,36,34,31,33,11,13,14],"12":[11,10,13,14],"31":[32,11,13],"2":[3,1,4,26,25],"33":[34,32,16,14],"16":[34,33,17,19,14,18],"17":[16,19,15,14],"3":[2,4,26],"11":[6,30,32,12,31,9,10,13],"7":[8,1],"29":[27,6,30,28],"28":[27,40,30,42,29,26],"19":[16,17,20,18],"8":[6,7,1,10],"5":[27,6,1,4],"1":[6,2,7,8,5,4],"4":[27,2,3,5,1,26],"9":[6,11,10],"10":[6,12,11,8,9],"15":[17,14],"13":[32,12,31,11,14],"14":[34,32,12,33,16,17,15,13],"24":[22,25,23],"22":[35,24,21,18,37,23],"21":[22,20,18],"20":[19,21,18],"18":[35,34,16,19,22,21,20],"39":[40,41,26,25,37,38],"26":[27,40,2,3,28,4,39,25],"25":[2,24,39,26,23,38],"37":[35,41,36,22,39,23,38],"23":[24,22,25,37,38],"38":[39,25,37,23]};
// Привяжите проект к Google Таблице. PIN храните в Script Properties под ключом JURY_PIN.
const STATE_SHEET='State', LOG_SHEET='Events', SCORE_SHEET='Scores';
function setup(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  for(const name of [STATE_SHEET,LOG_SHEET,SCORE_SHEET]) if(!ss.getSheetByName(name))ss.insertSheet(name);
  const s=ss.getSheetByName(STATE_SHEET);s.clear();s.getRange('A1').setValue('JSON state');s.getRange('A2').setValue(JSON.stringify(initialState()));
  const l=ss.getSheetByName(LOG_SHEET);l.clear();l.appendRow(['Время','Действие','Команда','Задача','Счёт 1','Счёт 2']);
  const o=ss.getSheetByName(SCORE_SHEET);o.clear();o.appendRow(['Команда','Баллы территорий','Штраф подсказок','Итого']);o.getRange('A2:A3').setValues([['Оранжевая'],['Зелёная']]);writeScores(initialState());
}
function initialState(){return {phase:'waiting',startedAt:null,teams:[{name:'Оранжевая',cards:[3],failed:[],hints:0,attempts:{},hinted:[]},{name:'Зелёная',cards:[19],failed:[],hints:0,attempts:{},hinted:[]}],territories:Array.from({length:43},(_,i)=>({id:i+1,owner:0,blocked:false})),events:[]}}
function loadState(){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATE_SHEET);if(!sh)throw Error('Сначала выполните setup()');return JSON.parse(sh.getRange('A2').getValue())}
function points(n){return n===43?12:n<=27?5:8}
function score(s,i){return s.territories.filter(x=>x.owner===i+1).reduce((a,x)=>a+points(x.id),0)-2*s.teams[i].hints}
function adjacent(s,i){const owned=s.territories.filter(x=>x.owner===i+1).map(x=>x.id),ids=new Set(owned.flatMap(n=>GRAPH[n]||[])),t=s.teams[i];return [...ids].filter(n=>{const x=s.territories[n-1];return !x.owner&&!x.blocked&&!t.cards.includes(n)&&!t.failed.includes(n)})}
function finishIfNeeded(s){if(s.phase!=='active')return;if(Date.now()-new Date(s.startedAt).getTime()>=5400000||s.teams.every((t,i)=>!t.cards.length&&!adjacent(s,i).length))s.phase='finished'}
function doGet(e){try{const s=loadState();finishIfNeeded(s);const cb=String(e.parameter.callback||'').replace(/[^a-zA-Z0-9_$.]/g,'');const json=JSON.stringify(s);return ContentService.createTextOutput(cb?cb+'('+json+')':json).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON)}catch(err){return ContentService.createTextOutput(JSON.stringify({error:String(err)})).setMimeType(ContentService.MimeType.JSON)}}
function doPost(e){const lock=LockService.getScriptLock();lock.waitLock(20000);try{const p=e.parameter;if(!PropertiesService.getScriptProperties().getProperty('JURY_PIN')||p.pin!==PropertiesService.getScriptProperties().getProperty('JURY_PIN'))throw Error('Неверный PIN');const s=loadState(),action=p.action,id=Number(p.id),team=Number(p.team),i=team-1,t=s.teams[i],x=s.territories[id-1];finishIfNeeded(s);
  if(action==='start'){if(s.phase!=='waiting')throw Error('Игра уже началась');s.phase='active';s.startedAt=new Date().toISOString()}
  else {if(s.phase!=='active')throw Error('Игра не активна');if(![1,2].includes(team)||!Number.isInteger(id)||id<1||id>43)throw Error('Неверная карточка');
    if(action==='take'){if(t.cards.length>=3||!adjacent(s,i).includes(id))throw Error('Нельзя взять карточку');t.cards.push(id)}
    else {if(!t.cards.includes(id))throw Error('Карточки нет на руках');
      if(action==='correct'){if(x.owner||x.blocked)throw Error('Уже занята');x.owner=team;t.cards=t.cards.filter(n=>n!==id);s.teams[1-i].cards=s.teams[1-i].cards.filter(n=>n!==id)}
      else if(action==='wrong'){t.attempts[id]=(t.attempts[id]||0)+1;if(t.attempts[id]>=(id===[3,19][i]?3:2)){t.cards=t.cards.filter(n=>n!==id);t.failed.push(id)}}
      else if(action==='hint'){if(!(t.attempts[id]>0)||t.hinted.includes(id))throw Error('Подсказка пока недоступна');t.hinted.push(id);t.hints++}
      else if(action==='drop'){if(id===[3,19][i])throw Error('Стартовую карточку нельзя сбросить');t.cards=t.cards.filter(n=>n!==id);t.failed.push(id)}
      else throw Error('Неизвестное действие');
    }
    if(!x.owner&&s.teams.every(z=>z.failed.includes(id)))x.blocked=true;
  }
  finishIfNeeded(s);SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATE_SHEET).getRange('A2').setValue(JSON.stringify(s));writeScores(s);SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET).appendRow([new Date(),action,team||'',id||'',score(s,0),score(s,1)]);return ContentService.createTextOutput('OK')
}catch(err){return ContentService.createTextOutput('ERROR: '+err)}finally{lock.releaseLock()}}
function writeScores(s){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCORE_SHEET);sh.getRange('A2:D3').setValues(s.teams.map((t,i)=>{const base=s.territories.filter(x=>x.owner===i+1).reduce((a,x)=>a+points(x.id),0);return [t.name,base,-2*t.hints,score(s,i)]}))}
