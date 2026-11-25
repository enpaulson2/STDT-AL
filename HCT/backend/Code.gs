const SHEET_ID='1MNRdUsyfvuyCknf4g_I-pl_nAwRK0KQ2hdXAbzVGL8U';
function sheet(){return SpreadsheetApp.openById(SHEET_ID).getSheetByName('Results')}
function migrate(){const s=sheet();if(s.getRange('I1').getValue()==='Proposed Technology')s.deleteColumn(9);}
function doGet(){return jsonResponse({status:'ready'})}
function doPost(e){const lock=LockService.getScriptLock();try{lock.waitLock(10000);const d=JSON.parse(e.postData.contents),t=Number(d.team);if(![1,2,3].includes(t)||!Array.isArray(d.pools)||d.pools.length!==6||!d.pools.every(p=>Array.isArray(p)&&p.length>=1&&p.length<=2&&p.every(v=>typeof v==='string'))||!['Difficulty accessing public transportation','Difficulty accessing hospital services'].includes(d.issue))throw Error('Invalid framework');const s=sheet();if(s.getRange('I1').getValue()!=='Submitted At')throw Error('Run migrate first');const safe=v=>"'"+String(v).slice(0,2000);s.getRange(t+1,1,1,9).setValues([['Team '+t,safe(d.issue),...d.pools.map(p=>safe(p.join(', '))),new Date()]]);return jsonResponse({success:true})}catch(e){return jsonResponse({success:false})}finally{if(lock.hasLock())lock.releaseLock()}}
function jsonResponse(d){return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON)}
function onOpen(){SpreadsheetApp.getUi().createMenu('HCT Results').addItem('Visual dashboard','showDashboard').addToUi()}
function showDashboard(){const html=HtmlService.createHtmlOutputFromFile('Dashboard').setWidth(1200).setHeight(720);SpreadsheetApp.getUi().showModalDialog(html,'HCT Team Results')}
function getResults(){return sheet().getRange('A2:I4').getDisplayValues()}
function openDashboard(){return HtmlService.createHtmlOutputFromFile('Dashboard')}
