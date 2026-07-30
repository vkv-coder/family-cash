// ============================================================
// FAMILY CASH MANAGER — DEMO Backend
// Deploy as a SEPARATE Apps Script project + Web App (Execute as Me |
// Anyone can access) — completely isolated from the real production
// script/spreadsheet. It uses its own spreadsheet ("FamilyCashManager -
// DEMO"), created automatically on first call, so it can never touch
// real family data.
//
// One extra public action vs. the production script: `demoLogin`, which
// returns a seeded Family Head user + a real session token immediately,
// no OTP/email step. Everything else (session tokens, ownership checks
// on writes) works exactly like the hardened production script — see
// Code.gs in the same repo for the shared logic/comments.
// ============================================================

var SHEET_NAME = "FamilyCashManager - DEMO";
var SESSION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;

function getOrCreateSpreadsheet() {
  var files = DriveApp.getFilesByName(SHEET_NAME);
  var ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    setupSheets(ss);
    seedDemoData(ss);
  }
  ensureSessionsSheet(ss);
  return ss;
}

function setupSheets(ss) {
  ss.getSheets()[0].setName("Users");
  ["Accounts","Transactions","Reminders","SubCategories"].forEach(function(n){
    if(!ss.getSheetByName(n)) ss.insertSheet(n);
  });
  ss.getSheetByName("Users").appendRow(["id","name","email","isHead","cashOnHand","createdAt"]);
  ss.getSheetByName("Accounts").appendRow(["id","userId","type","name","last4","balance","createdAt"]);
  ss.getSheetByName("Transactions").appendRow(["id","userId","date","fromAccountId","toAccountId","amount","isTransfer","mainCat","subCat","note","createdAt"]);
  ss.getSheetByName("Reminders").appendRow(["id","userId","name","amount","dueDay","accountId","active","createdAt"]);
  ss.getSheetByName("SubCategories").appendRow(["userId","mainCat","subCat"]);
  var defCats = {
    Personal:["Food & Dining","Transport & Fuel","Shopping & Clothes","Bills & Utilities","EMI & Loans","Health & Medical","Entertainment","Salary & Income","Other"],
    Business:["Office Supplies","Travel & Stay","Client Entertainment","Software & Tools","Salary Paid","Business Income","Other"],
    Family:["Groceries","School Fees","Family Outing","Home Maintenance","Medical","Other"]
  };
  var sub = ss.getSheetByName("SubCategories");
  Object.keys(defCats).forEach(function(cat){
    defCats[cat].forEach(function(s){ sub.appendRow(["default",cat,s]); });
  });
}

function ensureSessionsSheet(ss) {
  var sh = ss.getSheetByName("Sessions");
  if (!sh) {
    sh = ss.insertSheet("Sessions");
    sh.appendRow(["token","userId","createdAt"]);
  }
}

// ── Demo seed data ────────────────────────────────────────
// Fixed ids (not Date.now()-based) so re-seeding is deterministic and
// idempotent — always the same two people, same starting balances.
var DEMO_HEAD_ID = "demo-head";
var DEMO_MEMBER_ID = "demo-member";

function clearSheetRows(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();
}

// Wipes and re-populates all demo data. Safe to run anytime (manually,
// or on a nightly trigger — see installNightlyReset() below).
function seedDemoData(ssParam) {
  var ss = ssParam || getOrCreateSpreadsheet();
  var usersSh = ss.getSheetByName("Users");
  var acctSh = ss.getSheetByName("Accounts");
  var txnSh = ss.getSheetByName("Transactions");
  var remSh = ss.getSheetByName("Reminders");
  var sessSh = ss.getSheetByName("Sessions");

  clearSheetRows(usersSh);
  clearSheetRows(acctSh);
  clearSheetRows(txnSh);
  clearSheetRows(remSh);
  if (sessSh) clearSheetRows(sessSh);

  var now = new Date().toISOString();
  usersSh.appendRow([DEMO_HEAD_ID, "Priya (Head)", "demo.head@example.com", "true", 2500, now]);
  usersSh.appendRow([DEMO_MEMBER_ID, "Rahul", "demo.member@example.com", "false", 800, now]);

  var accts = [
    ["a-demo-1", DEMO_HEAD_ID, "bank", "HDFC Savings", "4521", 145000, now],
    ["a-demo-2", DEMO_HEAD_ID, "card", "SBI Credit Card", "8890", -12500, now],
    ["a-demo-3", DEMO_MEMBER_ID, "bank", "ICICI Savings", "2210", 58000, now]
  ];
  accts.forEach(function(r){ acctSh.appendRow(r); });

  var cats = [
    ["Personal","Food & Dining"], ["Personal","Transport & Fuel"], ["Personal","Bills & Utilities"],
    ["Personal","EMI & Loans"], ["Personal","Salary & Income"], ["Family","Groceries"],
    ["Family","School Fees"], ["Business","Client Entertainment"]
  ];
  var txns = [];
  var today = new Date();
  for (var i = 0; i < 18; i++) {
    var d = new Date(today); d.setDate(d.getDate() - i * 2);
    var cat = cats[i % cats.length];
    var isIncome = cat[1] === "Salary & Income";
    var amt = isIncome ? (45000 + (i % 3) * 1000) : -(150 + i * 37);
    var userId = i % 3 === 0 ? DEMO_MEMBER_ID : DEMO_HEAD_ID;
    var accId = userId === DEMO_HEAD_ID ? "a-demo-1" : "a-demo-3";
    txns.push(["t-demo-" + i, userId, d.toISOString().slice(0,10), accId, "", amt, "false", cat[0], cat[1], "", now]);
  }
  txns.forEach(function(r){ txnSh.appendRow(r); });

  remSh.appendRow(["r-demo-1", DEMO_HEAD_ID, "Home Loan EMI", 18500, 5, "a-demo-1", "true", now]);
  remSh.appendRow(["r-demo-2", DEMO_HEAD_ID, "Netflix", 649, 12, "a-demo-1", "true", now]);
  remSh.appendRow(["r-demo-3", DEMO_MEMBER_ID, "Gym Membership", 1200, 1, "a-demo-3", "true", now]);
}

// Run this ONCE manually from the Apps Script editor (select it in the
// function dropdown, click Run, approve permissions) to set up a nightly
// reset at 2am so the demo never accumulates visitor junk.
function installNightlyReset() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === "seedDemoData") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("seedDemoData").timeBased().everyDays(1).atHour(2).create();
}

function makeResp(data) {
  return ContentService.createTextOutput(JSON.stringify({status:"ok",data:data})).setMimeType(ContentService.MimeType.JSON);
}
function makeErr(msg) {
  return ContentService.createTextOutput(JSON.stringify({status:"error",message:msg})).setMimeType(ContentService.MimeType.JSON);
}

var PUBLIC_ACTIONS = ["login", "verifyLoginOtp", "demoLogin"];

function doGet(e) {
  try {
    var ss = getOrCreateSpreadsheet();
    var a = e.parameter.action;
    var authedUserId = requireSession(ss, e.parameter.token);
    if(a==="getUserData") return makeResp(getUserData(ss, e.parameter, authedUserId));
    if(a==="getFamily")   return makeResp(getFamily(ss, e.parameter, authedUserId));
    return makeErr("Unknown GET action");
  } catch(ex){ return makeErr(ex.message); }
}

function doPost(e) {
  try {
    var ss = getOrCreateSpreadsheet();
    var b = JSON.parse(e.postData.contents);
    var a = b.action;
    var authedUserId = null;
    if (PUBLIC_ACTIONS.indexOf(a) === -1) {
      authedUserId = requireSession(ss, b.token);
    }
    if(a==="demoLogin")        return makeResp(demoLogin(ss));
    if(a==="saveAccount")      return makeResp(saveAccount(ss,b,authedUserId));
    if(a==="updateBalance")    return makeResp(updateBalance(ss,b,authedUserId));
    if(a==="updateCashOnHand") return makeResp(updateCashOnHand(ss,b,authedUserId));
    if(a==="saveTransaction")  return makeResp(saveTransaction(ss,b,authedUserId));
    if(a==="saveReminder")     return makeResp(saveReminder(ss,b,authedUserId));
    if(a==="updateReminder")   return makeResp(updateReminder(ss,b,authedUserId));
    if(a==="deleteReminder")   return makeResp(deleteReminder(ss,b,authedUserId));
    if(a==="addMember")        return makeResp(addMember(ss,b,authedUserId));
    if(a==="removeMember")     return makeResp(removeMember(ss,b,authedUserId));
    if(a==="addSubCat")        return makeResp(addSubCat(ss,b,authedUserId));
    if(a==="removeSubCat")     return makeResp(removeSubCat(ss,b,authedUserId));
    if(a==="updateName")       return makeResp(updateName(ss,b,authedUserId));
    return makeErr("Unknown action: "+a);
  } catch(ex){ return makeErr(ex.message); }
}

// ── Demo instant login ────────────────────────────────────
function demoLogin(ss) {
  var users = toObjs(ss.getSheetByName("Users"));
  if (users.length === 0) { seedDemoData(ss); users = toObjs(ss.getSheetByName("Users")); }
  var head = users.find(function(u){ return String(u.id) === DEMO_HEAD_ID; }) || users[0];
  var token = createSession(ss, head.id);
  head.token = token;
  return head;
}

// ── Helpers (identical to production Code.gs) ─────────────
function toObjs(sheet) {
  var d=sheet.getDataRange().getValues();
  if(d.length<2) return [];
  var h=d[0];
  return d.slice(1).map(function(r){
    var o={};h.forEach(function(k,i){o[k]=r[i];});return o;
  });
}
function findRow(sheet, col, val) {
  var d=sheet.getDataRange().getValues();
  for(var i=1;i<d.length;i++) if(String(d[i][col])===String(val)) return i+1;
  return -1;
}
function colIdx(sheet, name) {
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].indexOf(name);
}
function setCell(sheet, rowNum, colName, val) {
  var c=colIdx(sheet,colName);
  if(c>=0&&rowNum>0) sheet.getRange(rowNum,c+1).setValue(val);
}
function getRowObj(sheet, idColName, id) {
  var row = findRow(sheet, colIdx(sheet, idColName), id);
  if (row < 0) return null;
  return toObjs(sheet).find(function(o){ return String(o[idColName]) === String(id); });
}

function createSession(ss, userId) {
  var token = Utilities.getUuid();
  ss.getSheetByName("Sessions").appendRow([token, userId, new Date().toISOString()]);
  return token;
}
function requireSession(ss, token) {
  if (!token) throw new Error("Missing session token — please log in again");
  var sessions = toObjs(ss.getSheetByName("Sessions"));
  var s = sessions.find(function(x){ return String(x.token) === String(token); });
  if (!s) throw new Error("Invalid or expired session — please log in again");
  var ageMs = Date.now() - new Date(s.createdAt).getTime();
  if (ageMs > SESSION_MAX_AGE_MS) throw new Error("Session expired — please log in again");
  return String(s.userId);
}
function requireHead(ss, authedUserId) {
  var users = toObjs(ss.getSheetByName("Users"));
  var u = users.find(function(x){ return String(x.id) === authedUserId; });
  if (!u || String(u.isHead) !== "true") throw new Error("Not authorized — Family Head only");
  return u;
}

function getUserData(ss, p, authedUserId) {
  var targetId = p.targetId || authedUserId;
  var users=toObjs(ss.getSheetByName("Users"));
  var requester=users.find(function(u){return String(u.id)===authedUserId;});
  if(!requester) throw new Error("User not found");
  if(targetId!==authedUserId && String(requester.isHead)!=="true") throw new Error("Not authorized");
  var target=users.find(function(u){return String(u.id)===targetId;});
  if(!target) throw new Error("Target not found");
  var accounts=toObjs(ss.getSheetByName("Accounts")).filter(function(a){return String(a.userId)===targetId;});
  var transactions=toObjs(ss.getSheetByName("Transactions")).filter(function(t){return String(t.userId)===targetId;}).slice(-500).reverse();
  var reminders=toObjs(ss.getSheetByName("Reminders")).filter(function(r){return String(r.userId)===targetId;});
  var allSubs=toObjs(ss.getSheetByName("SubCategories"));
  var subCats={};
  allSubs.filter(function(s){return s.userId==="default"||String(s.userId)===targetId;}).forEach(function(s){
    if(!subCats[s.mainCat])subCats[s.mainCat]=[];
    if(subCats[s.mainCat].indexOf(s.subCat)===-1)subCats[s.mainCat].push(s.subCat);
  });
  return {user:target, accounts:accounts, transactions:transactions, reminders:reminders, subCategories:subCats};
}

function getFamily(ss, p, authedUserId) {
  var users=toObjs(ss.getSheetByName("Users"));
  var req=users.find(function(u){return String(u.id)===authedUserId;});
  if(!req||String(req.isHead)!=="true") throw new Error("Not authorized");
  var accounts=toObjs(ss.getSheetByName("Accounts"));
  return users.map(function(u){
    var ua=accounts.filter(function(a){return String(a.userId)===String(u.id);});
    var banks=ua.filter(function(a){return a.type==="bank";}).reduce(function(s,a){return s+Number(a.balance);},0);
    var cards=ua.filter(function(a){return a.type==="card";}).reduce(function(s,a){return s+Number(a.balance);},0);
    var cash=Number(u.cashOnHand)||0;
    return {id:u.id,name:u.name,email:u.email,isHead:u.isHead,netWorth:banks+cards+cash,banks:banks,cards:cards,cash:cash};
  });
}

function saveAccount(ss, b, authedUserId) {
  var id="a"+Date.now();
  ss.getSheetByName("Accounts").appendRow([id,authedUserId,b.type,b.name||"",b.last4||"0000",Number(b.balance)||0,new Date().toISOString()]);
  return {id:id};
}
function updateBalance(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Accounts");
  var acc = getRowObj(sheet, "id", b.accountId);
  if (!acc) throw new Error("Account not found");
  if (String(acc.userId) !== authedUserId) throw new Error("Not authorized for this account");
  var row=findRow(sheet,colIdx(sheet,"id"),b.accountId);
  setCell(sheet,row,"balance",Number(b.balance));
  return "updated";
}
function updateCashOnHand(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Users");
  var row=findRow(sheet,colIdx(sheet,"id"),authedUserId);
  setCell(sheet,row,"cashOnHand",Number(b.cashOnHand)||0);
  return "updated";
}
function saveTransaction(ss, b, authedUserId) {
  var id="t"+Date.now();
  [b.fromAccountId, b.toAccountId].forEach(function(accId){
    if (accId && accId !== "cash") {
      var acc = getRowObj(ss.getSheetByName("Accounts"), "id", accId);
      if (!acc || String(acc.userId) !== authedUserId) throw new Error("Not authorized for account " + accId);
    }
  });
  ss.getSheetByName("Transactions").appendRow([
    id,authedUserId,b.date,b.fromAccountId,b.toAccountId||"",
    Number(b.amount),b.isTransfer?"true":"false",
    b.mainCat||"",b.subCat||"",b.note||"",new Date().toISOString()
  ]);
  if(b.fromAccountId && Number(b.fromDelta)!==0) applyDelta(ss,authedUserId,b.fromAccountId,Number(b.fromDelta));
  if(b.toAccountId   && Number(b.toDelta)!==0)   applyDelta(ss,authedUserId,b.toAccountId,  Number(b.toDelta));
  return {id:id};
}
function applyDelta(ss, userId, accId, delta) {
  if(!delta) return;
  if(accId==="cash") {
    var us=ss.getSheetByName("Users");
    var row=findRow(us,colIdx(us,"id"),userId);
    var cur=Number(us.getRange(row,colIdx(us,"cashOnHand")+1).getValue())||0;
    setCell(us,row,"cashOnHand",cur+delta);
  } else {
    var ac=ss.getSheetByName("Accounts");
    var row2=findRow(ac,colIdx(ac,"id"),accId);
    var cur2=Number(ac.getRange(row2,colIdx(ac,"balance")+1).getValue())||0;
    setCell(ac,row2,"balance",cur2+delta);
  }
}
function saveReminder(ss, b, authedUserId) {
  var id="r"+Date.now();
  ss.getSheetByName("Reminders").appendRow([id,authedUserId,b.name,Number(b.amount),Number(b.dueDay),b.accountId,"true",new Date().toISOString()]);
  return {id:id};
}
function updateReminder(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Reminders");
  var rem = getRowObj(sheet, "id", b.id);
  if (!rem) throw new Error("Reminder not found");
  if (String(rem.userId) !== authedUserId) throw new Error("Not authorized for this reminder");
  var row=findRow(sheet,colIdx(sheet,"id"),b.id);
  setCell(sheet,row,"active",b.active?"true":"false");
  return "updated";
}
function deleteReminder(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Reminders");
  var rem = getRowObj(sheet, "id", b.id);
  if (!rem) throw new Error("Reminder not found");
  if (String(rem.userId) !== authedUserId) throw new Error("Not authorized for this reminder");
  var row=findRow(sheet,colIdx(sheet,"id"),b.id);
  if(row>0) sheet.deleteRow(row);
  return "deleted";
}
function addMember(ss, b, authedUserId) {
  requireHead(ss, authedUserId);
  var users = toObjs(ss.getSheetByName("Users"));
  var email = (b.email||"").toLowerCase().trim();
  if (users.find(function(u){ return String(u.email) === email; })) throw new Error("Email already registered");
  var id="u"+Date.now();
  ss.getSheetByName("Users").appendRow([id,b.name,email,"false",0,new Date().toISOString()]);
  return {id:id,name:b.name,email:email};
}
function removeMember(ss, b, authedUserId) {
  requireHead(ss, authedUserId);
  var sheet=ss.getSheetByName("Users");
  var row=findRow(sheet,colIdx(sheet,"id"),b.memberId);
  if(row>0) sheet.deleteRow(row);
  return "removed";
}
function addSubCat(ss, b, authedUserId) {
  ss.getSheetByName("SubCategories").appendRow([authedUserId,b.mainCat,b.subCat]);
  return "added";
}
function removeSubCat(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("SubCategories");
  var d=sheet.getDataRange().getValues();
  for(var i=d.length-1;i>=1;i--){
    if(String(d[i][0])===authedUserId&&d[i][1]===b.mainCat&&d[i][2]===b.subCat){sheet.deleteRow(i+1);break;}
  }
  return "removed";
}
function updateName(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Users");
  var row=findRow(sheet,colIdx(sheet,"id"),authedUserId);
  setCell(sheet,row,"name",b.name);
  return "updated";
}
