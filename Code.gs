// ============================================================
// FAMILY CASH MANAGER — Google Apps Script Backend v4 (OTP login + sessions)
// Deploy as Web App: Execute as Me | Anyone can access
// ============================================================
//
// v4 change (security fix): every mutating action used to trust a
// client-supplied `userId` directly, and that id ("u"+Date.now()) was
// never re-verified against anything server-side after the initial OTP
// login — so anyone who guessed/observed a userId had permanent, full
// read/write access with no further login. This version issues a real
// session token on OTP verification (stored server-side in a new
// "Sessions" sheet), and every action that acts "as a user" now derives
// that user from the token, not from whatever the client claims.

var SHEET_NAME = "FamilyCashManager";
var SESSION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

function getOrCreateSpreadsheet() {
  var files = DriveApp.getFilesByName(SHEET_NAME);
  var ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    setupSheets(ss);
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

// Idempotent — safe to call on both brand-new and pre-existing spreadsheets.
function ensureSessionsSheet(ss) {
  var sh = ss.getSheetByName("Sessions");
  if (!sh) {
    sh = ss.insertSheet("Sessions");
    sh.appendRow(["token","userId","createdAt"]);
  }
}

function makeResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify({status:"ok",data:data}))
    .setMimeType(ContentService.MimeType.JSON);
}
function makeErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({status:"error",message:msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

var PUBLIC_ACTIONS = ["login", "verifyLoginOtp"];

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
    if(a==="login")            return makeResp(login(ss,b));
    if(a==="verifyLoginOtp")   return makeResp(verifyLoginOtp(ss,b));
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

// ── Helpers ───────────────────────────────────────────────
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
  var obj = toObjs(sheet).find(function(o){ return String(o[idColName]) === String(id); });
  return obj;
}

// ── Sessions ──────────────────────────────────────────────
function createSession(ss, userId) {
  var token = Utilities.getUuid();
  ss.getSheetByName("Sessions").appendRow([token, userId, new Date().toISOString()]);
  return token;
}

// Returns the AUTHENTICATED userId derived purely from the token — never
// trust a client-supplied userId for "who is acting" after this point.
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

// ── LOGIN (step 1: send OTP) ──────────────────────────────
function login(ss, b) {
  var email=(b.email||"").toLowerCase().trim();
  if(!email) throw new Error("Email required");
  var sheet=ss.getSheetByName("Users");
  var users=toObjs(sheet);
  var user=users.find(function(u){return String(u.email)===email;});
  if(!user) {
    var name=(b.name||"").trim();
    if(!name) throw new Error("Name required for new user");
    // Don't create the row yet — wait until the OTP is verified, so an
    // unverified email can't be used to spam-create accounts.
  }
  var otp = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put("otp_" + email, otp, 300); // 5 min
  if (b.name) CacheService.getScriptCache().put("otpname_" + email, b.name.trim(), 300);
  MailApp.sendEmail(email, "Your Family Cash Manager login code",
    "Your login code is: " + otp + "\n\nThis code expires in 5 minutes. If you didn't request this, you can ignore this email.");
  return { otpSent: true };
}

// ── LOGIN (step 2: verify OTP, create user if new, issue session) ─
function verifyLoginOtp(ss, b) {
  var email=(b.email||"").toLowerCase().trim();
  var otp=(b.otp||"").trim();
  var cache = CacheService.getScriptCache();
  var cached = cache.get("otp_" + email);
  if (!cached || cached !== otp) throw new Error("Incorrect or expired code");
  cache.remove("otp_" + email);

  var sheet=ss.getSheetByName("Users");
  var users=toObjs(sheet);
  var user=users.find(function(u){return String(u.email)===email;});
  if (!user) {
    var name = cache.get("otpname_" + email) || "Member";
    cache.remove("otpname_" + email);
    var isHead = users.length === 0;
    var id = "u" + Date.now();
    sheet.appendRow([id, name, email, isHead ? "true" : "false", 0, new Date().toISOString()]);
    user = { id: id, name: name, email: email, isHead: isHead ? "true" : "false", cashOnHand: 0 };
  }
  var token = createSession(ss, user.id);
  user.token = token;
  return user;
}

// ── GET USER DATA ─────────────────────────────────────────
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

// ── GET FAMILY ────────────────────────────────────────────
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

// ── SAVE ACCOUNT ──────────────────────────────────────────
function saveAccount(ss, b, authedUserId) {
  var id="a"+Date.now();
  ss.getSheetByName("Accounts").appendRow([id,authedUserId,b.type,b.name||"",b.last4||"0000",Number(b.balance)||0,new Date().toISOString()]);
  return {id:id};
}

// ── UPDATE BALANCE ────────────────────────────────────────
function updateBalance(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Accounts");
  var acc = getRowObj(sheet, "id", b.accountId);
  if (!acc) throw new Error("Account not found");
  if (String(acc.userId) !== authedUserId) throw new Error("Not authorized for this account");
  var row=findRow(sheet,colIdx(sheet,"id"),b.accountId);
  setCell(sheet,row,"balance",Number(b.balance));
  return "updated";
}

// ── UPDATE CASH ON HAND ───────────────────────────────────
function updateCashOnHand(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Users");
  var row=findRow(sheet,colIdx(sheet,"id"),authedUserId);
  setCell(sheet,row,"cashOnHand",Number(b.cashOnHand)||0);
  return "updated";
}

// ── SAVE TRANSACTION ──────────────────────────────────────
function saveTransaction(ss, b, authedUserId) {
  var id="t"+Date.now();
  // Ownership check: any account referenced must belong to the authed user.
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
  // Apply balance deltas
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

// ── REMINDERS ─────────────────────────────────────────────
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

// ── MEMBERS (Family Head only) ────────────────────────────
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

// ── SUB CATEGORIES ────────────────────────────────────────
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

// ── UPDATE NAME ───────────────────────────────────────────
function updateName(ss, b, authedUserId) {
  var sheet=ss.getSheetByName("Users");
  var row=findRow(sheet,colIdx(sheet,"id"),authedUserId);
  setCell(sheet,row,"name",b.name);
  return "updated";
}
