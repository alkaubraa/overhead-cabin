// Overhead Cabin API v5 (read and write) + Inner Circle and referral
// Order Log layout: A Customer, B Product, C Qty, D Variant, E Note, F Date, G Cancel, H Price, I Surcharge, J Manual IDR, K Store
// Customers layout added in v5: U Member?, V Referred by, W Referral status, X Cashback earned, Y (hidden engine),
//                               Z Used as credit, AA Transferred out, AB Cashback balance
// Keep the SAME secret key you used before.
const SECRET = "ApartemenPavilionTower1";

/* ---------- web entry points ---------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.key !== SECRET) return out_({ ok: false, error: "Wrong or missing key" });
    if (p.action === "ping") {
      return out_({ ok: true, message: "Connected to: " + SpreadsheetApp.getActiveSpreadsheet().getName() });
    }
    if (p.action === "snapshot") return out_({ ok: true, data: snapshot_() });
    return out_({ ok: false, error: "Unknown action" });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.key !== SECRET) return out_({ ok: false, error: "Wrong or missing key" });
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      const results = runBatch_(body.items || []);
      return out_({ ok: true, results: results, data: snapshot_() });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */
function clean_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return v;
}
function str_(v) { return String(v == null ? "" : v).trim(); }
function toDate_(s) {
  if (!s) return "";
  const p = String(s).split("-");
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}
function tab_(ss, name) {
  const s = ss.getSheetByName(name);
  if (!s) throw new Error("Missing tab: " + name);
  return s;
}
function firstEmpty_(sh, first, last, width, cols) {
  const v = sh.getRange(first, 1, last - first + 1, width).getValues();
  for (let i = 0; i < v.length; i++) {
    let empty = true;
    for (let k = 0; k < cols.length; k++) {
      if (v[i][cols[k]] !== "") { empty = false; break; }
    }
    if (empty) return first + i;
  }
  throw new Error("No empty row left in " + sh.getName());
}
function findRow_(sh, first, last, col, name) {
  const n = str_(name).toLowerCase();
  if (n === "") return 0;
  const v = sh.getRange(first, col, last - first + 1, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    if (str_(v[i][0]).toLowerCase() === n) return first + i;
  }
  return 0;
}
function has_(d, k) { return Object.prototype.hasOwnProperty.call(d, k); }

/* =====================================================================
   INNER CIRCLE + REFERRAL  (v5)
   Run setupInnerCircle() ONCE from the editor BEFORE deploying the new version.
   It is safe to run again: formulas are rewritten, but rates you already typed
   into Settings are kept.
   ===================================================================== */
const IC_HEADER = "INNER CIRCLE & REFERRAL";
const FMT = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;

function icReady_(ss) {
  try {
    return str_(tab_(ss, "Settings").getRange("B34").getValue()) === IC_HEADER &&
           str_(tab_(ss, "Customers").getRange("U5").getValue()) === "Member?" &&
           str_(tab_(ss, "Order Log").getRange("AN5").getValue()) === "Member?";
  } catch (err) { return false; }
}
function ensureCols_(sh, n) {
  const have = sh.getMaxColumns();
  if (have < n) sh.insertColumnsAfter(have, n - have);
}
function pxWidth_(w) { return Math.round(w * 7 + 5); }

// Base price from a yen amount. Regular and member versions.
function regBase_(y) { return "IF(" + y + "<Settings!$C$4," + y + "*Settings!$C$5+Settings!$C$6," + y + "*Settings!$C$7)"; }
function memBase_(y) { return "IF(" + y + "<Settings!$C$4," + y + "*Settings!$C$35+Settings!$C$36," + y + "*Settings!$C$37)"; }

function icOrderPrice_(r) {
  const yc = "INDEX(Products!$C$6:$C$389,L" + r + ")";
  const memCat = memBase_(yc) + "+N(INDEX(Products!$D$6:$D$389,L" + r + "))";
  return 'IF(B' + r + '="","",IF(L' + r + '<>"",' +
    'IF(INDEX(Products!$J$6:$J$389,L' + r + ')="","",' +
    'IF(AND(AN' + r + '="Yes",INDEX(Products!$E$6:$E$389,L' + r + ')="",N(' + yc + ')>0),' + memCat +
    ',INDEX(Products!$J$6:$J$389,L' + r + '))),' +
    'IF(J' + r + '<>"",J' + r + ',IF(H' + r + '="","",IF(AN' + r + '="Yes",' + memBase_("H" + r) + ',' + regBase_("H" + r) + ')+N(I' + r + ')))))';
}
function icMemberFlag_(r) {
  return 'IF(A' + r + '="","",IF(IFERROR(INDEX(Customers!$U$6:$U$205,MATCH(A' + r + ',Customers!$A$6:$A$205,0)),"")="Yes","Yes","No"))';
}
function icReferralStatus_(r) {
  return 'IF(OR(A' + r + '="",V' + r + '=""),"",' +
    'IF(COUNTIF($A$6:$A$205,V' + r + ')=0,"Referrer not found",' +
    'IF(LOWER(TRIM(V' + r + '))=LOWER(TRIM(A' + r + ')),"Cannot refer self",' +
    'IF(INDEX($U$6:$U$205,MATCH(V' + r + ',$A$6:$A$205,0))<>"Yes","Referrer is not a member",' +
    'IF(U' + r + '="Yes","Referred buyer is a member",' +
    'IF(P' + r + '<>"Delivered","Waiting delivery",' +
    'IF(ROUND(J' + r + '-K' + r + ',0)>0,"Waiting payment",' +
    'IF(J' + r + '<Settings!$C$39,"Below minimum","Earned"))))))))';
}

function setupInnerCircle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = tab_(ss, "Settings"), cu = tab_(ss, "Customers"), ol = tab_(ss, "Order Log"),
        pr = tab_(ss, "Products"), pa = tab_(ss, "Payments");
  ensureCols_(cu, 28);
  ensureCols_(ol, 40);
  setupSettingsIC_(st);
  setupCustomersIC_(cu);
  setupOrderLogIC_(ol);
  setupOtherTabsIC_(pr, pa);
  setupInvoiceIC_(ss);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Inner Circle setup is done. Next: Deploy > Manage deployments > edit > New version > Deploy.");
}

function setupSettingsIC_(st) {
  st.getRange("B22:D22").copyTo(st.getRange("B34:D34"), FMT, false);
  st.getRange("B34:D34").setValues([[IC_HEADER, "Value", "How it is used"]]);
  const rows = [
    [35, "Member Tier 1 rate (Rp per ¥)", "=C5", "Same as the regular Tier 1 rate. Type a number here only if members should get a different rate"],
    [36, "Member Tier 1 flat fee (Rp per item)", 35000, "Added once per item below the ¥ threshold, instead of the regular fee in C6"],
    [37, "Member Tier 2 rate (Rp per ¥)", 133, "Used from the ¥ threshold up, instead of the regular rate in C7"],
    [38, "Referral cashback (Rp per new buyer)", 20000, "Flat amount a member earns for each new buyer they referred"],
    [39, "Referral minimum order (Rp)", 200000, "Referred buyer needs at least this much billable items (after refunds), fully paid, with status Delivered"]
  ];
  rows.forEach(function (x) {
    const r = x[0];
    st.getRange("B5:D5").copyTo(st.getRange("B" + r + ":D" + r), FMT, false);
    st.getRange("B" + r).setValue(x[1]);
    st.getRange("D" + r).setValue(x[3]);
    const c = st.getRange("C" + r);
    if (r === 35) c.setFormula(x[2]);
    else if (str_(c.getValue()) === "") c.setValue(x[2]);
  });
  st.getRange("C35").setFontColor("#000000");
  st.getRange("B31").copyTo(st.getRange("B40"), FMT, false);
  st.getRange("B40").setValue("Member price applies to yen-priced items only. Rupiah all-in items and bulky surcharges never get the member rate.");
}

function setupCustomersIC_(cu) {
  const heads = [
    [21, "FILL IN", "Member?", 9, false, "B"],
    [22, "FILL IN", "Referred by", 18, false, "R"],
    [23, "AUTO", "Referral status", 22, false, "Q"],
    [24, "AUTO", "Cashback earned (Rp)", 14, false, "N"],
    [25, "ENGINE (hidden)", "Cashback value (engine)", 10, true, "S"],
    [26, "AUTO", "Used as credit (Rp)", 14, false, "N"],
    [27, "FILL IN", "Transferred out (Rp)", 14, false, "L"],
    [28, "AUTO", "Cashback balance (Rp)", 14, false, "N"]
  ];
  const head4 = { "FILL IN": "R4", "AUTO": "F4", "ENGINE (hidden)": "S4" };
  const money = "#,##0;[Red](#,##0);-";
  heads.forEach(function (h) {
    const c = h[0];
    cu.getRange(head4[h[1]]).copyTo(cu.getRange(4, c), FMT, false);
    cu.getRange(4, c).setValue(h[1]);
    cu.getRange("R5").copyTo(cu.getRange(5, c), FMT, false);
    cu.getRange(5, c).setValue(h[2]);
    cu.getRange(h[5] + "6:" + h[5] + "205").copyTo(cu.getRange(6, c, 200, 1), FMT, false);
    if (h[3] && !h[4]) { cu.showColumns(c); cu.setColumnWidth(c, pxWidth_(h[3])); }
    if (h[4]) cu.hideColumns(c);
  });
  [24, 26, 28].forEach(function (c) { cu.getRange(6, c, 200, 1).setNumberFormat(money); });

  const W = [], X = [], Y = [], Z = [], AB = [];
  for (let r = 6; r <= 205; r++) {
    W.push(["=" + icReferralStatus_(r)]);
    X.push(['=IF(OR(A' + r + '="",U' + r + '<>"Yes"),"",SUMIFS($Y$6:$Y$205,$V$6:$V$205,A' + r + '))']);
    Y.push(['=IF(A' + r + '="","",IF(W' + r + '="Earned",Settings!$C$38,0))']);
    Z.push(['=IF(OR(A' + r + '="",U' + r + '<>"Yes"),"",SUMIFS(Payments!$C$6:$C$505,Payments!$B$6:$B$505,A' + r + ',Payments!$E$6:$E$505,"Referral credit*"))']);
    AB.push(['=IF(OR(A' + r + '="",U' + r + '<>"Yes"),"",N(X' + r + ')-N(Z' + r + ')-N(AA' + r + '))']);
  }
  cu.getRange(6, 23, 200, 1).setFormulas(W);
  cu.getRange(6, 24, 200, 1).setFormulas(X);
  cu.getRange(6, 25, 200, 1).setFormulas(Y);
  cu.getRange(6, 26, 200, 1).setFormulas(Z);
  cu.getRange(6, 28, 200, 1).setFormulas(AB);

  cu.getRange("U6:U205").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Yes"], true).setAllowInvalid(false)
      .setHelpText("Yes = Inner Circle member. Leave blank for a regular customer.").build());
  cu.getRange("V6:V205").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(cu.getRange("A6:A205"), true).setAllowInvalid(true)
      .setHelpText("Only for a NEW buyer. Pick the Inner Circle member who referred them, before the invoice is made.").build());

  // conditional formats: drop any old rule that touches columns U and up, then add ours
  const keep = cu.getConditionalFormatRules().filter(function (rule) {
    return rule.getRanges().every(function (rg) { return rg.getColumn() < 21; });
  });
  const mk = function (a1, formula, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
      .setBackground(color).setRanges([cu.getRange(a1)]).build();
  };
  keep.push(mk("U6:U205", '=$U6="Yes"', "#E8DAEF"));
  keep.push(mk("W6:W205", '=$W6="Earned"', "#E2F0D9"));
  keep.push(mk("W6:W205", '=OR($W6="Referrer not found",$W6="Cannot refer self",$W6="Referrer is not a member",$W6="Referred buyer is a member")', "#FCE4E4"));
  keep.push(mk("W6:W205", '=OR($W6="Waiting delivery",$W6="Waiting payment",$W6="Below minimum")', "#FFEB9C"));
  cu.setConditionalFormatRules(keep);

  // total cashback still owed to members
  cu.getRange("AA3").setValue("Total owed").setFontFamily("Arial").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("right");
  cu.getRange("AB3").setFormula("=SUM(AB6:AB205)").setNumberFormat(money).setFontFamily("Arial").setFontSize(10).setFontWeight("bold");

  cu.getRange("A2").setValue(
    "One row per customer. Set Type: Trusted customers go into the shopping list at once, New customers only after they pay. " +
    "Inner Circle: set Member? = Yes and the member rate applies to all their yen items. " +
    "Referral: pick the member in \"Referred by\" for a new buyer. " +
    "Sample rows are examples: delete them before real orders.");
}

function setupOrderLogIC_(ol) {
  ol.getRange("AM4:AM5").copyTo(ol.getRange("AN4:AN5"), FMT, false);
  ol.getRange("AN4").setValue("ENGINE (hidden)");
  ol.getRange("AN5").setValue("Member?");
  ol.getRange("AM6:AM1005").copyTo(ol.getRange("AN6:AN1005"), FMT, false);
  const wrap = String(ol.getRange("O6").getFormula()).toUpperCase().indexOf("=ARRAYFORMULA(") === 0;
  const w = function (f) { return wrap ? "=ARRAYFORMULA(" + f + ")" : "=" + f; };
  const AN = [], O = [];
  for (let r = 6; r <= 1005; r++) {
    AN.push([w(icMemberFlag_(r))]);
    O.push([w(icOrderPrice_(r))]);
  }
  ol.getRange(6, 40, 1000, 1).setFormulas(AN);
  ol.getRange(6, 15, 1000, 1).setFormulas(O);
  ol.hideColumns(40);
}

// Invoice tab: a tag right after the customer name when they are an Inner Circle member
function setupInvoiceIC_(ss) {
  const inv = ss.getSheetByName("Invoice");
  if (!inv) return;
  inv.getRange("E4")
    .setFormula('=IF($C$4="","",IF(IFERROR(INDEX(Customers!$U$6:$U$205,MATCH($C$4,Customers!$A$6:$A$205,0)),"")="Yes","[with special Inner Circle rate]",""))')
    .setFontFamily("Arial").setFontSize(10).setFontWeight("bold").setFontColor("#CC4125").setHorizontalAlignment("left");
}

function setupOtherTabsIC_(pr, pa) {
  // Quick calc: regular price now uses "below the threshold" like the rest of the sheet, plus a member price
  pr.getRange("F2").setFormula('=IF(B2="","",' + regBase_("B2") + '+N(D2))');
  pr.getRange("E2").copyTo(pr.getRange("G2"), FMT, false);
  pr.getRange("G2").setValue("Member");
  pr.getRange("F2").copyTo(pr.getRange("H2"), FMT, false);
  pr.getRange("H2").setFormula('=IF(B2="","",' + memBase_("B2") + '+N(D2))');

  const note = String(pa.getRange("A2").getValue());
  if (note.indexOf("Referral credit") < 0) {
    pa.getRange("A2").setValue(note + ' To use a member\'s referral cashback as credit, log a payment with Method / Note = "Referral credit".');
  }
  pa.getRange("E6:E505").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Transfer", "Cash", "Referral credit"], true).setAllowInvalid(true).build());
}

/* ---------- purchases memory (product + variant key) ---------- */
function purchasesTab_(ss) {
  let s = ss.getSheetByName("Purchases");
  if (!s) s = ss.insertSheet("Purchases");
  if (str_(s.getRange("A1").getValue()) !== "Product" ||
      str_(s.getRange("B1").getValue()) !== "Variant") {
    s.getRange(1, 1, 1, 6).setValues([["Product", "Variant", "Bought Qty", "Actual Paid (¥)", "Can't get more?", "Updated at"]]);
  }
  return s;
}
function memReady_(ss) {
  return str_(purchasesTab_(ss).getRange("H1").getValue()) === "READY_VARIANT";
}
function purchaseKey_(product, variant) {
  return str_(product).toLowerCase() + "\u001f" + (str_(variant) || "standard").toLowerCase();
}
function readMem_(pu) {
  const last = pu.getLastRow(), mem = {};
  if (last < 2) return mem;
  pu.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) {
    const product = str_(r[0]), variant = str_(r[1]) || "standard";
    if (product) mem[purchaseKey_(product, variant)] = { b: r[2], c: r[3], d: r[4] };
  });
  return mem;
}
function upsertMem_(ss, product, variant, b, c, d) {
  const pu = purchasesTab_(ss), last = pu.getLastRow(), stamp = new Date().toISOString();
  const key = purchaseKey_(product, variant);
  if (last >= 2) {
    const rows = pu.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (purchaseKey_(rows[i][0], rows[i][1]) === key) {
        pu.getRange(2 + i, 3, 1, 3).setValues([[b, c, d]]);
        pu.getRange(2 + i, 6).setValue(stamp);
        return;
      }
    }
  }
  pu.getRange(last + 1, 1, 1, 6).setValues([[str_(product), str_(variant) || "standard", b, c, d, stamp]]);
}
// Initializes/migrates the saved purchase memory to a product+variant key.
function setupPurchasesMemory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), pu = purchasesTab_(ss), sl = tab_(ss, "Shopping List");
  const values = sl.getRange(7, 1, 300, 9).getValues();
  const existing = readMem_(pu), rows = [];
  values.forEach(function (r) {
    const product = str_(r[0]), variant = str_(r[1]) || "standard";
    if (!product) return;
    const old = existing[purchaseKey_(product, variant)];
    const bought = r[6] !== "" ? r[6] : (old ? old.b : "");
    const paid = r[7] !== "" ? r[7] : (old ? old.c : "");
    const sold = r[8] !== "" ? r[8] : (old ? old.d : "");
    rows.push([product, variant, bought, paid, sold, new Date().toISOString()]);
  });
  if (pu.getLastRow() >= 2) pu.getRange(2, 1, pu.getLastRow() - 1, 6).clearContent();
  if (rows.length) pu.getRange(2, 1, rows.length, 6).setValues(rows);
  pu.getRange("H1").setValue("READY_VARIANT");
  SpreadsheetApp.flush();
}
function applyPurchasesToShopping_(ss) {
  if (!memReady_(ss)) return;
  const sl = tab_(ss, "Shopping List"), values = sl.getRange(7, 1, 300, 2).getValues();
  const mem = readMem_(purchasesTab_(ss));
  const out = values.map(function (r) {
    const product = str_(r[0]), variant = str_(r[1]) || "standard";
    const m = product ? mem[purchaseKey_(product, variant)] : null;
    return m ? [m.b, m.c, m.d] : ["", "", ""];
  });
  sl.getRange(7, 7, 300, 3).setValues(out);
}
/* ---------- typing directly in the sheet keeps memory in step ---------- */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet(), ss = sh.getParent(), name = sh.getName();
    if (name !== "Shopping List" && name !== "Order Log") return;
    if (!memReady_(ss)) return;
    const r1 = e.range.getRow(), r2 = e.range.getLastRow();
    const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      SpreadsheetApp.flush();
      if (name === "Shopping List") {
        if (c2 < 7 || c1 > 9 || r2 < 7 || r1 > 306) return;
        const from = Math.max(r1, 7), to = Math.min(r2, 306);
        sh.getRange(from, 1, to - from + 1, 9).getValues().forEach(function (v) {
          const product = str_(v[0]), variant = str_(v[1]) || "standard";
          if (product) upsertMem_(ss, product, variant, v[6], v[7], v[8]);
        });
      } else if (r2 >= 6) {
        applyPurchasesToShopping_(ss);
      }
    } finally { lock.releaseLock(); }
  } catch (err) { /* never block typing */ }
}

/* ---------- batch runner (safe to retry: every item has its own id) ---------- */
function logTab_(ss) {
  let s = ss.getSheetByName("AppLog");
  if (!s) {
    s = ss.insertSheet("AppLog");
    s.getRange(1, 1, 1, 4).setValues([["cid", "action", "row", "at"]]);
  }
  return s;
}
function runBatch_(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!memReady_(ss)) setupPurchasesMemory();
  const lg = logTab_(ss);
  const lastLog = lg.getLastRow();
  const done = {};
  if (lastLog >= 2) lg.getRange(2, 1, lastLog - 1, 1).getValues().forEach(function (r) { done[r[0]] = true; });
  const results = [];
  let orderTouched = false;
  items.forEach(function (it) {
    if (done[it.cid]) { results.push({ cid: it.cid, ok: true, duplicate: true }); return; }
    try {
      const h = HANDLERS[it.action];
      if (!h) throw new Error("Unknown action " + it.action);
      const info = h(ss, it.data || {}) || {};
      if (info.orderListChanged) orderTouched = true;
      lg.appendRow([it.cid, it.action, info.row || "", new Date().toISOString()]);
      results.push({ cid: it.cid, ok: true, row: info.row, existed: !!info.existed });
    } catch (err) {
      results.push({ cid: it.cid, ok: false, error: String(err.message || err) });
    }
  });
  SpreadsheetApp.flush();
  if (orderTouched) applyPurchasesToShopping_(ss);
  return results;
}

/* ---------- actions ---------- */
// Customers: U (21) Member?, V (22) Referred by, AA (27) Transferred out, AB (28) Cashback balance
function referrerName_(sh, name, self) {
  const n = str_(name);
  if (n === "") return "";
  const r = findRow_(sh, 6, 205, 1, n);
  if (!r) throw new Error("Referrer not found: " + n);
  const real = str_(sh.getRange(r, 1).getValue());
  if (self && real.toLowerCase() === str_(self).toLowerCase()) throw new Error("A customer cannot refer themselves");
  return real;
}
function cashbackBalance_(ss, sh, r) {
  SpreadsheetApp.flush();
  if (str_(sh.getRange(r, 21).getValue()) !== "Yes") throw new Error("Not an Inner Circle member");
  return Number(sh.getRange(r, 28).getValue()) || 0;
}

const HANDLERS = {
  addOrder: function (ss, d) {
    const cu = tab_(ss, "Customers");
    if (!findRow_(cu, 6, 205, 1, d.customer)) throw new Error("Customer not found: " + d.customer);
    const qty = Number(d.qty);
    if (!str_(d.product) || !(qty > 0)) throw new Error("Order needs a product and a quantity");
    const sh = tab_(ss, "Order Log");
    const r = firstEmpty_(sh, 6, 1005, 3, [0, 1, 2]);
    // A Customer, B Product, C Qty, D Variant, E Note, F Date
    sh.getRange(r, 1, 1, 6).setValues([[str_(d.customer), str_(d.product), qty, str_(d.variant), d.note || "", toDate_(d.date)]]);
    if (d.newProduct) {
      if (d.yen !== "" && d.yen != null) sh.getRange(r, 8).setValue(Number(d.yen));          // H Price
      if (d.surcharge !== "" && d.surcharge != null) sh.getRange(r, 9).setValue(Number(d.surcharge)); // I Surcharge
      if (d.store) sh.getRange(r, 11).setValue(d.store);                                      // K Store
    }
    return { row: r, orderListChanged: true };
  },
  updateOrder: function (ss, d) {
    const sh = tab_(ss, "Order Log");
    const r = Number(d.row);
    const cur = sh.getRange(r, 1, 1, 2).getValues()[0];
    if (str_(cur[0]) !== str_(d.expectCustomer) || str_(cur[1]) !== str_(d.expectProduct)) {
      throw new Error("That order line changed in the sheet. Sync and try again.");
    }
    if (has_(d, "qty")) {
      if (!(Number(d.qty) > 0)) throw new Error("Quantity must be above 0");
      sh.getRange(r, 3).setValue(Number(d.qty));
    }
    if (has_(d, "variant")) sh.getRange(r, 4).setValue(str_(d.variant));
    if (has_(d, "note")) sh.getRange(r, 5).setValue(d.note || "");
    return { row: r, orderListChanged: true };
  },
  cancelOrder: function (ss, d) {
    const sh = tab_(ss, "Order Log");
    const r = Number(d.row);
    const cur = sh.getRange(r, 1, 1, 2).getValues()[0];
    if (str_(cur[0]) !== str_(d.expectCustomer) || str_(cur[1]) !== str_(d.expectProduct)) {
      throw new Error("That order line changed in the sheet. Sync and try again.");
    }
    sh.getRange(r, 7).setValue("Yes");   // G Cancel
    return { row: r, orderListChanged: true };
  },
  reactivateOrder: function (ss, d) {
    const sh = tab_(ss, "Order Log");
    const r = Number(d.row);
    const cur = sh.getRange(r, 1, 1, 2).getValues()[0];
    if (str_(cur[0]) !== str_(d.expectCustomer) || str_(cur[1]) !== str_(d.expectProduct)) {
      throw new Error("That order line changed in the sheet. Sync and try again.");
    }
    if (has_(d, "qty")) {
      if (!(Number(d.qty) > 0)) throw new Error("Quantity must be above 0");
      sh.getRange(r, 3).setValue(Number(d.qty));
    }
    if (has_(d, "variant")) sh.getRange(r, 4).setValue(str_(d.variant));
    if (has_(d, "note")) sh.getRange(r, 5).setValue(d.note || "");
    sh.getRange(r, 7).clearContent();    // G Cancel
    return { row: r, orderListChanged: true };
  },
  addPayment: function (ss, d) {
    const cu = tab_(ss, "Customers");
    if (!findRow_(cu, 6, 205, 1, d.customer)) throw new Error("Customer not found: " + d.customer);
    const amount = Number(d.amount);
    if (!(amount > 0)) throw new Error("Amount must be above 0");
    const sh = tab_(ss, "Payments");
    const r = firstEmpty_(sh, 6, 505, 5, [1, 2]);
    sh.getRange(r, 1, 1, 5).setValues([[toDate_(d.date), str_(d.customer), amount, d.shipOnly ? "Yes" : "", d.note || ""]]);
    return { row: r };
  },
  updatePayment: function (ss, d) {
    const sh = tab_(ss, "Payments");
    const r = Number(d.row);
    if (!(r >= 6 && r <= 505)) throw new Error("Invalid payment row");
    const cur = sh.getRange(r, 1, 1, 5).getValues()[0];
    if (str_(cur[1]) !== str_(d.expectCustomer) || Number(cur[2]) !== Number(d.expectAmount)) throw new Error("Payment changed in the sheet. Sync and try again.");
    const amount = Number(d.amount);
    if (!(amount > 0)) throw new Error("Amount must be above 0");
    if (!findRow_(tab_(ss, "Customers"), 6, 205, 1, d.customer)) throw new Error("Customer not found: " + d.customer);
    sh.getRange(r, 1, 1, 5).setValues([[toDate_(d.date), str_(d.customer), amount, d.shipOnly ? "Yes" : "", d.note || ""]]);
    return { row: r };
  },
  deletePayment: function (ss, d) {
    const sh = tab_(ss, "Payments");
    const r = Number(d.row);
    if (!(r >= 6 && r <= 505)) throw new Error("Invalid payment row");
    const cur = sh.getRange(r, 1, 1, 5).getValues()[0];
    if (str_(cur[1]) !== str_(d.expectCustomer) || Number(cur[2]) !== Number(d.expectAmount)) throw new Error("Payment changed in the sheet. Sync and try again.");
    sh.getRange(r, 1, 1, 5).clearContent();
    return { row: r };
  },
  addCustomer: function (ss, d) {
    const sh = tab_(ss, "Customers");
    const ex = findRow_(sh, 6, 205, 1, d.name);
    if (ex) return { row: ex, existed: true };
    if (!str_(d.name)) throw new Error("Customer needs a name");
    const ic = icReady_(ss);
    const ref = ic ? referrerName_(sh, d.referredBy, d.name) : "";
    const r = firstEmpty_(sh, 6, 205, 1, [0]);
    sh.getRange(r, 1, 1, 5).setValues([[str_(d.name), d.type || "New", d.contact || "", d.address || "", d.courier || ""]]);
    if (d.shipFee !== "" && d.shipFee != null) sh.getRange(r, 12).setValue(Number(d.shipFee));
    sh.getRange(r, 16).setValue(d.status || "Shopping");
    if (d.notes) sh.getRange(r, 18).setValue(d.notes);
    if (ic) {
      if (d.member === "Yes") sh.getRange(r, 21).setValue("Yes");
      if (ref) sh.getRange(r, 22).setValue(ref);
    }
    return { row: r };
  },
  updateCustomer: function (ss, d) {
    const sh = tab_(ss, "Customers");
    const r = findRow_(sh, 6, 205, 1, d.name);
    if (!r) throw new Error("Customer not found: " + d.name);
    const ic = icReady_(ss);
    const ref = ic && has_(d, "referredBy") ? referrerName_(sh, d.referredBy, d.name) : null;
    if (has_(d, "type")) sh.getRange(r, 2).setValue(d.type);
    if (has_(d, "contact")) sh.getRange(r, 3).setValue(d.contact);
    if (has_(d, "address")) sh.getRange(r, 4).setValue(d.address);
    if (has_(d, "courier")) sh.getRange(r, 5).setValue(d.courier);
    if (has_(d, "shipFee")) sh.getRange(r, 12).setValue(d.shipFee === "" ? "" : Number(d.shipFee));
    if (has_(d, "status")) sh.getRange(r, 16).setValue(d.status);
    if (has_(d, "notes")) sh.getRange(r, 18).setValue(d.notes);
    if (ic) {
      if (has_(d, "member")) sh.getRange(r, 21).setValue(d.member === "Yes" ? "Yes" : "");
      if (ref !== null) sh.getRange(r, 22).setValue(ref);
    }
    return { row: r };
  },
  // Member uses part of the cashback balance as credit: it is logged as a normal payment
  // with the note "Referral credit", so the invoice and balance update by themselves.
  useCashbackCredit: function (ss, d) {
    if (!icReady_(ss)) throw new Error("Run setupInnerCircle first");
    const cu = tab_(ss, "Customers");
    const cr = findRow_(cu, 6, 205, 1, d.customer);
    if (!cr) throw new Error("Customer not found: " + d.customer);
    const amount = Number(d.amount);
    if (!(amount > 0)) throw new Error("Amount must be above 0");
    const bal = cashbackBalance_(ss, cu, cr);
    if (amount > bal) throw new Error("Amount is above the cashback balance (Rp " + bal + ")");
    const sh = tab_(ss, "Payments");
    const r = firstEmpty_(sh, 6, 505, 5, [1, 2]);
    sh.getRange(r, 1, 1, 5).setValues([[toDate_(d.date), str_(d.customer), amount, "", "Referral credit"]]);
    return { row: r };
  },
  // Cashback paid out by bank transfer: only reduces the balance
  payoutCashback: function (ss, d) {
    if (!icReady_(ss)) throw new Error("Run setupInnerCircle first");
    const cu = tab_(ss, "Customers");
    const cr = findRow_(cu, 6, 205, 1, d.customer);
    if (!cr) throw new Error("Customer not found: " + d.customer);
    const amount = Number(d.amount);
    if (!(amount > 0)) throw new Error("Amount must be above 0");
    const bal = cashbackBalance_(ss, cu, cr);
    if (amount > bal) throw new Error("Amount is above the cashback balance (Rp " + bal + ")");
    const cell = cu.getRange(cr, 27);
    cell.setValue((Number(cell.getValue()) || 0) + amount);
    return { row: cr };
  },
  addProduct: function (ss, d) {
    const sh = tab_(ss, "Products");
    if (!str_(d.name)) throw new Error("Product needs a name");
    if (findRow_(sh, 6, 389, 1, d.name)) throw new Error("Product already exists: " + d.name);
    const r = firstEmpty_(sh, 6, 389, 1, [0]);
    const n = function (x) { return x === "" || x == null ? "" : Number(x); };
    sh.getRange(r, 1, 1, 8).setValues([[str_(d.name), d.store || "", n(d.yen), n(d.surcharge), n(d.manualIdr), d.link || "", d.ig || "Not posted", d.notes || ""]]);
    return { row: r };
  },
  updateProduct: function (ss, d) {
    const sh = tab_(ss, "Products");
    const r = findRow_(sh, 6, 389, 1, d.name);
    if (!r) throw new Error("Product not found: " + d.name);
    const n = function (x) { return x === "" || x == null ? "" : Number(x); };
    if (has_(d, "store")) sh.getRange(r, 2).setValue(d.store);
    if (has_(d, "yen")) sh.getRange(r, 3).setValue(n(d.yen));
    if (has_(d, "surcharge")) sh.getRange(r, 4).setValue(n(d.surcharge));
    if (has_(d, "manualIdr")) sh.getRange(r, 5).setValue(n(d.manualIdr));
    if (has_(d, "link")) sh.getRange(r, 6).setValue(d.link);
    if (has_(d, "ig")) sh.getRange(r, 7).setValue(d.ig);
    if (has_(d, "notes")) sh.getRange(r, 8).setValue(d.notes);
    return { row: r };
  },
  addStore: function (ss, d) {
    const sh = tab_(ss, "Settings");
    const name = str_(d.name);
    if (!name) throw new Error("Store needs a name");
    const ex = findRow_(sh, 4, 203, 6, name);
    if (ex) return { row: ex, existed: true };
    const v = sh.getRange(4, 6, 200, 1).getValues();
    for (let i = 0; i < v.length; i++) {
      if (v[i][0] === "") { sh.getRange(4 + i, 6).setValue(name); return { row: 4 + i }; }
    }
    throw new Error("Store list is full");
  },
  upsertPurchase: function (ss, d) {
    const sl = tab_(ss, "Shopping List");
    const product = str_(d.product), variant = str_(d.variant) || "standard";
    const rows = sl.getRange(7, 1, 300, 2).getValues();
    let r = 0;
    for (let i = 0; i < rows.length; i++) {
      if (str_(rows[i][0]).toLowerCase() === product.toLowerCase() &&
          (str_(rows[i][1]) || "standard").toLowerCase() === variant.toLowerCase()) {
        r = 7 + i; break;
      }
    }
    if (!r) throw new Error("Not on the Shopping List now: " + product + " / " + variant);
    const b = d.bought === "" || d.bought == null ? "" : Number(d.bought);
    const c = d.paidYen === "" || d.paidYen == null ? "" : Number(d.paidYen);
    const sold = d.soldOut ? "Sold Out" : "";
    sl.getRange(r, 7, 1, 3).setValues([[b, c, sold]]);
    upsertMem_(ss, product, variant, b, c, sold);
    return { row: r };
  }
};

/* ---------- reading everything for the app ---------- */
function snapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ic = icReady_(ss);
  return {
    generatedAt: new Date().toISOString(),
    settings: readSettings_(ss, ic),
    home: readHome_(ss),
    products: readProducts_(ss),
    customers: readCustomers_(ss, ic),
    orders: readOrders_(ss, ic),
    shopping: readShopping_(ss),
    payments: readPayments_(ss),
    invoiceMessage: readInvoiceMessage_(ss)
  };
}
function readInvoiceMessage_(ss) {
  const sh = ss.getSheetByName("Invoice");
  if (!sh) return "";
  return sh.getRange("B42").getDisplayValue();
}
function rowsOf_(v, first, keep, mapFn) {
  const out = [];
  for (let i = 0; i < v.length; i++) {
    if (keep(v[i])) out.push(mapFn(v[i], first + i));
  }
  return out;
}
function readSettings_(ss, ic) {
  const sh = tab_(ss, "Settings");
  const c = function (a1) { return clean_(sh.getRange(a1).getValue()); };
  const list = function (a1) {
    return sh.getRange(a1).getValues().map(function (r) { return clean_(r[0]); })
      .filter(function (v) { return v !== ""; });
  };
  const s = {
    threshold: c("C4"), t1Rate: c("C5"), t1Fee: c("C6"), t2Rate: c("C7"),
    surcharges: list("C10:C17"), costRate: c("C20"),
    batch: c("C26"), shipStart: c("C27"), shipFrom: c("C28"),
    bankAccount: c("C29"),
    stores: list("F4:F203"), couriers: list("M4:M33"),
    icReady: !!ic
  };
  if (ic) {
    s.memberT1Rate = c("C35"); s.memberFee = c("C36"); s.memberT2Rate = c("C37");
    s.cashback = c("C38"); s.cashbackMin = c("C39");
  }
  return s;
}
// Dashboard cells are read through named ranges, so you can add or move rows in the
// Dashboard tab without breaking the app. Run setupDashboardNames() once (see instructions).
const DASH_NAMES = {
  DB_days: "L2", DB_rate: "K3",
  DB_collect: "L5", DB_onHold: "L6", DB_overpaid: "L7", DB_waitingItems: "L8",
  DB_readyToPack: "L9", DB_shipFees: "L10",
  DB_ordered: "B14", DB_bought: "E14", DB_toBuy: "H14", DB_progress: "K14",
  DB_billable: "B17", DB_collected: "E17", DB_toCollect: "H17", DB_credit: "K17",
  DB_customers: "B20", DB_waitingUnits: "E20", DB_profit: "H20", DB_margin: "K20",
  DB_stores: "B24:F34", DB_itemStatus: "H24:J28", DB_pipeline: "H32:J36",
  DB_payStatus: "B39:E44", DB_checks: "B47:F56"
};
function setupDashboardNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = tab_(ss, "Dashboard");
  Object.keys(DASH_NAMES).forEach(function (n) {
    if (!ss.getRangeByName(n)) ss.setNamedRange(n, dash.getRange(DASH_NAMES[n]));
  });
}
function dashBlock_(ss, dash, name) {
  const r = ss.getRangeByName(name);
  return (r || dash.getRange(DASH_NAMES[name])).getValues();
}
function readHome_(ss) {
  const dash = tab_(ss, "Dashboard");
  const one = function (n) { return clean_(dashBlock_(ss, dash, n)[0][0]); };
  const blk = function (n) {
    return dashBlock_(ss, dash, n).map(function (r) { return r.map(clean_); });
  };
  return {
    daysText: one("DB_days"), rate: one("DB_rate"),
    actions: {
      collect: one("DB_collect"), onHold: one("DB_onHold"), overpaid: one("DB_overpaid"),
      waitingItems: one("DB_waitingItems"), readyToPack: one("DB_readyToPack"),
      shippingToCollect: one("DB_shipFees")
    },
    units: { ordered: one("DB_ordered"), bought: one("DB_bought"), toBuy: one("DB_toBuy"), progress: one("DB_progress") },
    money: { billable: one("DB_billable"), collected: one("DB_collected"), toCollect: one("DB_toCollect"), credit: one("DB_credit") },
    people: { customers: one("DB_customers"), waitingUnits: one("DB_waitingUnits"), profit: one("DB_profit"), margin: one("DB_margin") },
    stores: blk("DB_stores").map(function (r) {
      return { store: r[0], toBuy: r[2], bought: r[3], billable: r[4] };
    }).filter(function (x) { return x.store !== ""; }),
    itemStatus: blk("DB_itemStatus").map(function (r) { return { label: r[0], units: r[2] }; }),
    pipeline: blk("DB_pipeline").map(function (r) { return { label: r[0], customers: r[2] }; }),
    payStatus: blk("DB_payStatus").map(function (r) { return { label: r[0], customers: r[2], balance: r[3] }; }),
    checks: blk("DB_checks").map(function (r) { return { label: r[0], value: r[4] }; })
  };
}
function readProducts_(ss) {
  const v = tab_(ss, "Products").getRange(6, 1, 384, 15).getValues();
  return rowsOf_(v, 6, function (r) { return r[0] !== ""; }, function (r, row) {
    return {
      row: row, name: r[0], store: r[1], yen: r[2], surcharge: r[3], manualIdr: r[4], link: r[5],
      ig: r[6], notes: r[7], converted: r[8], selling: clean_(r[9]), label: r[10],
      cost: r[11], profit: r[12], margin: r[13], ordered: r[14]
    };
  });
}
function readCustomers_(ss, ic) {
  const v = tab_(ss, "Customers").getRange(6, 1, 200, ic ? 28 : 18).getValues();
  return rowsOf_(v, 6, function (r) { return r[0] !== ""; }, function (r, row) {
    const o = {
      row: row, name: r[0], type: r[1], contact: r[2], address: r[3], courier: r[4],
      units: r[5], bought: r[6], waiting: r[7], toBuy: r[8],
      billable: r[9], itemsPaid: r[10], shipFee: r[11], shipPaid: r[12],
      balance: r[13], payStatus: r[14], status: r[15], next: r[16], notes: r[17]
    };
    if (ic) {
      o.member = r[20] === "Yes" ? "Yes" : "";
      o.referredBy = r[21]; o.referralStatus = r[22];
      o.cashbackEarned = r[23]; o.creditUsed = r[25]; o.cashbackOut = r[26]; o.cashbackBalance = r[27];
    }
    return o;
  });
}
// Order Log now has 32 columns to read (one new Note column at E); member flag sits in AN (col 40).
function readOrders_(ss, ic) {
  const v = tab_(ss, "Order Log").getRange(6, 1, 1000, ic ? 40 : 32).getValues();
  return rowsOf_(v, 6, function (r) { return r[0] !== "" || r[1] !== ""; }, function (r, row) {
    return {
      row: row, customer: r[0], product: r[1], qty: r[2], variant: str_(r[3]), note: r[4], date: clean_(r[5]), cancel: r[6],
      unit: r[14], gross: r[15], bought: r[25], soldOut: r[26], toBuy: r[27],
      waiting: r[28], billable: r[30], status: r[31],
      member: ic ? (r[39] === "Yes" ? "Yes" : "") : ""
    };
  });
}
function readShopping_(ss) {
  const v = tab_(ss, "Shopping List").getRange(7, 1, 300, 13).getValues();
  return rowsOf_(v, 7, function (r) { return r[0] !== ""; }, function (r, row) {
    return {
      row: row, product: r[0], variant: str_(r[1]) || "standard", status: r[2], store: r[3],
      need: r[4], waiting: r[5], bought: r[6], paidYen: r[7], soldOut: r[8],
      toBuy: r[9], spare: r[10], given: r[11], soldOutUnits: r[12]
    };
  });
}
function readPayments_(ss) {
  const v = tab_(ss, "Payments").getRange(6, 1, 500, 5).getValues();
  return rowsOf_(v, 6, function (r) { return r[1] !== ""; }, function (r, row) {
    return { row: row, date: clean_(r[0]), customer: r[1], amount: r[2], shipOnly: r[3], note: r[4] };
  });
}
