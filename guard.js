/*
 * IRIS Recipient Guard - shared rules engine
 * ------------------------------------------------------------
 * Loaded by:
 *   - commands.html  (event runtime for Outlook on the web / new Outlook / Mac)
 *   - directly as the JavaScript runtime file for classic Outlook on Windows
 *   - taskpane.html  (status pane + company picker)
 *
 * Everything that decides whether an email may be sent lives in this one file,
 * so the send check and the status pane always agree.
 *
 * Plain ES2017, no DOM, no atob/TextDecoder: the classic Outlook JavaScript
 * runtime does not provide them.
 */

/* ============================== CONFIGURATION ============================== */

var RG_CONFIG = {
  VERSION: "1.0.0",

  // HTTPS address of the client list (CSV). Replace with the internal server
  // address IT gives you. The server must send CORS header:
  //   Access-Control-Allow-Origin: https://psrikanth72.github.io
  // While this points at the SAMPLE file, the status pane shows a red warning.
  REPOSITORY_URL: "https://psrikanth72.github.io/iris-recipient-guard/sample-clients.csv",
  SAMPLE_REPOSITORY_URL: "https://psrikanth72.github.io/iris-recipient-guard/sample-clients.csv",

  // Domains that are always allowed (exact domain match, case-insensitive).
  INTERNAL_DOMAINS: ["iriskpo.in", "iris.co.uk"],

  // How long the in-memory copy of the client list is reused by the
  // as-you-type checks. The Send check ALWAYS downloads a fresh copy.
  CACHE_SECONDS: 60,

  // Network timeout for downloading the client list.
  FETCH_TIMEOUT_MS: 10000,

  // Overall time budget for the Send check before it gives up and blocks.
  SEND_CHECK_TIMEOUT_MS: 90000,

  // Image attachments embedded in the body (e.g. signature logos) are skipped.
  SKIP_INLINE_IMAGES: true
};

/* ================================ UTILITIES ================================ */

function rgLower(s) { return String(s == null ? "" : s).trim().toLowerCase(); }

// Letters and digits only - used to compare company names with subjects/files.
function rgNorm(s) { return rgLower(s).replace(/[^a-z0-9]/g, ""); }

// Collapse whitespace, lower-case - used for subject prefix matching.
function rgSpaceNorm(s) { return rgLower(s).replace(/\s+/g, " "); }

function rgDomain(addr) {
  var a = rgLower(addr);
  var i = a.lastIndexOf("@");
  return i < 0 ? "" : a.slice(i + 1);
}

function rgIsInternal(addr) {
  var d = rgDomain(addr);
  for (var i = 0; i < RG_CONFIG.INTERNAL_DOMAINS.length; i++) {
    if (d === RG_CONFIG.INTERNAL_DOMAINS[i]) return true;
  }
  return false;
}

function rgIsEmail(s) { return /^[^@\s<>]+@[^@\s<>]+\.[^@\s<>]+$/.test(String(s || "").trim()); }

function rgUnique(list) {
  var seen = {}, out = [];
  for (var i = 0; i < list.length; i++) {
    var k = rgLower(list[i]);
    if (!seen[k]) { seen[k] = 1; out.push(list[i]); }
  }
  return out;
}

// Promise wrapper for Office "...Async(callback)" APIs.
function rgAsync(fn) {
  return new Promise(function (resolve, reject) {
    try {
      fn(function (res) {
        if (res && res.status === "failed") {
          reject(new Error((res.error && res.error.message) || "Outlook call failed"));
        } else {
          resolve(res ? res.value : undefined);
        }
      });
    } catch (e) { reject(e); }
  });
}

function rgWithTimeout(promise, ms, what) {
  return new Promise(function (resolve, reject) {
    var t = setTimeout(function () { reject(new Error(what + " timed out")); }, ms);
    promise.then(function (v) { clearTimeout(t); resolve(v); },
                 function (e) { clearTimeout(t); reject(e); });
  });
}

// Base64 -> Uint8Array without atob (not available in classic Outlook runtime).
var RG_B64 = (function () {
  var map = {}, chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  for (var i = 0; i < chars.length; i++) map[chars.charAt(i)] = i;
  map["-"] = 62; map["_"] = 63;
  return map;
})();

function rgBase64ToBytes(b64) {
  var s = String(b64 || "").replace(/[^A-Za-z0-9+/\-_]/g, "");
  var len = Math.floor(s.length * 3 / 4);
  var out = new Uint8Array(len);
  var o = 0, buf = 0, bits = 0;
  for (var i = 0; i < s.length; i++) {
    buf = (buf << 6) | RG_B64[s.charAt(i)];
    bits += 6;
    if (bits >= 8) { bits -= 8; out[o++] = (buf >> bits) & 0xff; }
  }
  return o === len ? out : out.subarray(0, o);
}

/* ============================ CLIENT REPOSITORY ============================ */

// Parses CSV text into rows of fields (handles quotes, commas, CRLF).
function rgParseCsv(text) {
  var rows = [], row = [], field = "", inQ = false;
  text = String(text || "").replace(/^﻿/, "");
  for (var i = 0; i < text.length; i++) {
    var c = text.charAt(i);
    if (inQ) {
      if (c === '"') {
        if (text.charAt(i + 1) === '"') { field += '"'; i++; } else { inQ = false; }
      } else { field += c; }
    } else if (c === '"') { inQ = true; }
    else if (c === ",") { row.push(field); field = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text.charAt(i + 1) === "\n") i++;
      row.push(field); rows.push(row); row = []; field = "";
    } else { field += c; }
  }
  if (field !== "" || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/*
 * Builds the repository from CSV text. The header row is found automatically
 * (first row with a "company" column and an "email" column), so the existing
 * Excel layout (Sl. No. | Company Name | Email ID | Status) can be saved as CSV
 * unchanged. Status is optional: blank / Active = active, anything starting
 * with "inactive", "terminated", "deactivated", "closed", "no" = inactive.
 */
function rgBuildRepository(csvText) {
  var rows = rgParseCsv(csvText);
  var hdr = -1, cCol = -1, eCol = -1, sCol = -1;
  for (var r = 0; r < rows.length && r < 30; r++) {
    var cc = -1, ec = -1, sc = -1;
    for (var k = 0; k < rows[r].length; k++) {
      var h = rgLower(rows[r][k]);
      if (cc < 0 && h.indexOf("company") >= 0) cc = k;
      else if (ec < 0 && (h.indexOf("email") >= 0 || h.indexOf("e-mail") >= 0)) ec = k;
      else if (sc < 0 && (h === "status" || h.indexOf("status") >= 0 || h === "active")) sc = k;
    }
    if (cc >= 0 && ec >= 0) { hdr = r; cCol = cc; eCol = ec; sCol = sc; break; }
  }
  if (hdr < 0) throw new Error("Client list has no 'Company' and 'Email' header row");

  var repo = {
    companies: {},        // normName -> { name, emails: [lower], inactiveEmails: [lower] }
    byEmail: {},          // lower email -> normName   (active only)
    inactiveByEmail: {},  // lower email -> normName
    conflicts: [],        // same address under two companies
    companyCount: 0, contactCount: 0, inactiveCount: 0,
    loadedAt: new Date()
  };
  var lastCompany = "";
  for (var i = hdr + 1; i < rows.length; i++) {
    var row = rows[i];
    var name = String(row[cCol] || "").trim();
    if (!name) name = lastCompany;            // tolerate merged-cell style blanks
    var emailCell = String(row[eCol] || "").trim();
    if (!name || !emailCell) continue;
    lastCompany = name;
    var statusTxt = sCol >= 0 ? rgLower(row[sCol]) : "";
    var inactive = /^(inactive|terminated|deactivated|disabled|closed|left|removed|no|n|false|0)\b/.test(statusTxt);
    var key = rgNorm(name);
    if (!key) continue;
    var comp = repo.companies[key];
    if (!comp) {
      comp = repo.companies[key] = { name: name, key: key, emails: [], inactiveEmails: [] };
    }
    var parts = emailCell.split(/[;,\s]+/);
    for (var p = 0; p < parts.length; p++) {
      var a = rgLower(parts[p]);
      if (!rgIsEmail(a)) continue;
      if (inactive) {
        if (comp.inactiveEmails.indexOf(a) < 0) comp.inactiveEmails.push(a);
        repo.inactiveByEmail[a] = key;
        continue;
      }
      if (repo.byEmail[a] && repo.byEmail[a] !== key) {
        repo.conflicts.push(a);
        continue;
      }
      if (comp.emails.indexOf(a) < 0) comp.emails.push(a);
      repo.byEmail[a] = key;
    }
  }
  // An address listed under two companies is ambiguous: treat as unknown (fail closed).
  for (var c = 0; c < repo.conflicts.length; c++) {
    var bad = repo.conflicts[c];
    delete repo.byEmail[bad];
    for (var ck in repo.companies) {
      var arr = repo.companies[ck].emails, ix = arr.indexOf(bad);
      if (ix >= 0) arr.splice(ix, 1);
    }
  }
  // A company whose contacts are all inactive is itself inactive.
  for (var key2 in repo.companies) {
    var co = repo.companies[key2];
    co.active = co.emails.length > 0;
    if (co.active) { repo.companyCount++; repo.contactCount += co.emails.length; }
    repo.inactiveCount += co.inactiveEmails.length;
  }
  // Longest names first, so "Andrew Cross Ltd" wins over "Andrew Cross".
  repo.sortedKeys = Object.keys(repo.companies).filter(function (k) {
    return repo.companies[k].active;
  }).sort(function (a, b) {
    return b.length - a.length;
  });
  return repo;
}

var rgRepoCache = { repo: null, at: 0, error: null };

function rgFetchText(url) {
  var sep = url.indexOf("?") >= 0 ? "&" : "?";
  var full = url + sep + "rg=" + Date.now();          // defeat caches: always the live file
  if (typeof fetch === "function") {
    return rgWithTimeout(
      fetch(full, { cache: "no-store", credentials: "omit" }).then(function (r) {
        if (!r.ok) throw new Error("Client list server answered HTTP " + r.status);
        return r.text();
      }), RG_CONFIG.FETCH_TIMEOUT_MS, "Client list download");
  }
  return rgWithTimeout(new Promise(function (resolve, reject) {
    var x = new XMLHttpRequest();
    x.open("GET", full, true);
    x.onload = function () {
      if (x.status >= 200 && x.status < 300) resolve(x.responseText);
      else reject(new Error("Client list server answered HTTP " + x.status));
    };
    x.onerror = function () { reject(new Error("Client list could not be reached")); };
    x.send();
  }), RG_CONFIG.FETCH_TIMEOUT_MS, "Client list download");
}

// fresh=true: always download (used by the Send check).
function rgLoadRepository(fresh) {
  var now = Date.now();
  if (!fresh && rgRepoCache.repo && now - rgRepoCache.at < RG_CONFIG.CACHE_SECONDS * 1000) {
    return Promise.resolve(rgRepoCache.repo);
  }
  return rgFetchText(RG_CONFIG.REPOSITORY_URL).then(function (text) {
    var repo = rgBuildRepository(text);
    if (repo.companyCount === 0) throw new Error("Client list is empty");
    rgRepoCache = { repo: repo, at: Date.now(), error: null };
    return repo;
  }, function (err) {
    rgRepoCache.error = err;
    throw err;
  });
}

/* =========================== COMPANY FROM SUBJECT ========================== */

// Removes any number of leading RE:/FW:/FWD:/AW:/TR:/SV: prefixes (and [EXT] tags).
function rgStripSubjectPrefixes(subject) {
  var s = String(subject || "").trim();
  var re = /^(\s*(re|fw|fwd|aw|wg|tr|sv|vs|antw|r)\s*(\[\d+\])?\s*[:：]\s*|\s*\[(ext|external)\]\s*)/i;
  var guard = 0;
  while (re.test(s) && guard++ < 20) s = s.replace(re, "");
  return s.trim();
}

// If the subject starts with the company name (ignoring case, spaces and
// punctuation) returns the index just after the name, otherwise -1.
// The name must end on a word boundary: "Andrew Cross" does not match "Andrew Crossley".
function rgPrefixEnd(subject, nameNorm) {
  var acc = "", s = String(subject || "");
  for (var i = 0; i < s.length; i++) {
    var ch = s.charAt(i).toLowerCase();
    if (!/[a-z0-9]/.test(ch)) continue;
    acc += ch;
    if (acc.length === nameNorm.length) {
      if (acc !== nameNorm) return -1;
      return /[a-z0-9]/i.test(s.charAt(i + 1)) ? -1 : i + 1;
    }
    if (nameNorm.indexOf(acc) !== 0) return -1;
  }
  return -1;
}

// Returns the company key whose name starts the subject, or null.
function rgCompanyFromSubject(repo, subject) {
  var s = rgStripSubjectPrefixes(subject);
  if (!s) return null;
  for (var i = 0; i < repo.sortedKeys.length; i++) {
    if (rgPrefixEnd(s, repo.sortedKeys[i]) >= 0) return repo.sortedKeys[i];
  }
  return null;
}

// Puts the company name at the start of the subject, replacing any existing one.
function rgSubjectWithCompany(repo, subject, key) {
  var rest = String(subject || "");
  var cur = rgCompanyFromSubject(repo, rest);
  if (cur) rest = rest.slice(rgPrefixEnd(rest, cur));
  rest = rest.replace(/^\s*[-–:|]?\s*/, "");
  return repo.companies[key].name + " - " + rest;
}

/* ========================== ATTACHMENT INSPECTION ========================== */

function rgExt(name) {
  var n = rgLower(name), i = n.lastIndexOf(".");
  return i < 0 ? "" : n.slice(i + 1);
}

function rgU16(b, o) { return b[o] | (b[o + 1] << 8); }
function rgU32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

function rgStartsWith(b, sig) {
  if (b.length < sig.length) return false;
  for (var i = 0; i < sig.length; i++) if (b[i] !== sig[i]) return false;
  return true;
}

function rgIndexOf(b, sig, from, to) {
  var end = Math.min(to == null ? b.length : to, b.length) - sig.length;
  outer: for (var i = from || 0; i <= end; i++) {
    for (var j = 0; j < sig.length; j++) if (b[i + j] !== sig[j]) continue outer;
    return i;
  }
  return -1;
}

function rgAscii(str) { var a = []; for (var i = 0; i < str.length; i++) a.push(str.charCodeAt(i)); return a; }
function rgUtf16(str) { var a = []; for (var i = 0; i < str.length; i++) { a.push(str.charCodeAt(i)); a.push(0); } return a; }

var RG_SIG_ZIP = [0x50, 0x4b, 0x03, 0x04];
var RG_SIG_ZIP_EMPTY = [0x50, 0x4b, 0x05, 0x06];
var RG_SIG_CFB = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
var RG_SIG_7Z = [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c];
var RG_SIG_PDF = rgAscii("%PDF");

// ZIP: every file entry in the central directory must have the "encrypted" flag.
function rgZipEncrypted(b) {
  var eocd = -1;
  for (var i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
    if (b[i] === 0x50 && b[i + 1] === 0x4b && b[i + 2] === 0x05 && b[i + 3] === 0x06) { eocd = i; break; }
  }
  if (eocd < 0) return { ok: false, why: "the ZIP file looks damaged" };
  var count = rgU16(b, eocd + 10), cdOff = rgU32(b, eocd + 16);
  if (count === 0xffff || cdOff === 0xffffffff) {
    // ZIP64: fall back to scanning local headers
    return rgZipScanLocal(b);
  }
  var p = cdOff, files = 0;
  for (var n = 0; n < count; n++) {
    if (rgU32(b, p) !== 0x02014b50) return rgZipScanLocal(b);
    var flags = rgU16(b, p + 8);
    var usize = rgU32(b, p + 24);
    var nlen = rgU16(b, p + 28), xlen = rgU16(b, p + 30), clen = rgU16(b, p + 32);
    var isDir = b[p + 46 + nlen - 1] === 0x2f; // name ends with '/'
    if (!isDir || usize > 0) {
      files++;
      if (!(flags & 1)) return { ok: false, why: "the ZIP file is not password protected (or some files inside it are not)" };
    }
    p += 46 + nlen + xlen + clen;
  }
  if (files === 0) return { ok: false, why: "the ZIP file is empty" };
  return { ok: true };
}

function rgZipScanLocal(b) {
  var p = 0, files = 0;
  while ((p = rgIndexOf(b, RG_SIG_ZIP, p)) >= 0) {
    var flags = rgU16(b, p + 6);
    files++;
    if (!(flags & 1)) return { ok: false, why: "the ZIP file is not password protected (or some files inside it are not)" };
    p += 4;
  }
  return files ? { ok: true } : { ok: false, why: "the ZIP file looks damaged" };
}

// 7-Zip: the header must declare the AES-256 coder (06 F1 07 01).
function rg7zEncrypted(b) {
  if (b.length < 32) return { ok: false, why: "the 7z file looks damaged" };
  var off = rgU32(b, 12) + rgU32(b, 16) * 4294967296;
  var size = rgU32(b, 20) + rgU32(b, 24) * 4294967296;
  var start = 32 + off, end = start + size;
  if (start >= b.length || end > b.length) return { ok: false, why: "the 7z file looks damaged" };
  if (rgIndexOf(b, [0x06, 0xf1, 0x07, 0x01], start, end) >= 0) return { ok: true };
  return { ok: false, why: "the 7z file could not be confirmed as password protected - in 7-Zip set a password AND tick \"Encrypt file names\"" };
}

function rgPdfEncrypted(b) {
  // Look for the /Encrypt dictionary entry (trailer or xref stream).
  if (rgIndexOf(b, rgAscii("/Encrypt"), 0) >= 0) return { ok: true };
  return { ok: false, why: "the PDF is not password protected" };
}

/* --- Minimal Compound File (CFB / OLE2) reader for legacy Office files --- */
function rgCfbOpen(b) {
  if (!rgStartsWith(b, RG_SIG_CFB)) return null;
  var secShift = rgU16(b, 30), miniShift = rgU16(b, 32);
  var secSize = 1 << secShift, miniSize = 1 << miniShift;
  var numFat = rgU32(b, 44), dirStart = rgU32(b, 48), miniCutoff = rgU32(b, 56);
  var miniFatStart = rgU32(b, 60), difStart = rgU32(b, 68), numDif = rgU32(b, 72);
  function secOff(s) { return (s + 1) * secSize; }
  var fatSecs = [];
  for (var i = 0; i < 109 && fatSecs.length < numFat; i++) fatSecs.push(rgU32(b, 76 + i * 4));
  var d = difStart, guard = 0;
  while (fatSecs.length < numFat && d < 0xfffffffa && guard++ < numDif + 1) {
    var o = secOff(d), per = secSize / 4 - 1;
    for (var j = 0; j < per && fatSecs.length < numFat; j++) fatSecs.push(rgU32(b, o + j * 4));
    d = rgU32(b, o + per * 4);
  }
  var fat = [];
  for (var f = 0; f < fatSecs.length; f++) {
    var fo = secOff(fatSecs[f]);
    for (var k = 0; k < secSize / 4; k++) fat.push(rgU32(b, fo + k * 4));
  }
  function chain(start) {
    var out = [], s = start, g = 0;
    while (s < 0xfffffffa && g++ < 1000000) { out.push(s); s = fat[s]; }
    return out;
  }
  function readChain(start, maxLen) {
    var secs = chain(start), out = new Uint8Array(secs.length * secSize);
    for (var i = 0; i < secs.length; i++) out.set(b.subarray(secOff(secs[i]), secOff(secs[i]) + secSize), i * secSize);
    return maxLen != null ? out.subarray(0, maxLen) : out;
  }
  var dir = readChain(dirStart);
  var entries = [];
  for (var e = 0; e + 128 <= dir.length; e += 128) {
    var nlen = rgU16(dir, e + 64);
    var name = "";
    for (var c = 0; c < nlen - 2 && c < 64; c += 2) name += String.fromCharCode(rgU16(dir, e + c));
    entries.push({ name: name, type: dir[e + 66], start: rgU32(dir, e + 116), size: rgU32(dir, e + 120) });
  }
  var root = entries[0];
  var miniStream = root ? readChain(root.start, root.size) : new Uint8Array(0);
  var miniFat = [];
  if (miniFatStart < 0xfffffffa) {
    var mf = readChain(miniFatStart);
    for (var m = 0; m + 4 <= mf.length; m += 4) miniFat.push(rgU32(mf, m));
  }
  return {
    entries: entries,
    has: function (n) { for (var i = 0; i < entries.length; i++) if (entries[i].name === n) return true; return false; },
    stream: function (n) {
      for (var i = 0; i < entries.length; i++) {
        var en = entries[i];
        if (en.name !== n || en.type !== 2) continue;
        if (en.size < miniCutoff) {
          var out = new Uint8Array(en.size), s = en.start, pos = 0, g = 0;
          while (s < 0xfffffffa && pos < en.size && g++ < 1000000) {
            var len = Math.min(miniSize, en.size - pos);
            out.set(miniStream.subarray(s * miniSize, s * miniSize + len), pos);
            pos += len; s = miniFat[s];
          }
          return out;
        }
        return readChain(en.start, en.size);
      }
      return null;
    }
  };
}

function rgOfficeEncrypted(b, ext) {
  var ooxml = /^(docx|docm|dotx|xlsx|xlsm|xlsb|xltx|pptx|pptm|potx|ppsx)$/.test(ext);
  if (rgStartsWith(b, RG_SIG_ZIP) || rgStartsWith(b, RG_SIG_ZIP_EMPTY)) {
    // A plain (unencrypted) Office file is a ZIP package.
    return { ok: false, why: "the Office file is not password protected (use File > Info > Protect > Encrypt with Password)" };
  }
  var cfb;
  try { cfb = rgCfbOpen(b); } catch (e) { cfb = null; }
  if (!cfb) return { ok: false, why: "the file could not be read to confirm password protection" };
  if (cfb.has("EncryptedPackage") || cfb.has("EncryptionInfo")) return { ok: true };
  if (ooxml) return { ok: false, why: "the Office file is not password protected" };

  try {
    if (ext === "doc" || ext === "dot") {
      var wd = cfb.stream("WordDocument");
      if (wd && wd.length > 12 && (rgU16(wd, 0x0a) & 0x0100)) return { ok: true };
      return { ok: false, why: "the Word file is not password protected" };
    }
    if (ext === "xls" || ext === "xlt") {
      var wb = cfb.stream("Workbook") || cfb.stream("Book");
      if (wb) {
        var p = 0, n = 0;
        while (p + 4 <= wb.length && n++ < 64) {
          var type = rgU16(wb, p), len = rgU16(wb, p + 2);
          if (type === 0x002f) return { ok: true };       // FILEPASS
          if (n > 1 && type === 0x0809) break;            // next BOF: past globals header
          p += 4 + len;
        }
      }
      return { ok: false, why: "the Excel file is not password protected" };
    }
    if (ext === "ppt" || ext === "pps" || ext === "pot") {
      var cu = cfb.stream("Current User");
      if (cu && cu.length >= 16 && rgU32(cu, 12) === 0xf3d1c4df) return { ok: true };
      return { ok: false, why: "the PowerPoint file is not password protected" };
    }
  } catch (e2) { /* fall through */ }
  return { ok: false, why: "the file could not be read to confirm password protection" };
}

var RG_SUPPORTED_EXT = {
  zip: 1, "7z": 1, pdf: 1,
  doc: 1, dot: 1, docx: 1, docm: 1, dotx: 1,
  xls: 1, xlt: 1, xlsx: 1, xlsm: 1, xlsb: 1, xltx: 1,
  ppt: 1, pps: 1, pot: 1, pptx: 1, pptm: 1, potx: 1, ppsx: 1
};

// Returns { ok, why } for one attachment's bytes.
function rgCheckProtection(name, bytes) {
  var ext = rgExt(name);
  if (!RG_SUPPORTED_EXT[ext]) {
    return { ok: false, why: "." + (ext || "?") + " files cannot be password protected - put the file in a password-protected ZIP" };
  }
  if (ext === "zip") return rgZipEncrypted(bytes);
  if (ext === "7z") return rgStartsWith(bytes, RG_SIG_7Z) ? rg7zEncrypted(bytes) : { ok: false, why: "the 7z file looks damaged" };
  if (ext === "pdf") return rgStartsWith(bytes, RG_SIG_PDF) || rgIndexOf(bytes, RG_SIG_PDF, 0, 1024) >= 0
    ? rgPdfEncrypted(bytes) : { ok: false, why: "the PDF looks damaged" };
  return rgOfficeEncrypted(bytes, ext);
}

/* =============================== OUTLOOK I/O =============================== */

function rgItem() { return Office.context.mailbox.item; }

function rgGetRecipients(field) {
  return rgAsync(function (cb) { rgItem()[field].getAsync(cb); }).then(function (v) { return v || []; });
}

function rgGetAllRecipients() {
  return Promise.all([rgGetRecipients("to"), rgGetRecipients("cc"), rgGetRecipients("bcc")])
    .then(function (r) { return { to: r[0], cc: r[1], bcc: r[2] }; });
}

function rgSetRecipients(field, list) {
  return rgAsync(function (cb) { rgItem()[field].setAsync(list, cb); });
}

function rgAddRecipients(field, list) {
  return rgAsync(function (cb) { rgItem()[field].addAsync(list, cb); });
}

function rgGetSubject() { return rgAsync(function (cb) { rgItem().subject.getAsync(cb); }).then(function (s) { return s || ""; }); }
function rgSetSubject(s) { return rgAsync(function (cb) { rgItem().subject.setAsync(s, cb); }); }

function rgGetComposeType() {
  var item = rgItem();
  if (!item.getComposeTypeAsync) return Promise.resolve("unknown");
  return rgAsync(function (cb) { item.getComposeTypeAsync(cb); })
    .then(function (v) { return rgLower(v && v.composeType ? v.composeType : v) || "unknown"; })
    .catch(function () { return "unknown"; });
}

// Session data is shared between the event runtime and the task pane for this draft.
function rgSessionGet(key) {
  var item = rgItem();
  if (!item.sessionData) return Promise.resolve(null);
  return rgAsync(function (cb) { item.sessionData.getAsync(key, cb); }).catch(function () { return null; });
}
function rgSessionSet(key, value) {
  var item = rgItem();
  if (!item.sessionData) return Promise.resolve();
  return rgAsync(function (cb) { item.sessionData.setAsync(key, String(value == null ? "" : value), cb); }).catch(function () {});
}

function rgNotify(key, text, opts) {
  var item = rgItem();
  if (!item.notificationMessages) return Promise.resolve();
  var msg = {
    type: (opts && opts.type) || Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: String(text).slice(0, 150)
  };
  if (msg.type === Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage ||
      msg.type === Office.MailboxEnums.ItemNotificationMessageType.InsightMessage) {
    msg.icon = "Icon.16x16";
  }
  if (msg.type === Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage) msg.persistent = false;
  if (opts && opts.actions) msg.actions = opts.actions;
  return rgAsync(function (cb) { item.notificationMessages.replaceAsync(key, msg, cb); }).catch(function () {
    // Insight messages with actions are not supported everywhere: retry plain.
    if (opts && opts.actions) {
      return rgNotify(key, text, { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage });
    }
  });
}

function rgAttachments() {
  var item = rgItem();
  if (!item.getAttachmentsAsync) return Promise.resolve(item.attachments || []);
  return rgAsync(function (cb) { item.getAttachmentsAsync(cb); }).then(function (v) { return v || []; });
}

function rgAttachmentBytes(id) {
  return rgAsync(function (cb) { rgItem().getAttachmentContentAsync(id, cb); }).then(function (c) {
    if (!c) throw new Error("no content");
    var fmt = rgLower(c.format);
    if (fmt === "base64") return { kind: "file", bytes: rgBase64ToBytes(c.content) };
    if (fmt === "url") return { kind: "cloud" };
    if (fmt === "eml" || fmt === "icalendar") return { kind: "item" };
    return { kind: "unknown" };
  });
}

/* ============================== RULES ENGINE =============================== */

function rgModeOf(composeType) {
  if (composeType === "forward") return "forward";
  if (composeType === "reply" || composeType === "replyall") return "reply";
  if (composeType === "newmail") return "new";
  return "unknown";
}

// Classifies every external recipient against the repository.
function rgClassify(repo, recips) {
  var all = [].concat(recips.to, recips.cc, recips.bcc);
  var out = { internal: [], external: [], unknown: [], inactive: [], companies: {} };
  for (var i = 0; i < all.length; i++) {
    var r = all[i];
    var addr = rgLower(r.emailAddress);
    var isDL = r.recipientType && rgLower(r.recipientType) === "distributionlist";
    if (addr && rgIsInternal(addr)) { out.internal.push(addr); continue; }
    if (!addr || !rgIsEmail(addr)) { out.unknown.push(r.displayName || addr || "(unresolved name)"); out.external.push(addr); continue; }
    out.external.push(addr);
    if (isDL) { out.unknown.push(addr); continue; }
    var key = repo ? repo.byEmail[addr] : null;
    if (key) { (out.companies[key] = out.companies[key] || []).push(addr); }
    else if (repo && repo.inactiveByEmail[addr]) { out.inactive.push(addr); }
    else { out.unknown.push(addr); }
  }
  out.companyKeys = Object.keys(out.companies);
  return out;
}

/*
 * Works out which client the email belongs to.
 *  - New email: the company at the start of the subject.
 *  - Reply: the company at the start of the subject (RE:/FW: ignored);
 *           otherwise the company recorded for the thread when the reply opened.
 */
function rgDetermineCompany(repo, mode, subject, threadCompany) {
  var fromSubject = rgCompanyFromSubject(repo, subject);
  if (fromSubject) return { key: fromSubject, source: "subject" };
  if (mode === "reply" && threadCompany && repo.companies[threadCompany] && repo.companies[threadCompany].active) {
    return { key: threadCompany, source: "thread" };
  }
  return { key: null, source: null };
}

/*
 * The full Send-time validation. Returns { allow, problems[], company, mode }.
 * Pure decision logic - called by the Send handler and by the status pane.
 */
function rgValidate(ctx) {
  var repo = ctx.repo, mode = ctx.mode, problems = [];
  var cls = rgClassify(repo, ctx.recipients);
  var company = rgDetermineCompany(repo, mode, ctx.subject, ctx.threadCompany);
  var hasExternal = cls.external.length > 0;
  var res = { allow: true, problems: problems, mode: mode, company: company.key,
              companyName: company.key ? repo.companies[company.key].name : null,
              externalCount: cls.external.length };

  if (!hasExternal) return res;                         // internal-only: always allowed

  if (mode === "forward") {
    problems.push("Forwarded emails can only go to @iriskpo.in or @iris.co.uk. Remove: " + cls.external.join(", "));
  }
  if (mode === "unknown") {
    problems.push("Outlook did not say whether this is a new email, reply or forward, so external recipients can't be checked.");
  }
  if (cls.inactive.length) {
    problems.push("These contacts are no longer authorised (inactive in the client list): " + cls.inactive.join(", "));
  }
  if (cls.unknown.length) {
    problems.push("Not in the client list: " + cls.unknown.join(", "));
  }
  if (cls.companyKeys.length > 1) {
    problems.push("Recipients belong to different clients (" +
      cls.companyKeys.map(function (k) { return repo.companies[k].name; }).join(" / ") +
      "). An email can only go to one client.");
  }
  if (mode === "new" && !company.key) {
    problems.push("Start the subject with the client's company name, e.g. \"" +
      (repo.sortedKeys.length ? repo.companies[repo.sortedKeys[repo.sortedKeys.length - 1]].name : "Company Ltd") + " - March payroll\".");
  }
  if (mode === "reply" && !company.key && cls.companyKeys.length === 1 && !cls.unknown.length) {
    // Reply with no company in subject: the recipients' own company is the thread company.
    company = { key: cls.companyKeys[0], source: "recipients" };
    res.company = company.key; res.companyName = repo.companies[company.key].name;
  }
  if (company.key) {
    for (var i = 0; i < cls.companyKeys.length; i++) {
      var k = cls.companyKeys[i];
      if (k !== company.key) {
        problems.push("This email is for " + repo.companies[company.key].name + ", but " +
          cls.companies[k].join(", ") + " belong to " + repo.companies[k].name + ".");
      }
    }
  } else if (mode === "reply" && !problems.length) {
    problems.push("Could not work out which client this reply is for.");
  }

  // Attachment results (pre-computed by the caller because they are async).
  var att = ctx.attachments || [];
  var compName = company.key ? repo.companies[company.key].name : null;
  for (var a = 0; a < att.length; a++) {
    var at = att[a];
    if (at.skipped) continue;
    if (at.error) { problems.push("\"" + at.name + "\": " + at.error); continue; }
    if (compName && rgNorm(at.name).indexOf(rgNorm(compName)) < 0) {
      problems.push("\"" + at.name + "\": the file name must contain \"" + compName + "\".");
    }
    if (!at.protection.ok) problems.push("\"" + at.name + "\": " + at.protection.why + ".");
  }

  res.allow = problems.length === 0;
  return res;
}

// Downloads and inspects every attachment (only needed when sending outside).
function rgInspectAttachments() {
  return rgAttachments().then(function (list) {
    return Promise.all(list.map(function (a) {
      var name = a.name || "(unnamed)";
      var type = rgLower(a.attachmentType);
      var isImage = /^image\//.test(rgLower(a.contentType)) || /\.(png|jpe?g|gif|bmp)$/i.test(name);
      if (a.isInline && RG_CONFIG.SKIP_INLINE_IMAGES && isImage) {
        return { name: name, skipped: true };
      }
      if (type === "cloud") return { name: name, error: "shared links (OneDrive/SharePoint) can't be checked - attach a password-protected file instead" };
      if (type === "item") return { name: name, error: "attached emails can't be sent to clients - attach a password-protected file instead" };
      return rgAttachmentBytes(a.id).then(function (c) {
        if (c.kind === "cloud") return { name: name, error: "shared links can't be checked - attach a password-protected file instead" };
        if (c.kind === "item") return { name: name, error: "attached emails can't be sent to clients" };
        if (c.kind !== "file") return { name: name, error: "the attachment could not be read" };
        return { name: name, protection: rgCheckProtection(name, c.bytes) };
      }, function () {
        return { name: name, error: "the attachment could not be read, so it can't be checked" };
      });
    }));
  });
}

/* ========================= AS-YOU-TYPE ENFORCEMENT ========================= */

/*
 * Removes recipients that are not allowed and, for new emails, fills To with the
 * company's contacts when the company in the subject changes.
 * Never throws: it is a convenience layer. The Send check is the real control.
 */
var rgEnforcing = false, rgEnforceAgain = false, rgEnforceRun = null;

// If a change arrives while a pass is running, run once more afterwards so no
// change is missed (e.g. an address typed while the previous one was removed).
function rgEnforce(reason) {
  if (rgEnforcing) { rgEnforceAgain = true; return rgEnforceRun; }
  rgEnforcing = true;
  rgEnforceRun = rgEnforceOnce().then(function () {
    rgEnforcing = false;
    if (rgEnforceAgain) { rgEnforceAgain = false; return rgEnforce("again"); }
  });
  return rgEnforceRun;
}

function rgEnforceOnce() {
  var state = {};
  return Promise.all([rgGetComposeType(), rgGetSubject(), rgGetAllRecipients(),
                      rgSessionGet("rg_company"), rgSessionGet("rg_thread")])
    .then(function (v) {
      state.mode = rgModeOf(v[0]); state.subject = v[1]; state.recips = v[2];
      state.lastCompany = v[3] || ""; state.thread = v[4] || "";
      var anyExternal = [].concat(v[2].to, v[2].cc, v[2].bcc).some(function (r) {
        return !rgIsInternal(r.emailAddress);
      });
      // Repository only needed if there are externals, or for subject-driven filling.
      if (!anyExternal && state.mode !== "new" && !state.lastCompany) return null;
      return rgLoadRepository(false).catch(function () { return null; });
    })
    .then(function (repo) {
      state.repo = repo;
      var removed = [], adds = [], fields = ["to", "cc", "bcc"];
      var mode = state.mode;
      var companyKey = repo ? rgDetermineCompany(repo, mode, state.subject, state.thread).key : null;
      var allowed = {};
      if (repo && companyKey) repo.companies[companyKey].emails.forEach(function (e) { allowed[e] = 1; });

      // Company switching / removal for new emails.
      var prevList = [];
      if (repo && mode === "new" && state.lastCompany !== (companyKey || "")) {
        if (state.lastCompany && repo.companies[state.lastCompany]) prevList = repo.companies[state.lastCompany].emails;
        if (companyKey) adds = repo.companies[companyKey].emails.slice();
      }

      var newLists = {};
      fields.forEach(function (f) {
        newLists[f] = state.recips[f].filter(function (r) {
          var a = rgLower(r.emailAddress);
          if (rgIsInternal(a)) return true;
          var keep;
          if (!repo) keep = mode !== "forward";              // repo unreachable: Send check will block
          else if (mode === "forward") keep = false;
          else if (mode === "new" || mode === "reply") keep = !!(companyKey && allowed[a]);
          else keep = true;
          if (keep && prevList.indexOf(a) >= 0 && !allowed[a]) keep = false;
          if (!keep) removed.push(r.emailAddress || r.displayName);
          return keep;
        });
      });

      // Don't re-add contacts already present anywhere.
      var present = {};
      fields.forEach(function (f) { newLists[f].forEach(function (r) { present[rgLower(r.emailAddress)] = 1; }); });
      adds = adds.filter(function (a) { return !present[a]; });

      var ops = [];
      fields.forEach(function (f) {
        if (newLists[f].length !== state.recips[f].length) {
          ops.push(rgSetRecipients(f, newLists[f].map(function (r) {
            return { displayName: r.displayName || r.emailAddress, emailAddress: r.emailAddress };
          })));
        }
      });
      if (adds.length) ops.push(rgAddRecipients("to", adds));
      if (repo && mode === "new") ops.push(rgSessionSet("rg_company", companyKey || ""));

      return Promise.all(ops).then(function () {
        var K = "rgEnforce";
        if (removed.length) {
          var why = !repo ? "" :
            mode === "forward" ? "Forwards can only go to IRIS addresses. " :
            !companyKey ? (mode === "new" ? "Start the subject with the client's company name. " : "Client not identified. ") :
            "Only " + repo.companies[companyKey].name + " contacts are allowed. ";
          return rgNotify(K, why + "Removed: " + rgUnique(removed).join(", "));
        }
        if (adds.length && companyKey) {
          return rgNotify(K, "Added " + adds.length + " " + repo.companies[companyKey].name + " contact(s) to To.");
        }
      });
    })
    .catch(function () { /* never break composing */ });
}

/* ============================== EVENT HANDLERS ============================= */

function onNewMessageComposeHandler(event) {
  var done = false;
  function finish() { if (!done) { done = true; try { event.completed(); } catch (e) {} } }
  setTimeout(finish, 25000);
  Promise.all([rgGetComposeType(), rgGetAllRecipients(), rgGetSubject()])
    .then(function (v) {
      var mode = rgModeOf(v[0]);
      var tasks = [rgSessionSet("rg_mode", mode)];
      if (mode === "reply") {
        // Remember which client this thread belongs to.
        tasks.push(rgLoadRepository(false).then(function (repo) {
          var key = rgCompanyFromSubject(repo, v[2]);
          if (!key) {
            var cls = rgClassify(repo, v[1]);
            if (cls.companyKeys.length === 1) key = cls.companyKeys[0];
          }
          return rgSessionSet("rg_thread", key || "");
        }).catch(function () {}));
      }
      if (mode === "new") {
        tasks.push(rgNotify("rgWelcome",
          "IRIS Recipient Guard: start the subject with the client's company name to add their contacts.",
          { type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
            actions: [{ actionType: "showTaskPane", actionText: "Pick company", commandId: "rgOpenPane", contextData: "{}" }] }));
      }
      return Promise.all(tasks).then(function () { return rgEnforce("open"); });
    })
    .catch(function () {})
    .then(finish);
}

function onMessageRecipientsChangedHandler(event) {
  var done = false;
  function finish() { if (!done) { done = true; try { event.completed(); } catch (e) {} } }
  setTimeout(finish, 25000);
  rgEnforce("recipients").then(finish, finish);
}

function rgBlock(event, text) {
  var opts = { allowEvent: false, errorMessage: String(text).slice(0, 500) };
  try {
    if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
      opts.cancelLabel = "Open Recipient Guard";
      opts.commandId = "rgOpenPane";
    }
  } catch (e) {}
  event.completed(opts);
}

function onMessageSendHandler(event) {
  var finished = false;
  function allow() { if (!finished) { finished = true; event.completed({ allowEvent: true }); } }
  function block(t) { if (!finished) { finished = true; rgBlock(event, t); } }

  // Fail closed if anything takes too long.
  setTimeout(function () {
    block("IRIS Recipient Guard could not finish checking this email in time, so it has not been sent. Please try Send again.");
  }, RG_CONFIG.SEND_CHECK_TIMEOUT_MS);

  rgRunSendCheck().then(function (res) {
    if (res.allow) allow();
    else block("Not sent - IRIS Recipient Guard:\n• " + res.problems.join("\n• "));
  }, function (err) {
    block("Not sent - IRIS Recipient Guard could not complete its checks (" +
      (err && err.message ? err.message : "unknown error") + "). Please try again or contact IT.");
  });
}

// Gathers everything, downloads a FRESH client list, and validates.
function rgRunSendCheck() {
  var ctx = {};
  return Promise.all([rgGetComposeType(), rgGetSubject(), rgGetAllRecipients(), rgSessionGet("rg_thread"), rgSessionGet("rg_mode")])
    .then(function (v) {
      var mode = rgModeOf(v[0]);
      if (mode === "unknown" && v[4]) mode = v[4];
      ctx.mode = mode; ctx.subject = v[1]; ctx.recipients = v[2]; ctx.threadCompany = v[3] || "";
      var externals = [].concat(v[2].to, v[2].cc, v[2].bcc).filter(function (r) { return !rgIsInternal(r.emailAddress); });
      if (!externals.length) { ctx.internalOnly = true; return null; }
      return rgLoadRepository(true).catch(function (e) {
        throw new Error("the client list could not be loaded: " + (e && e.message ? e.message : e));
      });
    })
    .then(function (repo) {
      if (ctx.internalOnly) return { allow: true, problems: [], mode: ctx.mode };
      ctx.repo = repo;
      return rgInspectAttachments().then(function (att) {
        ctx.attachments = att;
        return rgValidate(ctx);
      });
    });
}

/* ============================ REGISTER HANDLERS ============================ */

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    RG_CONFIG: RG_CONFIG, rgBuildRepository: rgBuildRepository, rgCompanyFromSubject: rgCompanyFromSubject,
    rgSubjectWithCompany: rgSubjectWithCompany, rgCheckProtection: rgCheckProtection, rgValidate: rgValidate,
    rgBase64ToBytes: rgBase64ToBytes, rgStripSubjectPrefixes: rgStripSubjectPrefixes,
    onNewMessageComposeHandler: onNewMessageComposeHandler,
    onMessageRecipientsChangedHandler: onMessageRecipientsChangedHandler,
    onMessageSendHandler: onMessageSendHandler, rgEnforce: rgEnforce, rgRunSendCheck: rgRunSendCheck,
    _cache: function () { return rgRepoCache; }, _resetCache: function () { rgRepoCache = { repo: null, at: 0, error: null }; }
  };
}
