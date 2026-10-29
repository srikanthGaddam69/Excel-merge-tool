/*!
 * Excel Merge Tool — merge-core.js
 * Pure merge logic, shared by the browser UI and the Node test-suite.
 *
 *  1. buildPlan()  – works out WHICH cells change and to WHAT (no file I/O).
 *  2. patchXlsx()  – writes those values straight into the original .xlsx
 *                    package (XML level), so every style, colour, border,
 *                    freeze pane, filter, validation, chart and pivot stays
 *                    exactly as it was.
 *
 * MIT License — see LICENSE.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MergeCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var EXCEL_CHAR_LIMIT = 32767;
  var MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var XML_NS = 'http://www.w3.org/XML/1998/namespace';
  var XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  /* Characters that are illegal in XML 1.0 and make Excel report a corrupt file */
  var BAD_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

  function isEmpty(v) {
    return v === undefined || v === null || String(v).trim() === '';
  }

  function cellText(cell) {
    if (!cell || cell.v === undefined || cell.v === null) return '';
    if (cell.t === 'e') return '';
    return cell.w !== undefined && cell.w !== null ? String(cell.w) : String(cell.v);
  }

  function normKey(v, ignoreCase) {
    var s = v === undefined || v === null ? '' : String(v).trim();
    return ignoreCase ? s.toUpperCase() : s;
  }

  /**
   * Read a worksheet as a table whose header sits on `headerRow` (1-based).
   * Duplicate / blank headers are made unique with their column letter.
   */
  function readTable(XLSX, ws, headerRow) {
    var out = { headers: [], rows: [] };
    if (!ws || !ws['!ref']) return out;
    var range = XLSX.utils.decode_range(ws['!ref']);
    var hr = Math.max(0, (headerRow || 1) - 1);
    var seen = {};
    for (var c = range.s.c; c <= range.e.c; c++) {
      var hc = ws[XLSX.utils.encode_cell({ r: hr, c: c })];
      if (!hc || isEmpty(hc.v)) continue;
      var name = cellText(hc).trim();
      if (seen[name]) name = name + ' (' + XLSX.utils.encode_col(c) + ')';
      seen[name] = true;
      out.headers.push({ name: name, col: c });
    }
    for (var r = hr + 1; r <= range.e.r; r++) {
      var cells = {};
      var any = false;
      for (var i = 0; i < out.headers.length; i++) {
        var h = out.headers[i];
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: h.col })] || null;
        cells[h.name] = cell;
        if (cell && !isEmpty(cell.v)) any = true;
      }
      if (any) out.rows.push({ r: r, cells: cells });
    }
    return out;
  }

  /** Is a source cell a plain number (not a date)? */
  function isPlainNumber(XLSX, cell) {
    if (!cell || cell.t !== 'n') return false;
    if (cell.z && XLSX.SSF && XLSX.SSF.is_date && XLSX.SSF.is_date(cell.z)) return false;
    return true;
  }

  function clean(s) {
    return String(s).replace(BAD_XML_CHARS, '');
  }

  /**
   * Work out every cell change.
   * opts = { origHeaderRow, sharedHeaderRow, origKey, sharedKey,
   *          ignoreCase, rules:[{src,tgt,op:'replace'|'concat',sep}] }
   */
  function buildPlan(XLSX, origWs, sharedWs, opts) {
    var orig = readTable(XLSX, origWs, opts.origHeaderRow);
    var shared = readTable(XLSX, sharedWs, opts.sharedHeaderRow);
    var origCol = {};
    orig.headers.forEach(function (h) { origCol[h.name] = h.col; });
    var sharedNames = {};
    shared.headers.forEach(function (h) { sharedNames[h.name] = true; });

    if (!(opts.origKey in origCol)) throw new Error('Key column "' + opts.origKey + '" not found in the original sheet.');
    if (!sharedNames[opts.sharedKey]) throw new Error('Key column "' + opts.sharedKey + '" not found in the shared sheet.');
    (opts.rules || []).forEach(function (rule) {
      if (!sharedNames[rule.src]) throw new Error('Source column "' + rule.src + '" not found in the shared sheet.');
      if (!(rule.tgt in origCol)) throw new Error('Target column "' + rule.tgt + '" not found in the original sheet.');
    });

    var ic = !!opts.ignoreCase;
    var combine = !!opts.combineDuplicates;
    var combineSep = opts.combineSep === undefined || opts.combineSep === null ? ' | ' : String(opts.combineSep);
    var sharedIdx = {};   /* key -> array of shared rows, in sheet order */
    var dupShared = {};
    var dupExtraRows = 0;
    shared.rows.forEach(function (row) {
      var k = normKey(cellText(row.cells[opts.sharedKey]), ic);
      if (!k) return;
      if (sharedIdx[k]) { dupShared[k] = true; dupExtraRows++; sharedIdx[k].push(row); }
      else sharedIdx[k] = [row];
    });

    /*
     * Value(s) for one rule from the shared rows of one key.
     * Without combine: the last row wins (blank there = blank).
     * With combine: every non-blank value, repeats removed, joined by combineSep.
     * Returns null when there is nothing to write.
     */
    function sourceFor(rows, srcName) {
      var usable = function (c) { return c && !isEmpty(c.v) && c.t !== 'e'; };
      if (!combine || rows.length === 1) {
        var last = rows[rows.length - 1].cells[srcName];
        return usable(last) ? { cell: last, text: cellText(last).trim(), combined: 0 } : null;
      }
      var seen = {}, texts = [], firstCell = null;
      rows.forEach(function (r) {
        var c = r.cells[srcName];
        if (!usable(c)) return;
        var t = cellText(c).trim();
        if (seen[t]) return;
        seen[t] = true;
        texts.push(t);
        if (!firstCell) firstCell = c;
      });
      if (!texts.length) return null;
      if (texts.length === 1) return { cell: firstCell, text: texts[0], combined: 0 };
      return { cell: null, text: texts.join(combineSep), combined: texts.length };
    }

    var stats = {
      totalRows: 0, matchedRows: 0, unmatchedRows: 0, blankKeyRows: 0,
      cellsUpdated: 0, blanksRetained: 0, unchanged: 0, truncated: 0,
      duplicateKeysShared: Object.keys(dupShared).length, duplicateExtraRows: dupExtraRows,
      cellsCombined: 0
    };
    var seenOrig = {};
    var dupOrig = {};
    var updates = [];
    var truncatedCells = [];
    var matchedKeys = {};

    orig.rows.forEach(function (row) {
      stats.totalRows++;
      var k = normKey(cellText(row.cells[opts.origKey]), ic);
      if (!k) { stats.blankKeyRows++; return; }
      if (seenOrig[k]) dupOrig[k] = true;
      seenOrig[k] = true;
      var sRows = sharedIdx[k];
      if (!sRows) { stats.unmatchedRows++; return; }
      stats.matchedRows++;
      matchedKeys[k] = true;

      (opts.rules || []).forEach(function (rule) {
        var src = sourceFor(sRows, rule.src);
        if (!src) { stats.blanksRetained++; return; }

        var tgtCellObj = row.cells[rule.tgt];
        var oldText = cellText(tgtCellObj);
        var srcText = src.text;
        var value, type;

        if (rule.op === 'concat' && !isEmpty(oldText)) {
          value = oldText.replace(/\s+$/, '') + (rule.sep === undefined ? ' | ' : rule.sep) + srcText;
          type = 's';
        } else if (src.cell && isPlainNumber(XLSX, src.cell)) {
          value = src.cell.v;
          type = 'n';
        } else {
          value = srcText;
          type = 's';
        }

        if (type === 's') {
          value = clean(value);
          if (value.length > EXCEL_CHAR_LIMIT) {
            stats.truncated++;
            truncatedCells.push({ key: k, column: rule.tgt, excelRow: row.r + 1, chars: value.length });
            value = value.slice(0, EXCEL_CHAR_LIMIT);
          }
        }

        var oldCompare = tgtCellObj && tgtCellObj.t === 'n' ? tgtCellObj.v : oldText;
        if (String(oldCompare) === String(value)) { stats.unchanged++; return; }

        stats.cellsUpdated++;
        if (src.combined) stats.cellsCombined++;
        updates.push({
          r: row.r, c: origCol[rule.tgt],
          ref: XLSX.utils.encode_cell({ r: row.r, c: origCol[rule.tgt] }),
          key: k, column: rule.tgt, op: rule.op,
          oldValue: oldText, value: value, type: type,
          combinedFrom: src.combined || 0
        });
      });
    });

    var sharedOnly = Object.keys(sharedIdx).filter(function (k) { return !matchedKeys[k] && !seenOrig[k]; });

    return {
      updates: updates,
      stats: stats,
      truncatedCells: truncatedCells,
      duplicateKeysShared: Object.keys(dupShared),
      duplicateKeysOriginal: Object.keys(dupOrig),
      keysOnlyInShared: sharedOnly
    };
  }

  /** Apply the plan to a SheetJS worksheet (used for the .xls/.csv fallback path). */
  function applyToSheetJS(ws, updates) {
    updates.forEach(function (u) {
      var old = ws[u.ref] || {};
      var cell = { t: u.type, v: u.value };
      if (old.s) cell.s = old.s;
      if (old.z && u.type === 'n') cell.z = old.z;
      ws[u.ref] = cell;
    });
  }

  /* ---------- XLSX package patching (format-preserving) ---------- */

  function colToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }
  function numToCol(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function splitRef(ref) {
    var m = /^([A-Z]+)(\d+)$/.exec(ref);
    return m ? { col: colToNum(m[1]), row: parseInt(m[2], 10) } : null;
  }
  function childElements(node, localName) {
    var out = [];
    for (var n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && (!localName || n.localName === localName)) out.push(n);
    }
    return out;
  }
  function resolveTarget(target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = ('xl/' + target).split('/');
    var stack = [];
    parts.forEach(function (p) { if (p === '..') stack.pop(); else if (p && p !== '.') stack.push(p); });
    return stack.join('/');
  }
  function serialize(Ser, doc) {
    var s = new Ser().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
    return XML_DECL + s;
  }

  /**
   * Write `updates` into sheet `sheetName` of the .xlsx/.xlsm in `data`.
   * env = { JSZip, DOMParser, XMLSerializer }
   * Resolves to { zip, report:{ written, formulasReplaced, skipped:[...] } }
   */
  function patchXlsx(data, sheetName, updates, env) {
    var JSZip = env.JSZip, Parser = env.DOMParser, Ser = env.XMLSerializer;
    var parse = function (xml) { return new Parser().parseFromString(xml, 'application/xml'); };
    var zip, wbDoc, relsDoc, sheetPath, sheetDoc;
    var report = { written: 0, formulasReplaced: 0, skipped: [] };

    return JSZip.loadAsync(data).then(function (z) {
      zip = z;
      var wbFile = zip.file('xl/workbook.xml');
      var relFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!wbFile || !relFile) throw new Error('This file is not a standard .xlsx workbook.');
      return Promise.all([wbFile.async('string'), relFile.async('string')]);
    }).then(function (xml) {
      wbDoc = parse(xml[0]);
      relsDoc = parse(xml[1]);
      var rid = null;
      var sheets = wbDoc.getElementsByTagNameNS(wbDoc.documentElement.namespaceURI, 'sheet');
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getAttribute('name') === sheetName) {
          var attrs = sheets[i].attributes;
          for (var a = 0; a < attrs.length; a++) {
            if (attrs[a].localName === 'id' && attrs[a].namespaceURI && attrs[a].namespaceURI.indexOf('relationships') >= 0) rid = attrs[a].value;
          }
        }
      }
      if (!rid) throw new Error('Sheet "' + sheetName + '" not found inside the workbook.');
      var rels = relsDoc.getElementsByTagName('Relationship');
      for (var j = 0; j < rels.length; j++) {
        if (rels[j].getAttribute('Id') === rid) sheetPath = resolveTarget(rels[j].getAttribute('Target'));
      }
      if (!sheetPath || !zip.file(sheetPath)) throw new Error('Worksheet part for "' + sheetName + '" is missing.');
      return zip.file(sheetPath).async('string');
    }).then(function (sheetXml) {
      sheetDoc = parse(sheetXml);
      var root = sheetDoc.documentElement;
      var ns = root.namespaceURI || MAIN_NS;
      var sheetData = sheetDoc.getElementsByTagNameNS(ns, 'sheetData')[0];
      if (!sheetData) throw new Error('Worksheet has no sheetData section.');
      var pfx = sheetData.prefix ? sheetData.prefix + ':' : '';
      var mk = function (name) { return sheetDoc.createElementNS(ns, pfx + name); };

      /* Index rows (and give implicit rows/cells explicit refs) */
      var rowMap = {};
      var implicitRow = 0;
      childElements(sheetData, 'row').forEach(function (rowEl) {
        var rAttr = rowEl.getAttribute('r');
        var rn = rAttr ? parseInt(rAttr, 10) : implicitRow + 1;
        if (!rAttr) rowEl.setAttribute('r', String(rn));
        implicitRow = rn;
        rowMap[rn] = rowEl;
      });

      var cellCache = {};
      function cellsOf(rowEl, rn) {
        if (cellCache[rn]) return cellCache[rn];
        var map = {};
        var implicitCol = 0;
        childElements(rowEl, 'c').forEach(function (cEl) {
          var ref = cEl.getAttribute('r');
          var cn = ref ? splitRef(ref).col : implicitCol + 1;
          if (!ref) cEl.setAttribute('r', numToCol(cn) + rn);
          implicitCol = cn;
          map[cn] = cEl;
        });
        cellCache[rn] = map;
        return map;
      }

      function getRow(rn) {
        if (rowMap[rn]) return rowMap[rn];
        var rowEl = mk('row');
        rowEl.setAttribute('r', String(rn));
        var next = null;
        childElements(sheetData, 'row').some(function (el) {
          if (parseInt(el.getAttribute('r'), 10) > rn) { next = el; return true; }
          return false;
        });
        sheetData.insertBefore(rowEl, next);
        rowMap[rn] = rowEl;
        return rowEl;
      }

      function getCell(rowEl, rn, cn) {
        var map = cellsOf(rowEl, rn);
        if (map[cn]) return map[cn];
        var cEl = mk('c');
        cEl.setAttribute('r', numToCol(cn) + rn);
        if (rowEl.getAttribute('customFormat') === '1' && rowEl.getAttribute('s')) cEl.setAttribute('s', rowEl.getAttribute('s'));
        var next = null;
        childElements(rowEl, 'c').some(function (el) {
          if (splitRef(el.getAttribute('r')).col > cn) { next = el; return true; }
          return false;
        });
        rowEl.insertBefore(cEl, next);
        map[cn] = cEl;
        return cEl;
      }

      updates.forEach(function (u) {
        var rn = u.r + 1, cn = u.c + 1;
        var rowEl = getRow(rn);
        var cEl = getCell(rowEl, rn, cn);
        var f = childElements(cEl, 'f')[0];
        if (f && f.getAttribute('t') === 'shared' && f.getAttribute('ref')) {
          report.skipped.push({ ref: u.ref, key: u.key, column: u.column, reason: 'cell is the master of a shared formula' });
          return;
        }
        if (f) report.formulasReplaced++;
        while (cEl.firstChild) cEl.removeChild(cEl.firstChild);
        cEl.removeAttribute('t');
        cEl.removeAttribute('cm');
        cEl.removeAttribute('vm');
        if (u.type === 'n') {
          var v = mk('v');
          v.appendChild(sheetDoc.createTextNode(String(u.value)));
          cEl.appendChild(v);
        } else {
          cEl.setAttribute('t', 'inlineStr');
          var is = mk('is');
          var t = mk('t');
          t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
          t.appendChild(sheetDoc.createTextNode(String(u.value)));
          is.appendChild(t);
          cEl.appendChild(is);
        }
        report.written++;
      });

      zip.file(sheetPath, serialize(Ser, sheetDoc));

      /* Ask Excel to recalculate on open so status formulas pick up new values */
      var wbNs = wbDoc.documentElement.namespaceURI;
      var calcPr = wbDoc.getElementsByTagNameNS(wbNs, 'calcPr')[0];
      if (!calcPr) {
        var wbRoot = wbDoc.documentElement;
        var wpfx = wbRoot.prefix ? wbRoot.prefix + ':' : '';
        calcPr = wbDoc.createElementNS(wbNs, wpfx + 'calcPr');
        var after = null;
        childElements(wbRoot).forEach(function (el) {
          if (['sheets', 'functionGroups', 'externalReferences', 'definedNames'].indexOf(el.localName) >= 0) after = el;
        });
        wbRoot.insertBefore(calcPr, after ? after.nextSibling : null);
      }
      calcPr.setAttribute('fullCalcOnLoad', '1');
      zip.file('xl/workbook.xml', serialize(Ser, wbDoc));

      /* Removing a formula invalidates the calculation chain; Excel rebuilds it */
      if (report.formulasReplaced > 0 && zip.file('xl/calcChain.xml')) {
        zip.remove('xl/calcChain.xml');
        var rels = relsDoc.getElementsByTagName('Relationship');
        for (var i = rels.length - 1; i >= 0; i--) {
          if (/calcChain$/.test(rels[i].getAttribute('Type') || '')) rels[i].parentNode.removeChild(rels[i]);
        }
        zip.file('xl/_rels/workbook.xml.rels', serialize(Ser, relsDoc));
        return zip.file('[Content_Types].xml').async('string').then(function (ct) {
          var ctDoc = parse(ct);
          var ovs = ctDoc.getElementsByTagName('Override');
          for (var k = ovs.length - 1; k >= 0; k--) {
            if (ovs[k].getAttribute('PartName') === '/xl/calcChain.xml') ovs[k].parentNode.removeChild(ovs[k]);
          }
          zip.file('[Content_Types].xml', serialize(Ser, ctDoc));
        });
      }
    }).then(function () {
      return { zip: zip, report: report };
    });
  }

  /** Build a CSV change log (audit trail) from a plan. */
  function changeLogCsv(plan, skipped) {
    var skip = {};
    (skipped || []).forEach(function (s) { skip[s.ref] = s.reason; });
    var q = function (v) { return '"' + String(v === undefined || v === null ? '' : v).replace(/"/g, '""') + '"'; };
    var lines = [['Key', 'Target column', 'Cell', 'Operation', 'Old value', 'New value', 'Note'].map(q).join(',')];
    plan.updates.forEach(function (u) {
      var note = skip[u.ref] ? 'SKIPPED: ' + skip[u.ref] : (u.combinedFrom ? 'Combined ' + u.combinedFrom + ' values from duplicate IDs' : '');
      lines.push([u.key, u.column, u.ref, u.op, u.oldValue, u.value, note].map(q).join(','));
    });
    return '﻿' + lines.join('\r\n');
  }

  return {
    EXCEL_CHAR_LIMIT: EXCEL_CHAR_LIMIT,
    isEmpty: isEmpty,
    cellText: cellText,
    readTable: readTable,
    buildPlan: buildPlan,
    applyToSheetJS: applyToSheetJS,
    patchXlsx: patchXlsx,
    changeLogCsv: changeLogCsv
  };
});
