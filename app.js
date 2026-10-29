/*!
 * Excel Merge Tool — app.js (browser UI)
 * Depends on: XLSX (SheetJS), JSZip, MergeCore. MIT License.
 */
(function () {
  'use strict';

  var KEY_HINTS = ['reqt_id', 'req_id', 'requirement id', 'requirement_id', 'id', 'key'];
  var state = {
    orig: null,   /* { name, ext, buf, wb } */
    shared: null,
    lastPlan: null,
    lastReport: null
  };

  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  };

  /* ---------- helpers ---------- */
  function extOf(name) { var m = /\.([^.]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
  function baseOf(name) { return String(name || 'output').replace(/\.[^.]+$/, ''); }

  function setStep(n, on) {
    $('step' + n).classList.toggle('disabled', !on);
    $('b' + n).classList.toggle('done', on);
  }

  function msg(el, kind, html) { el.innerHTML = '<div class="msg ' + kind + '">' + html + '</div>'; }

  function fillSelect(sel, names, preferred) {
    var cur = preferred !== undefined ? preferred : sel.value;
    sel.innerHTML = names.map(function (n) { return '<option>' + esc(n) + '</option>'; }).join('');
    if (cur && names.indexOf(cur) >= 0) sel.value = cur;
  }

  function headersOf(side) {
    var f = state[side];
    if (!f) return [];
    var sheet = $(side + 'Sheet').value;
    var hr = parseInt($(side + 'Header').value, 10) || 1;
    return MergeCore.readTable(XLSX, f.wb.Sheets[sheet], hr).headers.map(function (h) { return h.name; });
  }

  function guessKey(names, otherNames) {
    var lower = names.map(function (n) { return n.toLowerCase(); });
    for (var i = 0; i < KEY_HINTS.length; i++) {
      var idx = lower.indexOf(KEY_HINTS[i]);
      if (idx >= 0) return names[idx];
    }
    if (otherNames) for (var j = 0; j < names.length; j++) if (otherNames.indexOf(names[j]) >= 0) return names[j];
    return names[0];
  }

  /* ---------- file loading ---------- */
  function loadFile(side, file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var buf = new Uint8Array(e.target.result);
        var wb = XLSX.read(buf, { type: 'array', cellNF: true, cellDates: false });
        state[side] = { name: file.name, ext: extOf(file.name), buf: buf, wb: wb };
        $(side + 'Name').textContent = file.name + '  ·  ' + wb.SheetNames.length + ' sheet(s)';
        $(side + 'Loaded').style.display = 'flex';
        fillSelect($(side + 'Sheet'), wb.SheetNames);
        refreshHeaders(side);
        if (side === 'orig') $('outName').value = baseOf(file.name) + '_merged.' + (state.orig.ext === 'xlsm' ? 'xlsm' : 'xlsx');
        setStep(2, !!(state.orig && state.shared));
        $('loadMsg').innerHTML = '';
        if (side === 'orig' && ['xlsx', 'xlsm'].indexOf(state.orig.ext) < 0) {
          msg($('loadMsg'), 'warn', '⚠ <strong>' + esc(file.name) + '</strong> is not .xlsx/.xlsm. The merge works, but cell formatting cannot be preserved for this format. Save it as .xlsx first for a format-identical result.');
        }
      } catch (err) {
        msg($('loadMsg'), 'err', 'Could not read <strong>' + esc(file.name) + '</strong>: ' + esc(err.message));
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function refreshHeaders(side) {
    var names = headersOf(side);
    var other = side === 'orig' ? headersOf('shared') : headersOf('orig');
    var keySel = $(side + 'Key');
    var cur = keySel.value;
    fillSelect(keySel, names, names.indexOf(cur) >= 0 ? cur : guessKey(names, other));
    document.querySelectorAll('#rulesBody tr').forEach(function (tr) {
      fillSelect(tr.querySelector(side === 'shared' ? '.src-col' : '.tgt-col'), names);
    });
  }

  /* ---------- options ---------- */
  function currentOpts() {
    return {
      origSheet: $('origSheet').value,
      sharedSheet: $('sharedSheet').value,
      origHeaderRow: parseInt($('origHeader').value, 10) || 1,
      sharedHeaderRow: parseInt($('sharedHeader').value, 10) || 1,
      origKey: $('origKey').value,
      sharedKey: $('sharedKey').value,
      ignoreCase: $('ignoreCase').checked,
      combineDuplicates: $('combineDup').checked,
      combineSep: $('combineSep').value,
      rules: getRules()
    };
  }

  /* ---------- preview ---------- */
  function buildPreview() {
    try {
      var o = currentOpts();
      var plan = MergeCore.buildPlan(XLSX, state.orig.wb.Sheets[o.origSheet], state.shared.wb.Sheets[o.sharedSheet],
        Object.assign({}, o, { rules: [] }));
      var s = plan.stats;
      var info = '<strong>' + s.matchedRows + '</strong> rows matched · <strong>' + s.unmatchedRows + '</strong> not in shared file (left unchanged)';
      if (s.blankKeyRows) info += ' · ' + s.blankKeyRows + ' rows with blank key (skipped)';
      if (plan.keysOnlyInShared.length) info += '<br>' + plan.keysOnlyInShared.length + ' key(s) exist only in the shared file and will be ignored: ' + esc(plan.keysOnlyInShared.slice(0, 8).join(', ')) + (plan.keysOnlyInShared.length > 8 ? '…' : '');
      if (plan.duplicateKeysShared.length) {
        info += '<br>⚠ <strong>' + s.duplicateKeysShared + '</strong> ID(s) appear more than once in the shared file (' + s.duplicateExtraRows + ' extra row(s)) — ' +
          (o.combineDuplicates ? 'their values will be <strong>combined</strong>' : 'only the <strong>last row</strong> of each will be used') +
          ': ' + esc(plan.duplicateKeysShared.slice(0, 8).join(', ')) + (plan.duplicateKeysShared.length > 8 ? '…' : '');
      } else {
        info += '<br>No duplicate IDs in the shared file.';
      }
      if (plan.duplicateKeysOriginal.length) info += '<br>⚠ Duplicate keys in original (every copy is updated): ' + esc(plan.duplicateKeysOriginal.slice(0, 8).join(', '));
      $('matchInfo').innerHTML = info;

      var t = MergeCore.readTable(XLSX, state.orig.wb.Sheets[o.origSheet], o.origHeaderRow);
      var sharedT = MergeCore.readTable(XLSX, state.shared.wb.Sheets[o.sharedSheet], o.sharedHeaderRow);
      var keys = {};
      sharedT.rows.forEach(function (r) {
        var k = MergeCore.cellText(r.cells[o.sharedKey]).trim();
        keys[o.ignoreCase ? k.toUpperCase() : k] = true;
      });
      var cols = [o.origKey].concat(t.headers.map(function (h) { return h.name; }).filter(function (n) { return n !== o.origKey; }).slice(0, 3));
      var html = '<thead><tr><th>Row</th>' + cols.map(function (c) { return '<th>' + esc(c) + '</th>'; }).join('') + '<th>Match result</th></tr></thead><tbody>';
      t.rows.slice(0, 50).forEach(function (row) {
        var k = MergeCore.cellText(row.cells[o.origKey]).trim();
        var hit = !!keys[o.ignoreCase ? k.toUpperCase() : k];
        html += '<tr><td class="muted">' + (row.r + 1) + '</td>' + cols.map(function (c) {
          var v = MergeCore.cellText(row.cells[c]);
          return '<td>' + esc(v.length > 60 ? v.slice(0, 60) + '…' : v) + '</td>';
        }).join('') + '<td><span class="' + (hit ? 'tag-match' : 'tag-miss') + '">' + (hit ? '✓ match' : '— no match') + '</span></td></tr>';
      });
      $('prevTable').innerHTML = html + '</tbody>';
      $('previewOut').style.display = 'block';
      setStep(3, true);
      setStep(4, true);
      if (!$('rulesBody').children.length) addRule();
    } catch (err) {
      $('previewOut').style.display = 'block';
      $('matchInfo').innerHTML = '<span class="red">' + esc(err.message) + '</span>';
    }
  }

  /* ---------- rules ---------- */
  function addRule(rule) {
    var srcCols = headersOf('shared'), tgtCols = headersOf('orig');
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td><select class="src-col"></select></td>' +
      '<td><select class="tgt-col"></select></td>' +
      '<td><div class="op-cell"><div class="op-seg">' +
      '<button type="button" data-op="replace" class="active">Replace</button>' +
      '<button type="button" data-op="concat">Concat</button></div>' +
      '<span class="sep-wrap"><span class="sep-label">sep</span><input type="text" class="sep-val" value=" | "></span></div></td>' +
      '<td><button type="button" class="del-btn" title="Remove rule">×</button></td>';
    fillSelect(tr.querySelector('.src-col'), srcCols, rule ? rule.src : undefined);
    fillSelect(tr.querySelector('.tgt-col'), tgtCols, rule ? rule.tgt : undefined);
    tr.querySelectorAll('.op-seg button').forEach(function (b) {
      b.addEventListener('click', function () { setOp(tr, b.getAttribute('data-op')); });
    });
    tr.querySelector('.del-btn').addEventListener('click', function () { tr.remove(); });
    $('rulesBody').appendChild(tr);
    if (rule) {
      setOp(tr, rule.op || 'replace');
      if (rule.sep !== undefined) tr.querySelector('.sep-val').value = rule.sep;
    }
  }

  function setOp(tr, op) {
    tr.querySelectorAll('.op-seg button').forEach(function (b) { b.classList.toggle('active', b.getAttribute('data-op') === op); });
    tr.querySelector('.sep-wrap').style.display = op === 'concat' ? 'flex' : 'none';
  }

  function getRules() {
    return Array.prototype.map.call(document.querySelectorAll('#rulesBody tr'), function (tr) {
      var act = tr.querySelector('.op-seg button.active');
      return {
        src: tr.querySelector('.src-col').value,
        tgt: tr.querySelector('.tgt-col').value,
        op: act ? act.getAttribute('data-op') : 'replace',
        sep: tr.querySelector('.sep-val').value
      };
    });
  }

  function autoMap() {
    var src = headersOf('shared'), tgt = headersOf('orig');
    var o = currentOpts();
    var existing = {};
    getRules().forEach(function (r) { existing[r.tgt] = true; });
    var added = 0;
    src.forEach(function (n) {
      if (n === o.sharedKey || n === o.origKey) return;
      if (tgt.indexOf(n) >= 0 && !existing[n]) { addRule({ src: n, tgt: n, op: 'replace' }); added++; }
    });
    $('ruleMsg').innerHTML = added ? '<div class="msg ok">Added ' + added + ' rule(s) for columns with identical names.</div>' : '<div class="msg info">No new identically-named columns found.</div>';
  }

  /* ---------- save / load a rule set ---------- */
  function saveConfig() {
    var o = currentOpts();
    var cfg = { tool: 'excel-merge-tool', version: 2, saved: new Date().toISOString(), settings: o };
    download(new Blob([JSON.stringify(cfg, null, 2)], { type: 'application/json' }), 'merge-rules.json');
  }

  function loadConfig(file) {
    var r = new FileReader();
    r.onload = function () {
      try {
        var cfg = JSON.parse(r.result);
        var s = cfg.settings || cfg;
        if (s.origSheet && state.orig && state.orig.wb.SheetNames.indexOf(s.origSheet) >= 0) $('origSheet').value = s.origSheet;
        if (s.sharedSheet && state.shared && state.shared.wb.SheetNames.indexOf(s.sharedSheet) >= 0) $('sharedSheet').value = s.sharedSheet;
        if (s.origHeaderRow) $('origHeader').value = s.origHeaderRow;
        if (s.sharedHeaderRow) $('sharedHeader').value = s.sharedHeaderRow;
        $('ignoreCase').checked = !!s.ignoreCase;
        if (s.combineDuplicates !== undefined) $('combineDup').checked = !!s.combineDuplicates;
        if (s.combineSep !== undefined) $('combineSep').value = s.combineSep;
        syncCombine();
        refreshHeaders('orig');
        refreshHeaders('shared');
        if (s.origKey) $('origKey').value = s.origKey;
        if (s.sharedKey) $('sharedKey').value = s.sharedKey;
        buildPreview();
        $('rulesBody').innerHTML = '';
        var missing = [];
        var src = headersOf('shared'), tgt = headersOf('orig');
        (s.rules || []).forEach(function (rule) {
          if (src.indexOf(rule.src) < 0 || tgt.indexOf(rule.tgt) < 0) missing.push(rule.src + ' → ' + rule.tgt);
          else addRule(rule);
        });
        $('ruleMsg').innerHTML = missing.length
          ? '<div class="msg warn">Loaded rules, but these columns were not found and were skipped: ' + esc(missing.join('; ')) + '</div>'
          : '<div class="msg ok">Loaded ' + (s.rules || []).length + ' rule(s) from ' + esc(file.name) + '.</div>';
      } catch (err) {
        $('ruleMsg').innerHTML = '<div class="msg err">Could not load rules: ' + esc(err.message) + '</div>';
      }
    };
    r.readAsText(file);
  }

  /* ---------- process ---------- */
  function download(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }

  function process() {
    var btn = $('processBtn'), out = $('processMsg');
    var o = currentOpts();
    if (!o.rules.length) { msg(out, 'err', 'Add at least one column rule in step 3.'); return; }
    btn.disabled = true;
    btn.textContent = 'Processing…';
    out.innerHTML = '';

    var plan;
    var outName = ($('outName').value || 'merged.xlsx').trim();
    var isZip = ['xlsx', 'xlsm'].indexOf(state.orig.ext) >= 0;
    if (!isZip && !/\.xlsx$/i.test(outName)) outName = baseOf(outName) + '.xlsx';

    Promise.resolve().then(function () {
      plan = MergeCore.buildPlan(XLSX, state.orig.wb.Sheets[o.origSheet], state.shared.wb.Sheets[o.sharedSheet], o);
      if (isZip) {
        return MergeCore.patchXlsx(state.orig.buf, o.origSheet, plan.updates, {
          JSZip: JSZip, DOMParser: DOMParser, XMLSerializer: XMLSerializer
        }).then(function (res) {
          return res.zip.generateAsync({
            type: 'blob', compression: 'DEFLATE',
            mimeType: state.orig.ext === 'xlsm'
              ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
              : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }).then(function (blob) { return { blob: blob, report: res.report }; });
        });
      }
      var wb = XLSX.read(state.orig.buf, { type: 'array', cellNF: true });
      MergeCore.applyToSheetJS(wb.Sheets[o.origSheet], plan.updates);
      var data = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      return { blob: new Blob([data], { type: 'application/octet-stream' }), report: { written: plan.updates.length, formulasReplaced: 0, skipped: [] } };
    }).then(function (res) {
      state.lastPlan = plan;
      state.lastReport = res.report;
      download(res.blob, outName);
      renderSummary(plan, res.report, outName, isZip);
    }).catch(function (err) {
      msg(out, 'err', 'Error: ' + esc(err.message));
    }).then(function () {
      btn.disabled = false;
      btn.textContent = 'Process & download';
    });
  }

  function renderSummary(plan, report, outName, isZip) {
    var s = plan.stats;
    var written = report.written;
    var html = '<div class="msg ok">✓ Downloaded <strong>' + esc(outName) + '</strong>' +
      (isZip ? ' — only the changed cells were touched; all formatting kept.' : ' — values merged (formatting not kept for this input format).') + '</div>' +
      '<div class="stat-row">' +
      stat(written, 'cells updated', 'green') +
      stat(s.blanksRetained, 'blanks retained', 'amber') +
      stat(s.unchanged, 'already identical', 'blue') +
      stat(s.unmatchedRows, 'rows unmatched', 'red') +
      stat(s.truncated, 'cells truncated', 'purple') +
      '</div>';
    if (s.cellsCombined) html += '<div class="msg info">' + s.cellsCombined + ' updated cell(s) combine values from duplicate IDs in the shared file (marked in the change log).</div>';
    if (report.formulasReplaced) html += '<div class="msg warn">⚠ ' + report.formulasReplaced + ' target cell(s) held a formula and now hold the merged value instead.</div>';
    if (report.skipped.length) html += '<div class="msg warn">⚠ ' + report.skipped.length + ' cell(s) skipped because they anchor a shared formula: ' + esc(report.skipped.slice(0, 10).map(function (x) { return x.ref; }).join(', ')) + '. See the change log.</div>';
    if (plan.truncatedCells.length) {
      html += '<div class="msg warn">⚠ ' + plan.truncatedCells.length + ' cell(s) exceeded Excel\'s 32,767-character limit and were trimmed:<ul class="trunc-list">' +
        plan.truncatedCells.slice(0, 50).map(function (c) {
          return '<li><strong>' + esc(c.key) + '</strong> → ' + esc(c.column) + ' (row ' + c.excelRow + ') — ' + c.chars.toLocaleString() + ' chars</li>';
        }).join('') + (plan.truncatedCells.length > 50 ? '<li>…and ' + (plan.truncatedCells.length - 50) + ' more</li>' : '') + '</ul></div>';
    }
    if (plan.updates.length) html += '<button type="button" class="btn ghost" id="logBtn">Download change log (.csv)</button>';
    $('processMsg').innerHTML = html;
    var lb = $('logBtn');
    if (lb) lb.addEventListener('click', function () {
      download(new Blob([MergeCore.changeLogCsv(plan, report.skipped)], { type: 'text/csv;charset=utf-8' }), baseOf(outName) + '_changelog.csv');
    });
  }

  function stat(n, label, cls) {
    return '<div class="stat"><div class="n ' + cls + '">' + n + '</div><div class="l">' + label + '</div></div>';
  }

  /* ---------- wiring ---------- */
  ['orig', 'shared'].forEach(function (side) {
    var zone = $(side + 'Zone'), input = $(side + 'Input');
    zone.addEventListener('click', function () { input.click(); });
    zone.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', function () { loadFile(side, input.files[0]); input.value = ''; });
    zone.addEventListener('dragover', function (e) { e.preventDefault(); zone.classList.add('over'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('over'); });
    zone.addEventListener('drop', function (e) { e.preventDefault(); zone.classList.remove('over'); loadFile(side, e.dataTransfer.files[0]); });
    $(side + 'Sheet').addEventListener('change', function () { refreshHeaders(side); });
    $(side + 'Header').addEventListener('change', function () { refreshHeaders(side); });
  });
  function syncCombine() { $('combineSepWrap').classList.toggle('off', !$('combineDup').checked); }
  $('combineDup').addEventListener('change', function () {
    syncCombine();
    if ($('previewOut').style.display === 'block') buildPreview();
  });
  syncCombine();
  $('previewBtn').addEventListener('click', buildPreview);
  $('addRuleBtn').addEventListener('click', function () { addRule(); });
  $('autoMapBtn').addEventListener('click', autoMap);
  $('saveCfgBtn').addEventListener('click', saveConfig);
  $('loadCfgBtn').addEventListener('click', function () { $('cfgInput').click(); });
  $('cfgInput').addEventListener('change', function () { if (this.files[0]) loadConfig(this.files[0]); this.value = ''; });
  $('processBtn').addEventListener('click', process);
})();
