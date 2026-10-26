// OnlyOffice DSH AI 插件 v1
// 本地 DSH 服务（默认 127.0.0.1:3080）的第二个 AI 入口，与 AI 助手（plugin-ai）并存。
// 双路径：① 面板内直聊（POST /api/dsh/chat → 按结构化 actions 精确写入）
//        ② 队列轮询（POST /api/dsh/pending，instruction 支持 JSON actions 数组）
// 写入能力：光标处 / 文末(含真分页) / 锚点段落前后插入 / 替换锚点段落（复用 plugin-ai 执行器）
(function () {
  const qs = new URLSearchParams(location.search)
  let session = qs.get('session') || ''

  const dbgLines = []
  function dbg(msg) {
    dbgLines.push(new Date().toISOString().slice(11, 19) + ' ' + msg)
    if (dbgLines.length > 10) dbgLines.shift()
    const el = document.getElementById('dbg')
    if (el) el.textContent = dbgLines.join('\n')
    console.log('[dsh-plugin] ' + msg)
    try {
      fetch(apiBase + '/api/dsh/log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ source: 'plugin', msg: String(msg).slice(0, 480) }),
      }).catch(() => {})
    } catch {}
  }
  window.onerror = function (m, s, l) {
    dbg('onerror: ' + m + ' @' + (s || '').split('/').pop() + ':' + l)
  }

  function setStatus(text, cls) {
    // topbar 的 #status 常显（轮询计数/写文档状态直接可见）
    const tag = document.getElementById('status')
    if (tag) {
      tag.style.display = ''
      tag.textContent = text
      tag.className = cls || 'idle'
    }
  }

  // apiBase：config.js 可配置（独立安装到 DS 时填后端地址）；留空取同源（内置注入式）
  const apiBase = window.OO_API_BASE || window.location.origin
  dbg('plugin.js v11 (dsh) 加载')

  // 兜底：iframe 可能丢 query 或后端尚未登记会话，从后端查活跃 session（15s 节流重试）
  let lastSessionTry = 0
  async function ensureSession() {
    if (session) return
    const now = Date.now()
    if (now - lastSessionTry < 15000) return
    lastSessionTry = now
    try {
      const r = await fetch(apiBase + '/api/dsh/session', { cache: 'no-store' })
      const j = await r.json()
      if (j.session) {
        session = j.session
        dbg('session=' + session + '（后端兜底）')
      }
    } catch (e) {
      dbg('获取 session 失败: ' + String(e.message || e).slice(0, 60))
    }
  }

  // 轻量剥掉常见 markdown 符号
  function plain(t) {
    return String(t)
      .replace(/^#{1,6}\s+/gm, '')
      .replace(/\*\*([^*]+)\*\*/g, '$1')
      .replace(/`([^`]+)`/g, '$1')
  }

  let busy = false
  let inserted = 0
  let pollCount = 0
  let started = false

  function pasteText(text, cb) {
    window.Asc.plugin.executeMethod('PasteText', [text], function (ret) {
      dbg('PasteText 返回 ' + JSON.stringify(ret))
      if (ret === null || ret === undefined) {
        window.Asc.plugin.executeMethod('InputText', [text], function (r2) {
          dbg('InputText 返回 ' + JSON.stringify(r2))
          cb()
        })
      } else cb()
    })
  }

  // 真分页 + 文末写入：全部走文档模型（doc.Push），完全不依赖光标位置
  // parts 为 [[PAGE_BREAK]] 切分后的片段；空首段 = 先分页（第N页语义）
  function pushEnd(content, cb) {
    const parts = String(content).split(/\[\[PAGE_BREAK\]\]/gi)
    const implSrc =
      '(function(parts){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      'var br0 = (parts[0].trim()==="");' +
      'for (var i=0;i<parts.length;i++){' +
      ' var t = parts[i].trim();' +
      ' if (!t) continue;' +
      ' if ((i>0) || br0){ var pb=Api.CreateParagraph(); pb.AddPageBreak(); doc.Push(pb); }' +
      'var lines = t.split("\\n");' +
      ' for (var k=0;k<lines.length;k++){ var p=Api.CreateParagraph(); if(lines[k]) p.AddText(lines[k]); doc.Push(p); }' +
      '}' +
      'var D=doc.Document||doc; if(D&&D.Recalculate)D.Recalculate();' +
      'return true;' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(parts) + ')'
    const fn = new Function('return ' + implSrc + ';')
    // 第3参 recalc=true 触发画布重绘（同 docFormat）
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('pushEnd 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 真表格：Api.CreateTable 追加到文档末尾，data 二维数组填内容
  function docTable(rows, cols, data, cb) {
    const implSrc =
      '(function(rows, cols, data){' +
      'try{' +
      'rows = Math.max(1, Math.min(50, rows|0)); cols = Math.max(1, Math.min(20, cols|0));' +
      'var doc = Api.GetDocument();' +
      'var t = Api.CreateTable(cols, rows);' +
      't.SetTableBorderAll("single", 4, 0, 0, 0, 0);' +
      'if (data && data.length){' +
      ' for (var r=0; r<rows && r<data.length; r++){' +
      '  var row = data[r]; if (!row) continue;' +
      '  for (var c=0; c<cols && c<row.length; c++){' +
      '   var v = row[c]; if (v===null || v===undefined || v==="") continue;' +
      '   var cell = t.GetCell(r, c); if (!cell) continue;' +
      '   var cc = cell.GetContent();' +
      '   var p = !cc ? null : (cc.length ? cc[0] : (cc.GetElement ? cc.GetElement(0) : null));' +
      '   if (p && p.AddText) p.AddText(String(v));' +
      '  }' +
      ' }' +
      '}' +
      'doc.Push(t);' +
      'doc.Push(Api.CreateParagraph());' +
      'var D=doc.Document||doc; if(D&&D.Recalculate)D.Recalculate();' +
      'return "ok:"+rows+"x"+cols;' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(rows|0 || 3) + ',' + JSON.stringify(cols|0 || 3) + ',' + JSON.stringify(data || null) + ')'
    const fn = new Function('return ' + implSrc + ';')
    // 第3参 recalc=true 触发画布重绘（同 pushEnd）
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      console.log('[oo] docTable 返回:', JSON.stringify(ret))
      dbg('docTable 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 超链接：段落级 AddHyperlink，追加到文档末尾
  function docLink(url, text, cb) {
    const implSrc =
      '(function(url, text){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      'var p = Api.CreateParagraph();' +
      'p.AddHyperlink({ sUrl: url, sText: text || url, sToolTip: url });' +
      'doc.Push(p);' +
      'var D=doc.Document||doc; if(D&&D.Recalculate)D.Recalculate();' +
      'return "ok";' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(String(url)) + ',' + JSON.stringify(String(text || '')) + ')'
    const fn = new Function('return ' + implSrc + ';')
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('docLink 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 图片：URL 或 base64，宽高毫米（默认 120x90），追加到文档末尾
  function docImage(src, w, h, cb) {
    w = Math.min(300, Math.max(10, parseInt(w, 10) || 120))
    h = Math.min(300, Math.max(10, parseInt(h, 10) || 90))
    const implSrc =
      '(function(src, w, h){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      // CreateImage 宽高是 EMU 单位（1mm=36000），入参为毫米须换算，否则画布上仅 0.003mm 不可见
      'var img = Api.CreateImage(src, w*36000, h*36000);' +
      'var p = Api.CreateParagraph();' +
      'p.AddDrawing(img);' +
      'doc.Push(p);' +
      'var D=doc.Document||doc; if(D&&D.Recalculate)D.Recalculate();' +
      'return "ok:"+w+"x"+h;' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(String(src)) + ',' + JSON.stringify(w) + ',' + JSON.stringify(h) + ')'
    const fn = new Function('return ' + implSrc + ';')
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('docImage 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 图表：Api.CreateChart 追加到文档末尾；series=[{name,data[]}]，categories=[...]
  function docChart(chartType, categories, series, w, h, cb) {
    const OK = ['bar', 'column', 'line', 'pie', 'area', 'scatter', 'doughnut']
    if (OK.indexOf(chartType) < 0) chartType = 'bar'
    // DS 内部类型名映射：此版本 sdkjs 无 "line"/"column"，须用 lineNormal/bar（column 即竖向柱状）
    const TYPEMAP = { line: 'lineNormal', column: 'bar' }
    if (TYPEMAP[chartType]) chartType = TYPEMAP[chartType]
    w = Math.min(400, Math.max(40, parseInt(w, 10) || 120))
    h = Math.min(400, Math.max(40, parseInt(h, 10) || 80))
    const data = series.slice(0, 10).map(function (s) { return (s && s.data) || [] })
    const names = series.slice(0, 10).map(function (s, i) { return (s && s.name) || '系列' + (i + 1) })
    const implSrc =
      '(function(type, data, names, cats, w, h){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      // CreateChart 宽高是 EMU 单位（1mm=36000），同 CreateImage
      'var chart = Api.CreateChart(type, data, names, cats, w*36000, h*36000, 0);' +
      'var p = Api.CreateParagraph();' +
      'p.AddDrawing(chart);' +
      'doc.Push(p);' +
      'var D=doc.Document||doc; if(D&&D.Recalculate)D.Recalculate();' +
      'return "ok:"+type;' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(chartType) + ',' + JSON.stringify(data) + ',' + JSON.stringify(names) + ',' + JSON.stringify(categories || []) + ',' + JSON.stringify(w) + ',' + JSON.stringify(h) + ')'
    const fn = new Function('return ' + implSrc + ';')
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('docChart 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 合并：把某表格 row 行 colStart~colEnd 横向合并（anchor 定位表格，缺省第一个）
  function docMerge(anchor, row, colStart, colEnd, cb) {
    const implSrc =
      '(function(anchor, row, c0, c1){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      'var tables = [];' +
      'for (var i=0;i<doc.GetElementsCount();i++){' +
      ' var el = doc.GetElement(i);' +
      ' if (el && typeof el.GetClassType==="function" && el.GetClassType()==="table") tables.push(el);' +
      '}' +
      'if (!tables.length) return "err:文档里没有表格";' +
      'var t = tables[0];' +
      'if (anchor){' +
      ' outer: for (var ti=0;ti<tables.length;ti++){' +
      '  var tt = tables[ti];' +
      '  for (var r=0;r<tt.GetRowsCount();r++){' +
      '   var rr = tt.GetRow(r);' +
      '   for (var c=0;c<rr.GetCellsCount();c++){' +
      '    var cc = rr.GetCell(c);' +
      '    if (cc && cc.GetText && String(cc.GetText()).indexOf(anchor)>=0){ t = tt; break outer; }' +
      '   }' +
      '  }' +
      ' }' +
      '}' +
      'row = Math.max(0, Math.min(t.GetRowsCount()-1, row|0));' +
      'var rr = t.GetRow(row);' +
      'var n = rr.GetCellsCount();' +
      'c0 = Math.max(0, Math.min(n-2, c0|0)); c1 = Math.max(c0+1, Math.min(n-1, c1|0));' +
      'var cells = [];' +
      'for (var c=c0;c<=c1;c++) cells.push(rr.GetCell(c));' +
      'if (cells.length<2) return "err:合并范围少于两格";' +
      't.MergeCells(cells);' +
      'var D=doc.Document||doc; if(D&&D.Recalculate)D.Recalculate();' +
      'return "ok:row"+row+" c"+c0+"-"+c1;' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(String(anchor || '')) + ',' + JSON.stringify(parseInt(row, 10) || 0) + ',' + JSON.stringify(parseInt(colStart, 10) || 0) + ',' + JSON.stringify(parseInt(colEnd, 10) || 0) + ')'
    const fn = new Function('return ' + implSrc + ';')
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('docMerge 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 格式化：对含 anchor 文字的段落应用完整格式集，返回诊断信息（哪个属性真正应用了几个 run）
  function docFormat(anchor, o, cb) {
    o = o || {}
    const implSrc =
      '(function(anchor, f){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      'var paras = doc.GetAllParagraphs();' +
      'var col = null;' +
      'if (f.color && /^[0-9A-Fa-f]{6}$/.test(f.color)) col = { r: parseInt(f.color.substr(0,2),16), g: parseInt(f.color.substr(2,2),16), b: parseInt(f.color.substr(4,2),16) };' +
      'var hit = -1, nRuns = 0, cnt = {font:0,size:0,bold:0,color:0,align:0,italic:0,underline:0,strikeout:0,highlight:0};' +
      'var before = null, after = null, RCret = "skipped";' +
      'var RC=function(){var D=doc.Document||doc; if(D&&D.Recalculate){try{D.Recalculate();RCret="ok"}catch(e){RCret="err:"+e.message}}else{RCret="no-Recalc(doc="+(typeof doc)+",D="+(typeof D)+")"}};' +
      'for (var i=0;i<paras.length;i++){' +
      ' if (paras[i].GetText().indexOf(anchor) < 0) continue;' +
      ' hit = i;' +
      ' if (f.align && paras[i].SetAlign) { paras[i].SetAlign(f.align); cnt.align++; }' +
      ' var n = paras[i].GetElementsCount();' +
      ' for (var j=0;j<n;j++){ var el = paras[i].GetElement(j); if (!el) continue;' +
      '  nRuns++;' +
      '  if (nRuns === 1) {' +
      '   try { var pr0 = el.GetTextPr ? el.GetTextPr() : null; before = pr0 ? { size: (pr0.GetFontSize ? pr0.GetFontSize() : "?"), font: (pr0.GetFontName ? pr0.GetFontName() : "?") } : "no-TextPr"; } catch(e) { before = "readback-err:" + e.message; }' +
      '  }' +
      '  if (f.bold !== null && el.SetBold) { el.SetBold(f.bold); cnt.bold++; }' +
      '  if (f.italic !== null && el.SetItalic) { el.SetItalic(f.italic); cnt.italic++; }' +
      '  if (f.underline !== null && el.SetUnderline) { el.SetUnderline(f.underline); cnt.underline++; }' +
      '  if (f.strikeout !== null && el.SetStrikeout) { el.SetStrikeout(f.strikeout); cnt.strikeout++; }' +
      '  if (f.fontSize && el.SetFontSize) { el.SetFontSize(f.fontSize); cnt.size++; }' +
      '  if (f.fontName && el.SetFontName) { el.SetFontName(f.fontName); cnt.font++; }' +
      '  if (col && el.SetColor) { el.SetColor(col.r, col.g, col.b); cnt.color++; }' +
      '  if (f.highlight && el.SetHighlight) { el.SetHighlight(f.highlight); cnt.highlight++; }' +
      '  if (nRuns === 1) {' +
      '   try { var pr1 = el.GetTextPr ? el.GetTextPr() : null; after = pr1 ? { size: (pr1.GetFontSize ? pr1.GetFontSize() : "?"), font: (pr1.GetFontName ? pr1.GetFontName() : "?") } : "no-TextPr"; } catch(e) { after = "readback-err:" + e.message; }' +
      '  }' +
      ' }' +
      ' RC();' +
      ' break;' +
      '}' +
      'return hit < 0 ? "anchor-not-found" : JSON.stringify({para:hit, runs:nRuns, applied:cnt, before:before, after:after, rc:RCret});' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(anchor) + ',' + JSON.stringify({
        bold: o.bold === undefined ? null : !!o.bold,
        italic: o.italic === undefined ? null : !!o.italic,
        underline: o.underline === undefined ? null : !!o.underline,
        strikeout: o.strikeout === undefined ? null : !!o.strikeout,
        fontSize: o.fontSize || null,
        fontName: o.fontName || null,
        color: o.color || null,
        highlight: o.highlight || null,
        align: o.align || null,
      }) + ')'
    const fn = new Function('return ' + implSrc + ';')
    // 第3参 recalc=true：触发画布重绘（实测 false 时模型已改但画布不刷新）；
    // 第2参必须 false（true 发 close 命令，回调挂起且插件轮询死亡）
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('docFormat 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 锚点操作：在含 anchor 文字的段落 前后插入新段落 / 替换该段落
  // callCommand 序列化函数源码，参数用 JSON 字面量烘焙进函数体
  function docOp(kind, anchor, content, position, cb) {
    const implSrc =
      '(function(kind, anchor, content, position){' +
      'try{' +
      'var doc = Api.GetDocument();' +
      'var paras = doc.GetAllParagraphs();' +
      'var lines = content.split("\\n");' +
      'var RC=function(){var D=doc.Document||doc; if(D&&D.Recalculate)try{D.Recalculate()}catch(e){}};' +
      'for (var i=0;i<paras.length;i++){' +
      ' var t = paras[i].GetText();' +
      ' if (t.indexOf(anchor) < 0) continue;' +
      ' var made=[];' +
      ' for (var k=0;k<lines.length;k++){' +
      '  if (!lines[k].trim() && lines.length>1) continue;' +
      '  var np = Api.CreateParagraph(); np.AddText(lines[k]); made.push(np);' +
      ' }' +
      ' if (!made.length) made.push(Api.CreateParagraph());' +
      ' if (kind==="replace"){ made[0].InsertParagraph ? paras[i].InsertParagraph(made[0],"before") : doc.InsertParagraph(made[0],paras[i],"before");' +
      '   for (var m=1;m<made.length;m++) made[m-1].InsertParagraph(made[m],"after");' +
      '   paras[i].Delete(); RC(); return true; }' +
      ' if (position==="before"){ for (var m=made.length-1;m>=0;m--) paras[i].InsertParagraph(made[m],"before"); RC(); return true; }' +
      ' var prev=paras[i];' +
      ' for (var m=0;m<made.length;m++){ prev.InsertParagraph(made[m],"after"); prev=made[m]; }' +
      ' RC(); return true; }' +
      'return "anchor-not-found";' +
      '}catch(e){ return "err:"+e.message; }' +
      '})(' + JSON.stringify(kind) + ',' + JSON.stringify(anchor) + ',' + JSON.stringify(content) + ',' + JSON.stringify(position) + ')'
    const fn = new Function('return ' + implSrc + ';')
    // 第3参 recalc=true 触发画布重绘（同 docFormat）
    window.Asc.plugin.callCommand(fn, false, true, function (ret) {
      dbg('docOp(' + kind + ') 返回 ' + JSON.stringify(ret))
      cb(ret)
    })
  }

  // 抽取文档正文（选址上下文，最多 200 段）
  function getDocText(cb) {
    window.Asc.plugin.callCommand(function () {
      var doc = Api.GetDocument()
      var paras = doc.GetAllParagraphs()
      var out = []
      for (var i = 0; i < paras.length && i < 200; i++) {
        var t = paras[i].GetText()
        out.push(t.length > 200 ? t.slice(0, 200) : t)
      }
      return out.join('\n')
    }, false, false, function (ret) {
      cb(typeof ret === 'string' ? ret : '')
    })
  }

  function syncDocContext() {
    getDocText(function (text) {
      if (!text) return
      fetch(apiBase + '/api/dsh/context', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session: session, text: text }),
      }).catch(function () {})
    })
  }

  // === 结构化动作执行 ===
  function handleActions(actions, done) {
    const list = (Array.isArray(actions) ? actions : []).filter(function (a) { return a && a.type })
    if (!list.length) { done(); return }
    let idx = 0
    const next = function () {
      if (idx >= list.length) {
        inserted++
        setStatus('✓ 已插' + inserted + '条', 'done')
        syncDocContext()
        done(); return
      }
      const a = list[idx++]
      // 前后 trim：AI 内容里首尾的空行/空格不写进文档
      const content = plain(a.content || '').trim()
      if (a.type === 'format') {
        if (!a.anchor) { dbg('format 缺 anchor，跳过'); next(); return }
        dbg('format 载荷: ' + JSON.stringify(a))
        docFormat(a.anchor, a, function (ret) { dbg('docFormat 返回 ' + ret); next() })
        return
      }
      if (a.type === 'table') {
        docTable(parseInt(a.rows, 10) || 3, parseInt(a.cols, 10) || 3, a.data, function () { next() })
        return
      }
      if (a.type === 'link') {
        if (!a.url) { dbg('link 缺 url，跳过'); next(); return }
        docLink(String(a.url), a.text ? plain(a.text) : '', function () { next() })
        return
      }
      if (a.type === 'image') {
        if (!a.src) { dbg('image 缺 src，跳过'); next(); return }
        docImage(a.src, a.width, a.height, function () { next() })
        return
      }
      if (a.type === 'chart') {
        if (!Array.isArray(a.series) || !a.series.length) { dbg('chart 缺 series，跳过'); next(); return }
        docChart(a.chartType, Array.isArray(a.categories) ? a.categories : [], a.series, a.width, a.height, function () { next() })
        return
      }
      if (a.type === 'merge') {
        docMerge(a.anchor ? plain(a.anchor) : '', a.row, a.colStart, a.colEnd, function () { next() })
        return
      }
      if (!content) { next(); return }
      if (a.type === 'replace') {
        if (!a.anchor) { dbg('replace 缺 anchor，跳过'); next(); return }
        docOp('replace', a.anchor, content, 'after', function () { next() })
      } else if (a.target === 'after' || a.target === 'before') {
        if (!a.anchor) { dbg('insert 缺 anchor，跳过'); next(); return }
        docOp('insert', a.anchor, content, a.target, function () { next() })
      } else if (a.target === 'cursor') {
        pasteText(content + '\n', function () { next() })
      } else {
        // end（默认）：文档模型追加，不依赖光标
        pushEnd(content, function () { next() })
      }
    }
    next()
  }

  // instruction 兼容两种格式：JSON actions 数组 / 纯文本
  function handleInstruction(instruction, done) {
    const s = String(instruction || '').trim()
    if (s[0] === '[') {
      try { handleActions(JSON.parse(s), done); return } catch (e) { dbg('actions JSON 解析失败: ' + e.message) }
    }
    pushEnd(plain(s).trim(), done)
  }

  async function poll() {
    if (busy || !window.Asc?.plugin) return
    if (!session) {
      await ensureSession()
      if (!session) { setStatus('无session', 'err'); return }
    }
    busy = true
    try {
      pollCount++
      setStatus('…轮询' + pollCount, 'busy')
      const r = await fetch(apiBase + '/api/dsh/pending?session=' + encodeURIComponent(session), { cache: 'no-store' })
      if (!r.ok) { setStatus('后端' + r.status, 'err'); return }
      const j = await r.json()
      console.log('[oo] poll 拿到:', j.instruction ? String(j.instruction).slice(0, 60) : '(空)')
      if (!j.instruction) return
      handleInstruction(j.instruction, function () {})
    } catch (e) {
      setStatus('⚠️ ' + String(e.message || e).slice(0, 14), 'err')
      dbg('poll 异常: ' + String(e.message || e).slice(0, 60))
    } finally {
      busy = false
    }
  }

  async function startPolling() {
    if (started) return
    started = true
    setStatus('启动中', 'busy')
    setInterval(poll, 1500)
    poll()
    syncDocContext()
  }

  // OnlyOffice 插件生命周期钩子
  window.Asc.plugin.init = function () {
    dbg('init 触发')
    startPolling()
  }
  // 看门狗：某些版本 isVisual 插件不回调 init，3s 后强启
  setTimeout(function () {
    if (!started) {
      dbg('init 3s 未触发，看门狗强启')
      startPolling()
    }
  }, 3000)

  window.Asc.plugin.button = function (id) {
    dbg('toolbar button 点击 id=' + id)
    // 面板右上角 X（id=-1）：插件定义了 button 回调后，DS 不再自动关闭，须自己触发
    if (id === -1) this.executeCommand('close', '')
  }

  // ===== 原生对话 UI：面板内直接聊天，DSH 回复写入文档 =====
  const box = document.getElementById('msgs')
  const inp = document.getElementById('inp')
  const sendBtn = document.getElementById('send')

  // 跟随 OnlyOffice 编辑器主题（dark/light）
  function applyTheme(theme) {
    const dark = (theme && theme.type === 'dark') ||
      document.documentElement.classList.contains('theme-dark') ||
      document.body.classList.contains('theme-dark')
    document.body.classList.toggle('dark', !!dark)
  }
  applyTheme()
  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.onThemeChanged = function (theme) { applyTheme(theme) }
  }

  function bubble(role, text) {
    const div = document.createElement('div')
    div.className = 'msg ' + role
    const b = document.createElement('div')
    b.className = 'bubble'
    b.textContent = text
    div.appendChild(b)
    box.appendChild(div)
    box.scrollTop = box.scrollHeight
    return b
  }

  function typingBubble() {
    const div = document.createElement('div')
    div.className = 'msg ai'
    // DSH web 风格思考块：单行「思考 · 最新进度」，点击展开全部；回复到达后整体移除
    const think = document.createElement('div')
    think.className = 'think'
    const thinkLine = document.createElement('div')
    thinkLine.className = 'think-line'
    const spin = document.createElement('span')
    spin.className = 'spin'
    const label = document.createElement('span')
    label.textContent = '思考 · '
    const cur = document.createElement('span')
    cur.className = 'think-cur'
    cur.textContent = '连接 DSH…'
    thinkLine.appendChild(spin)
    thinkLine.appendChild(label)
    thinkLine.appendChild(cur)
    const thinkAll = document.createElement('div')
    thinkAll.className = 'think-all'
    thinkAll.style.display = 'none'
    thinkLine.addEventListener('click', function () {
      thinkAll.style.display = thinkAll.style.display === 'none' ? 'block' : 'none'
    })
    think.appendChild(thinkLine)
    think.appendChild(thinkAll)
    const b = document.createElement('div')
    b.className = 'bubble typing'
    b.style.display = 'none'
    b.innerHTML = '<span></span><span></span><span></span>'
    div.appendChild(think)
    div.appendChild(b)
    box.appendChild(div)
    box.scrollTop = box.scrollHeight
    return b
  }

  // 复制文本（iframe 里 navigator.clipboard 可能受限，带 execCommand 兜底）
  function copyText(text, ok) {
    const done = function () { ok() }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done).catch(function () { fallbackCopy(text, done) })
    } else fallbackCopy(text, done)
  }
  function fallbackCopy(text, done) {
    const ta = document.createElement('textarea')
    ta.value = text
    ta.style.cssText = 'position:fixed;left:-999px;top:0'
    document.body.appendChild(ta)
    ta.select()
    try { document.execCommand('copy') } catch {}
    document.body.removeChild(ta)
    done()
  }

  // 回复下方元信息行：左（复制+点赞）右（耗时）
  let lastUserMsg = ''
  let lastActions = null
  function addMeta(msgDiv, replyText, secs, actions, memHit) {
    const meta = document.createElement('div')
    meta.className = 'meta'
    const left = document.createElement('span')
    left.className = 'left'

    const copy = document.createElement('span')
    copy.className = 'copy'
    copy.title = '复制回复'
    copy.textContent = '⧉'
    copy.addEventListener('click', function () {
      copyText(replyText, function () {
        copy.textContent = '✓'
        setTimeout(function () { copy.textContent = '⧉' }, 1500)
      })
    })

    const like = document.createElement('span')
    like.className = 'like'
    like.title = '点赞并记住这次操作'
    // 线条点赞图标，点击后变实心填充
    like.innerHTML =
      '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M14 9V5a3 3 0 0 0-3-3l-4 9v11h11.28a2 2 0 0 0 2-1.7l1.38-9a2 2 0 0 0-2-2.3zM7 22H4a2 2 0 0 1-2-2v-7a2 2 0 0 1 2-2h3"/></svg>'
    like.addEventListener('click', function () {
      if (like.classList.contains('liked')) return
      like.classList.add('liked')
      like.title = '已记住'
      like.querySelector('path').setAttribute('d', 'M14 9V5a3 3 0 0 0-3-3l-4 9v11h11.28a2 2 0 0 0 2-1.7l1.38-9a2 2 0 0 0-2-2.3zM7 22H4a2 2 0 0 1-2-2v-7a2 2 0 0 1 2-2h3')
      like.querySelector('svg').setAttribute('fill', 'currentColor')
      fetch(apiBase + '/api/dsh/like', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: lastUserMsg, summary: replyText, actions: actions || null }),
      }).catch(function () {})
    })

    left.appendChild(copy)
    left.appendChild(like)
    const time = document.createElement('span')
    time.textContent = (memHit ? '⚡记忆 ' : '') + secs + 's'
    meta.appendChild(left)
    meta.appendChild(time)
    msgDiv.appendChild(meta)
    box.scrollTop = box.scrollHeight
  }

  async function send() {
    const text = inp.value.trim()
    if (!text || busy) return
    inp.value = ''
    inp.style.height = ''  // 清空后重置回 1 行高
    sendBtn.textContent = '停止'  // 发送中按钮变停止（点击走 stopChat，见下方 click 监听）
    busy = true
    bubble('user', text)
    setStatus('思考中', 'busy')
    const typing = typingBubble()
    const t0 = Date.now()
    // DSH headless 单发模式无中间流，长任务用秒数计时给出「活着」反馈
    const thinkTimer = setInterval(function () {
      setStatus('思考中 ' + Math.floor((Date.now() - t0) / 1000) + 's', 'busy')
    }, 500)
    try {
      // 实时抽取文档正文作为选址上下文
      const docText = await new Promise(function (res) { getDocText(res) })
      // 思考过程：轮询喂给 typingBubble 里的 think 块（单行最新/点击展开全部）
      const reqId = 'r' + Date.now().toString(36)
      const think = typing.previousElementSibling
      const thinkCur = think ? think.querySelector('.think-cur') : null
      const thinkAll = think ? think.querySelector('.think-all') : null
      const progTimer = setInterval(function () {
        fetch(apiBase + '/api/dsh/chat-progress?reqId=' + reqId)
          .then(function (r) { return r.json() })
          .then(function (j) {
            const lines = j.lines || []
            if (!lines.length || !thinkCur) return
            thinkCur.textContent = lines[lines.length - 1]
            if (thinkAll) {
              thinkAll.textContent = lines.join('\n')
              if (thinkAll.style.display !== 'none') thinkAll.scrollTop = thinkAll.scrollHeight
            }
          })
          .catch(function () {})
      }, 1000)
      let r
      try {
        r = await fetch(apiBase + '/api/dsh/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ message: text, docUrl: session, docContext: docText, reqId: reqId }),
        })
      } finally {
        clearInterval(progTimer)
        if (think) think.parentNode.removeChild(think)  // 结束后思考块让位给正式回复
        typing.style.display = ''
      }
      const j = await r.json()
      typing.parentNode.className = 'msg ai'
      typing.className = 'bubble'
      if (j.session && j.session !== session) { session = j.session; dbg('session=' + session) }  // 记住会话号，下条消息续会话（DSH 记忆关键）
      if (!r.ok || !j.ok) {
        const errMsg = j.error || 'HTTP ' + r.status
        // 主动停止：中性提示，不当失败处理
        if (errMsg.indexOf('停止') >= 0) {
          typing.parentNode.className = 'msg ai'
          typing.textContent = '⏹ 已停止'
          setStatus('已停止', 'idle')
          lastUserMsg = ''
          lastActions = null
          addMeta(typing.parentNode, '', ((Date.now() - t0) / 1000).toFixed(1), null, false)
          return
        }
        // DSH 调用失败：明确显示错误（CLI 退出码/网络原因），便于排查本机服务
        typing.textContent = '⚠️ DSH 调用失败：' + errMsg + '\n（检查本机 DSH 服务/登录态）'
        setStatus('调用失败', 'err')
        lastUserMsg = ''
        lastActions = null
        addMeta(typing.parentNode, '', ((Date.now() - t0) / 1000).toFixed(1), null, false)
        return
      }
      typing.textContent = j.message || '(无回复)'
      lastUserMsg = text
      lastActions = j.actions || null
      addMeta(typing.parentNode, j.message || '', ((Date.now() - t0) / 1000).toFixed(1), lastActions, j.memoryHit)
      if (j.actions && j.actions.length) {
        setStatus('写入中', 'busy')
        handleActions(j.actions, function () {})
      } else {
        // 无结构化动作：纯文本回复直接落到文档末尾（纯文本降级，同 pushEnd 通道）
        setStatus('写入中', 'busy')
        handleActions([{ type: 'text', content: j.message || '' }], function () {})
      }
    } catch (e) {
      typing.parentNode.className = 'msg err'
      typing.className = 'bubble'
      typing.textContent = '后端不可达：' + String(e.message || e) + '（检查 DSH 服务是否启动）'
      setStatus('API错误', 'err')
    } finally {
      clearInterval(thinkTimer)
      sendBtn.textContent = '发送'
      busy = false
      inp.value = ''
      inp.style.height = ''  // 恢复 1 行高
      inp.focus()
    }
  }

  const form = document.getElementById('form')
  if (form) {
    form.addEventListener('submit', function (e) { e.preventDefault(); send() })
    // 发送中按钮即「停止」：终止后端 dsh 子进程树；submit 事件里 send() 因 busy 直接 return，不会误发新消息
    sendBtn.addEventListener('click', function () {
      if (busy) fetch(apiBase + '/api/dsh/stop', { method: 'POST' }).catch(function () {})
    })
    // textarea：Enter 换行，Ctrl/Cmd+Enter 发送；输入时自动长高（上限约6行）
    inp.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault()
        send()
      }
    })
    inp.addEventListener('input', function () {
      inp.style.height = 'auto'
      inp.style.height = Math.min(inp.scrollHeight, 140) + 'px'
    })
    setTimeout(function () { inp.focus() }, 500)
  }

  window.Asc.plugin.onExternalPluginMessage = function (data) {
    if (data && data.instruction) {
      handleInstruction(data.instruction, function () {})
    }
  }
})()
