/* ============================================================
 * PPT Studio — PPTX（OOXML）生成器
 * 输入 layout.js 产出的绘制指令，输出一个完整可用的 .pptx 字节流。
 * 零依赖：包结构、关系、内容类型、母版/版式/主题、表格、备注页
 * 全部手写。经 python-pptx 与 LibreOffice 渲染验证。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});
  var EMU = 914400;
  function E(inch) { return String(Math.round((+inch || 0) * EMU)); }

  /* ---------- 基础工具 ---------- */
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  function normColor(c) {
    if (c == null) return { c: '000000', a: 1 };
    if (typeof c === 'object') {
      return { c: normColor(c.c).c, a: (c.a == null ? 1 : +c.a) };
    }
    var s = String(c).replace(/^#/, '').toUpperCase();
    if (s.length === 3) s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    if (!/^[0-9A-F]{6}$/.test(s)) s = '000000';
    return { c: s, a: 1 };
  }
  function fillXml(fill, alphaOverride) {
    if (fill == null) return '<a:noFill/>';
    var nc = normColor(fill);
    var a = (alphaOverride == null ? nc.a : alphaOverride);
    var inner = '<a:srgbClr val="' + nc.c + '">' +
      (a != null && a < 1 ? '<a:alpha val="' + Math.round(a * 100000) + '"/>' : '') +
      '</a:srgbClr>';
    return '<a:solidFill>' + inner + '</a:solidFill>';
  }
  function gradXml(from, to, angle) {
    return '<a:gradFill rotWithShape="1"><a:gsLst>' +
      '<a:gs pos="0">' + fillXml(from).replace(/^<a:solidFill>|<\/a:solidFill>$/g, '') + '</a:gs>' +
      '<a:gs pos="100000">' + fillXml(to).replace(/^<a:solidFill>|<\/a:solidFill>$/g, '') + '</a:gs>' +
      '</a:gsLst><a:lin ang="' + Math.round((angle || 0) * 60000) + '" scaled="0"/></a:gradFill>';
  }
  function bgXml(bg) {
    if (!bg) return '';
    var body = bg.grad
      ? gradXml(bg.grad[0], bg.grad[1], bg.angle || 0)
      : fillXml(bg.fill == null ? 'FFFFFF' : bg.fill);
    return '<p:bg><p:bgPr>' + body + '<a:effectLst/></p:bgPr></p:bg>';
  }
  var ANCHOR = { t: 't', ctr: 'ctr', b: 'b', c: 'ctr' };
  var ALIGN = { l: 'l', c: 'ctr', r: 'r', j: 'just' };

  var SPRTREE_HEAD =
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

  /* ---------- 运行 / 段落 ---------- */
  function runXml(r, th) {
    var sz = Math.round((r.sz || 18) * 100);
    var attrs = 'lang="zh-CN" sz="' + sz + '"';
    if (r.b) attrs += ' b="1"';
    if (r.i) attrs += ' i="1"';
    if (r.spc) attrs += ' spc="' + Math.round(r.spc) + '"';
    attrs += ' dirty="0"';
    var inner = '';
    if (r.color) inner += fillXml(r.color, r.alpha);
    var latin = r.font || th.latin, ea = r.font || th.ea;
    inner += '<a:latin typeface="' + esc(latin) + '"/><a:ea typeface="' + esc(ea) +
             '"/><a:cs typeface="' + esc(ea) + '"/>';
    return '<a:r><a:rPr ' + attrs + '>' + inner + '</a:rPr>' +
           '<a:t xml:space="preserve">' + esc(r.t) + '</a:t></a:r>';
  }

  function paraXml(p, th, emptySz) {
    var attrs = '';
    if (p.marL) attrs += ' marL="' + E(p.marL) + '"';
    if (p.indent) attrs += ' indent="' + E(p.indent) + '"';
    if (p.align && p.align !== 'l') attrs += ' algn="' + (ALIGN[p.align] || 'l') + '"';
    var inner = '';
    if (p.line) inner += '<a:lnSpc><a:spcPct val="' + Math.round(p.line * 100000) + '"/></a:lnSpc>';
    if (p.before) inner += '<a:spcBef><a:spcPts val="' + Math.round(p.before * 100) + '"/></a:spcBef>';
    if (p.after) inner += '<a:spcAft><a:spcPts val="' + Math.round(p.after * 100) + '"/></a:spcAft>';
    if (p.bullet) {
      inner += '<a:buClr><a:srgbClr val="' + normColor(p.bullet.color).c + '"/></a:buClr>' +
               '<a:buSzPct val="' + Math.round((p.bullet.size || 1) * 100000) + '"/>' +
               '<a:buFont typeface="Arial"/>' +
               '<a:buChar char="' + esc(p.bullet.char) + '"/>';
    } else {
      inner += '<a:buNone/>';
    }
    var runs = (p.runs || []).filter(function (r) { return r && r.t !== '' && r.t != null; })
      .map(function (r) { return runXml(r, th); }).join('');
    var pPr = '<a:pPr' + attrs + '>' + inner + '</a:pPr>';
    if (!runs) return '<a:p>' + pPr + '<a:endParaRPr lang="zh-CN" sz="' + (emptySz || 1800) + '"/></a:p>';
    return '<a:p>' + pPr + runs + '</a:p>';
  }

  /* ---------- 形状 ---------- */
  function spXml(id, op) {
    var geom;
    if (op.type === 'round') {
      var r = op.radius == null ? 0.1 : op.radius;
      var adj = Math.max(0, Math.min(50000, Math.round(r / Math.max(0.01, Math.min(op.w, op.h)) * 100000)));
      geom = '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ' + adj + '"/></a:avLst></a:prstGeom>';
    } else {
      geom = '<a:prstGeom prst="' + (op.type === 'ellipse' ? 'ellipse' : 'rect') + '"><a:avLst/></a:prstGeom>';
    }
    var ln = op.line
      ? '<a:ln w="' + Math.round((op.lineW || 1) * 12700) + '" cap="flat" cmpd="sng" algn="ctr">' + fillXml(op.line) + '</a:ln>'
      : '<a:ln><a:noFill/></a:ln>';
    return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="Shape ' + id + '"/>' +
      '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
      '<a:xfrm><a:off x="' + E(op.x) + '" y="' + E(op.y) + '"/><a:ext cx="' + E(op.w) + '" cy="' + E(op.h) + '"/></a:xfrm>' +
      geom + fillXml(op.fill, op.alpha) + ln + '</p:spPr></p:sp>';
  }

  function textSpXml(id, op, th) {
    var ins = op.ins || 0;
    var bodyPr = 'wrap="' + (op.wrap === false ? 'none' : 'square') + '" lIns="' + E(ins) + '" tIns="' + E(ins) +
      '" rIns="' + E(ins) + '" bIns="' + E(ins) + '" anchor="' + (ANCHOR[op.anchor] || 't') + '"';
    var paras = op.paras.map(function (p) { return paraXml(p, th); }).join('');
    return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="Text ' + id + '"/>' +
      '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>' +
      '<a:xfrm><a:off x="' + E(op.x) + '" y="' + E(op.y) + '"/><a:ext cx="' + E(op.w) + '" cy="' + E(op.h) + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
      '<p:txBody><a:bodyPr ' + bodyPr + '><a:noAutofit/></a:bodyPr><a:lstStyle/>' + paras + '</p:txBody></p:sp>';
  }

  function tableCellPara(text, sz, bold, color, align, th) {
    var attrs = 'lang="zh-CN" sz="' + Math.round(sz * 100) + '"' + (bold ? ' b="1"' : '') + ' dirty="0"';
    var inner = fillXml(color) + '<a:latin typeface="' + esc(th.latin) + '"/><a:ea typeface="' + esc(th.ea) +
      '"/><a:cs typeface="' + esc(th.ea) + '"/>';
    return '<a:p><a:pPr algn="' + (ALIGN[align] || 'l') + '"><a:buNone/></a:pPr>' +
      '<a:r><a:rPr ' + attrs + '>' + inner + '</a:rPr><a:t xml:space="preserve">' + esc(text) + '</a:t></a:r></a:p>';
  }

  function tableXml(id, op, th) {
    var rIdx;
    var ct = function () { return fillXml(op.border); };
    var grid = op.colW.map(function (w) { return '<a:gridCol w="' + E(w) + '"/>'; }).join('');
    var rows = op.rows.map(function (r, ri) {
      rIdx = ri;
      var fill = r.head ? op.headFill : (ri % 2 === 0 ? op.altFill : op.bodyFill);
      var cells = r.cells.map(function (txt, ci) {
        var align = op.aligns && op.aligns[ci] ? op.aligns[ci] : 'l';
        return '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>' +
          tableCellPara(txt, op.fontSize, !!r.head, r.head ? op.headText : op.text, align, th) +
          '</a:txBody><a:tcPr marL="' + E(0.10) + '" marR="' + E(0.10) + '" marT="' + E(0.03) +
          '" marB="' + E(0.03) + '" anchor="ctr">' +
          '<a:lnL w="6350" cap="flat" cmpd="sng" algn="ctr">' + ct() + '</a:lnL>' +
          '<a:lnR w="6350" cap="flat" cmpd="sng" algn="ctr">' + ct() + '</a:lnR>' +
          '<a:lnT w="6350" cap="flat" cmpd="sng" algn="ctr">' + ct() + '</a:lnT>' +
          '<a:lnB w="6350" cap="flat" cmpd="sng" algn="ctr">' + ct() + '</a:lnB>' +
          fillXml(fill) + '</a:tcPr></a:tc>';
      }).join('');
      return '<a:tr h="' + E(r.h) + '">' + cells + '</a:tr>';
    }).join('');
    var totalH = op.rows.reduce(function (a, r) { return a + r.h; }, 0);
    return '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="' + id + '" name="Table ' + id + '"/>' +
      '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="' + E(op.x) + '" y="' + E(op.y) + '"/><a:ext cx="' + E(op.w) + '" cy="' + E(totalH) + '"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
      '<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>' + grid + '</a:tblGrid>' + rows + '</a:tbl>' +
      '</a:graphicData></a:graphic></p:graphicFrame>';
  }

  /* ---------- 各部件 XML ---------- */
  var NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
           'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  var XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  function relsXml(list) {
    return XMLDECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      list.map(function (r) {
        return '<Relationship Id="' + r.id + '" Type="http://schemas.openxmlformats.org/' + r.type +
          '" Target="' + esc(r.target) + '"/>';
      }).join('') + '</Relationships>';
  }

  function themeXml(th, name) {
    var accents = [th.accent, th.accent2, th.muted, th.text, th.accent2, th.accent];
    var clr = '<a:clrScheme name="' + esc(name) + '">' +
      '<a:dk1><a:srgbClr val="' + normColor(th.text).c + '"/></a:dk1>' +
      '<a:lt1><a:srgbClr val="' + normColor(th.bg).c + '"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="' + normColor(th.text).c + '"/></a:dk2>' +
      '<a:lt2><a:srgbClr val="' + normColor(th.surface).c + '"/></a:lt2>' +
      accents.map(function (c, i) {
        return '<a:accent' + (i + 1) + '><a:srgbClr val="' + normColor(c).c + '"/></a:accent' + (i + 1) + '>';
      }).join('') +
      '<a:hlink><a:srgbClr val="' + normColor(th.accent).c + '"/></a:hlink>' +
      '<a:folHlink><a:srgbClr val="' + normColor(th.accent2).c + '"/></a:folHlink></a:clrScheme>';
    var font = function (tag) {
      return '<a:' + tag + '><a:latin typeface="' + esc(th.latin) + '"/><a:ea typeface="' + esc(th.ea) +
        '"/><a:cs typeface="' + esc(th.ea) + '"/></a:' + tag + '>';
    };
    var lnStyle = '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
    var solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    return XMLDECL + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="' + esc(name) + '">' +
      '<a:themeElements>' + clr +
      '<a:fontScheme name="' + esc(name) + '">' + font('majorFont') + font('minorFont') + '</a:fontScheme>' +
      '<a:fmtScheme name="' + esc(name) + '">' +
      '<a:fillStyleLst>' + solid + solid + solid + '</a:fillStyleLst>' +
      '<a:lnStyleLst>' + lnStyle + lnStyle + lnStyle + '</a:lnStyleLst>' +
      '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
      '<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
      '<a:bgFillStyleLst>' + solid + solid + solid + '</a:bgFillStyleLst>' +
      '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
  }

  function clrMapXml() {
    return '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
      'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  }

  function slideMasterXml(th) {
    var ph = function (id, name, type, idx, x, y, w, h, sz) {
      return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + name + '"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
        '<p:nvPr><p:ph type="' + type + '"' + (idx ? ' idx="' + idx + '"' : '') + '/></p:nvPr></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="' + E(x) + '" y="' + E(y) + '"/><a:ext cx="' + E(w) + '" cy="' + E(h) +
        '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN" sz="' + (sz || 1800) + '"/></a:p></p:txBody></p:sp>';
    };
    return XMLDECL + '<p:sldMaster ' + NS + '>' +
      '<p:cSld><p:bg><p:bgPr>' + fillXml(th.bg) + '<a:effectLst/></p:bgPr></p:bg><p:spTree>' + SPRTREE_HEAD +
      ph(2, 'Title Placeholder', 'title', 0, 0.92, 0.5, 11.5, 1.45, 3600) +
      ph(3, 'Body Placeholder', 'body', 1, 0.92, 2.16, 11.5, 4.61, 1800) +
      '</p:spTree></p:cSld>' + clrMapXml() +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3600" b="1"/></a:lvl1pPr></p:titleStyle>' +
      '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle>' +
      '<p:otherStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:otherStyle></p:txStyles>' +
      '</p:sldMaster>';
  }

  function slideLayoutXml() {
    return XMLDECL + '<p:sldLayout ' + NS + ' type="blank" preserve="1">' +
      '<p:cSld name="空白"><p:spTree>' + SPRTREE_HEAD + '</p:spTree></p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
  }

  function notesMasterXml(th) {
    var ph = function (id, name, type, idx, x, y, w, h) {
      return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + name + '"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
        '<p:nvPr><p:ph type="' + type + '"' + (idx ? ' idx="' + idx + '"' : '') + '/></p:nvPr></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="' + E(x) + '" y="' + E(y) + '"/><a:ext cx="' + E(w) + '" cy="' + E(h) +
        '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>';
    };
    return XMLDECL + '<p:notesMaster ' + NS + '>' +
      '<p:cSld><p:bg><p:bgPr>' + fillXml('FFFFFF') + '<a:effectLst/></p:bgPr></p:bg><p:spTree>' + SPRTREE_HEAD +
      ph(2, 'Slide Image Placeholder', 'sldImg', 0, 0.68, 0.46, 4.5, 2.53) +
      ph(3, 'Notes Placeholder', 'body', 1, 0.68, 3.22, 6.14, 6.1) +
      '</p:spTree></p:cSld>' + clrMapXml() +
      '<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle>' +
      '</p:notesMaster>';
  }

  function slideXml(plan, slide, th) {
    var id = 2, shapes = '';
    slide.ops.forEach(function (op) {
      if (op.type === 'text') shapes += textSpXml(id++, op, th);
      else if (op.type === 'table') shapes += tableXml(id++, op, th);
      else shapes += spXml(id++, op);
    });
    return XMLDECL + '<p:sld ' + NS + '><p:cSld>' + bgXml(slide.bg) +
      '<p:spTree>' + SPRTREE_HEAD + shapes + '</p:spTree></p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  }

  function notesSlideXml(notes, th) {
    var paras = String(notes).split(/\n+/).filter(function (l) { return l.trim(); }).map(function (line) {
      return paraXml({ runs: [{ t: line.trim(), sz: 12, color: '333333' }], line: 1.3 }, th);
    }).join('') || paraXml({ runs: [], line: 1.3 }, th, 1200);
    return XMLDECL + '<p:notes ' + NS + '><p:cSld><p:spTree>' + SPRTREE_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/>' +
      '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="' + E(0.68) + '" y="' + E(3.22) + '"/><a:ext cx="' + E(6.14) + '" cy="' + E(6.1) +
      '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/>' + paras + '</p:txBody></p:sp>' +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>';
  }

  function presentationXml(deck, nSlides, hasNotes) {
    var sldIds = '';
    for (var i = 0; i < nSlides; i++) sldIds += '<p:sldId id="' + (256 + i) + '" r:id="rId' + (4 + i) + '"/>';
    return XMLDECL + '<p:presentation ' + NS + ' saveSubsetFonts="1" autoCompressPictures="0">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      (hasNotes ? '<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst>' : '') +
      '<p:sldIdLst>' + sldIds + '</p:sldIdLst>' +
      '<p:sldSz cx="' + E(13.333) + '" cy="' + E(7.5) + '" type="screen16x9"/>' +
      '<p:notesSz cx="' + E(7.5) + '" cy="' + E(10) + '"/>' +
      '</p:presentation>';
  }

  function coreXml(deck) {
    var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    return XMLDECL +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(deck.title || '演示文稿') + '</dc:title>' +
      '<dc:creator>' + esc(deck.author || 'PPT Studio') + '</dc:creator>' +
      '<cp:lastModifiedBy>' + esc(deck.author || 'PPT Studio') + '</cp:lastModifiedBy>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>';
  }

  function appXml(nSlides, nNotes) {
    return XMLDECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<Application>PPT Studio</Application><PresentationFormat>宽屏</PresentationFormat>' +
      '<Slides>' + nSlides + '</Slides><Notes>' + nNotes + '</Notes>' +
      '<HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop>' +
      '<Company></Company><AppVersion>16.0000</AppVersion></Properties>';
  }

  var REL = {
    office: 'officeDocument/2006/relationships/officeDocument',
    core: 'package/2006/relationships/metadata/core-properties',
    app: 'officeDocument/2006/relationships/extended-properties',
    slide: 'officeDocument/2006/relationships/slide',
    slideMaster: 'officeDocument/2006/relationships/slideMaster',
    slideLayout: 'officeDocument/2006/relationships/slideLayout',
    notesMaster: 'officeDocument/2006/relationships/notesMaster',
    notesSlide: 'officeDocument/2006/relationships/notesSlide',
    theme: 'officeDocument/2006/relationships/theme'
  };

  /** 只负责生成所有 part（便于自动化测试逐个检查 XML）。 */
  function partsFor(deck) {
    var plan = PPT.layout.build(deck);
    var th = plan.theme;
    var slides = plan.slides;
    var notesFlags = slides.map(function (s) { return !!(s.notes && String(s.notes).trim()); });
    var hasNotes = notesFlags.some(Boolean);
    var parts = [];
    var overrides = [];

    function add(name, data, ct) {
      parts.push({ name: name, data: data });
      if (ct) overrides.push('<Override PartName="/' + name + '" ContentType="' + ct + '"/>');
    }

    var CT = {
      pres: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
      master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
      layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
      slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
      notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
      theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
      core: 'application/vnd.openxmlformats-package.core-properties+xml',
      app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
    };

    /* 主题 */
    add('ppt/theme/theme1.xml', themeXml(th, 'PPT Studio'), CT.theme);
    if (hasNotes) add('ppt/theme/theme2.xml', themeXml(th, 'PPT Studio Notes'), CT.theme);

    /* 母版与版式 */
    add('ppt/slideMasters/slideMaster1.xml', slideMasterXml(th), CT.master);
    add('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml(), CT.layout);
    add('ppt/slideMasters/_rels/slideMaster1.xml.rels', relsXml([
      { id: 'rId1', type: REL.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
      { id: 'rId2', type: REL.theme, target: '../theme/theme1.xml' }
    ]));
    add('ppt/slideLayouts/_rels/slideLayout1.xml.rels', relsXml([
      { id: 'rId1', type: REL.slideMaster, target: '../slideMasters/slideMaster1.xml' }
    ]));

    /* 备注母版 */
    if (hasNotes) {
      add('ppt/notesMasters/notesMaster1.xml', notesMasterXml(th), CT.notesMaster);
      add('ppt/notesMasters/_rels/notesMaster1.xml.rels', relsXml([
        { id: 'rId1', type: REL.theme, target: '../theme/theme2.xml' }
      ]));
    }

    /* 幻灯片 + 备注页 */
    var nNotes = 0;
    slides.forEach(function (s, i) {
      var n = i + 1;
      var rels = [{ id: 'rId1', type: REL.slideLayout, target: '../slideLayouts/slideLayout1.xml' }];
      var notesRel = '';
      if (notesFlags[i]) {
        nNotes++;
        notesRel = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/' + REL.notesSlide +
          '" Target="../notesSlides/notesSlide' + n + '.xml"/>';
        add('ppt/notesSlides/notesSlide' + n + '.xml', notesSlideXml(s.notes, th), CT.notesSlide);
        add('ppt/notesSlides/_rels/notesSlide' + n + '.xml.rels', relsXml([
          { id: 'rId1', type: REL.slide, target: '../slides/slide' + n + '.xml' },
          { id: 'rId2', type: REL.notesMaster, target: '../notesMasters/notesMaster1.xml' }
        ]));
      }
      add('ppt/slides/slide' + n + '.xml', slideXml(plan, s, th), CT.slide);
      add('ppt/slides/_rels/slide' + n + '.xml.rels',
        relsXml(rels).replace('</Relationships>', notesRel + '</Relationships>'));
    });

    /* presentation */
    add('ppt/presentation.xml', presentationXml(deck, slides.length, hasNotes), CT.pres);
    var presRels = [
      { id: 'rId1', type: REL.slideMaster, target: 'slideMasters/slideMaster1.xml' },
      { id: 'rId3', type: REL.theme, target: 'theme/theme1.xml' }
    ];
    if (hasNotes) {
      presRels.push({ id: 'rId2', type: REL.notesMaster, target: 'notesMasters/notesMaster1.xml' });
    }
    slides.forEach(function (s, i) {
      presRels.push({ id: 'rId' + (4 + i), type: REL.slide, target: 'slides/slide' + (i + 1) + '.xml' });
    });
    add('ppt/_rels/presentation.xml.rels', relsXml(presRels));

    /* docProps */
    add('docProps/core.xml', coreXml(deck), CT.core);
    add('docProps/app.xml', appXml(slides.length, nNotes), CT.app);

    /* 根关系 */
    add('_rels/.rels', relsXml([
      { id: 'rId1', type: REL.office, target: 'ppt/presentation.xml' },
      { id: 'rId2', type: REL.core, target: 'docProps/core.xml' },
      { id: 'rId3', type: REL.app, target: 'docProps/app.xml' }
    ]));

    /* 内容类型（放最前，符合常规顺序） */
    var ct = XMLDECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      overrides.join('') + '</Types>';
    parts.unshift({ name: '[Content_Types].xml', data: ct });

    return { parts: parts, plan: plan };
  }

  function build(deck) {
    var r = partsFor(deck);
    return PPT.zip.buildZip(r.parts);
  }

  PPT.pptx = { build: build, partsFor: partsFor, esc: esc };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
