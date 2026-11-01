/*
 * 原稿用紙レイアウト ＋ Word(.docx) 生成（依存ライブラリなし）
 * ブラウザでは window.Genkou、Node では module.exports で使える。
 */
(function (root) {
  'use strict';

  const COLS = 20;

  // 行頭に来てはいけない文字（前の行の最終マスに入れる）
  const PUNCT = new Set(['、', '。', '，', '．']);
  const CLOSE_BRACKETS = new Set(['」', '』', '）', '］', '】', '〕', '〉', '》', '｝']);

  /** 半角文字か（ASCII 印字可能文字・半角スペース・半角カナ） */
  function isHalfWidth(ch) {
    const c = ch.codePointAt(0);
    return (c >= 0x20 && c <= 0x7e) || (c >= 0xff61 && c <= 0xff9f);
  }

  /**
   * 文字列をマス（cell）の配列にする。
   * cell = { text, half: bool, hang: string }  hang は行末にぶら下げた句読点
   */
  function toCells(text) {
    const cells = [];
    for (const ch of Array.from(text)) {
      if (isHalfWidth(ch)) {
        const last = cells[cells.length - 1];
        if (last && last.half && Array.from(last.text).length === 1) {
          last.text += ch; // 半角は1マスに2文字
          continue;
        }
        cells.push({ text: ch, half: true, hang: '' });
      } else {
        cells.push({ text: ch, half: false, hang: '' });
      }
    }
    return cells;
  }

  /**
   * マス配列を20マスの行に割り付ける。
   * 行頭に来る句読点（opts.brackets なら閉じ括弧も）は前行の最終マスに入れる。
   * @returns {Array<Array<cell>>}
   */
  function toRows(cells, opts) {
    opts = opts || {};
    const forbidden = (cell) =>
      !cell.half && (PUNCT.has(cell.text) || (opts.brackets && CLOSE_BRACKETS.has(cell.text)));
    const rows = [];
    let row = [];
    for (const cell of cells) {
      if (row.length === COLS) {
        if (forbidden(cell)) {
          row[COLS - 1].hang += cell.text;
          continue;
        }
        rows.push(row);
        row = [];
      }
      row.push(cell);
    }
    if (row.length) rows.push(row);
    return rows;
  }

  function normalizeDigits(s) {
    return s.replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0));
  }

  /**
   * 入力テキストを設問ブロックに分ける。
   * 「#」で始まる行は見出し。見出しに「100字」などがあれば制限字数として使う。
   * 見出しがない場合は全体を1つの答案として扱う。
   */
  function parse(input, opts) {
    opts = opts || {};
    const lines = input.replace(/\r\n?/g, '\n').split('\n');
    const blocks = [];
    let cur = null;
    const start = (title) => {
      cur = { title, limit: null, lines: [] };
      if (title) {
        const m = normalizeDigits(title).match(/(\d+)\s*字/);
        if (m) cur.limit = parseInt(m[1], 10);
      }
      blocks.push(cur);
    };
    for (const line of lines) {
      const h = line.match(/^\s*[#＃]\s*(.*)$/);
      if (h) {
        start(h[1].trim());
        continue;
      }
      if (!cur) start('');
      cur.lines.push(line);
    }
    return blocks
      .map((b) => {
        let body;
        if (opts.keepNewlines) {
          body = b.lines.join('\n').replace(/^\n+|\s+$/g, '');
        } else {
          body = b.lines.map((l) => l.trim()).join('');
        }
        if (b.limit == null && opts.defaultLimit) b.limit = opts.defaultLimit;
        return { title: b.title, limit: b.limit, body };
      })
      .filter((b) => b.title || b.body);
  }

  /** 1ブロックをレイアウトする（改行を残す場合は改行ごとに次の行へ） */
  function layoutBlock(block, opts) {
    const paragraphs = block.body.split('\n');
    let rows = [];
    for (const p of paragraphs) {
      const r = toRows(toCells(p), opts);
      rows = rows.concat(r.length ? r : [[]]);
    }
    if (rows.length === 1 && rows[0].length === 0) rows = [];
    const used = rows.reduce((n, r) => n + r.length, 0);
    // 最後の行以外の空きマス（改行による）も字数に数える
    const usedWithBreaks = rows.length ? (rows.length - 1) * COLS + rows[rows.length - 1].length : 0;
    const count = opts && opts.keepNewlines ? usedWithBreaks : used;
    const minRows = block.limit ? Math.ceil(block.limit / COLS) : 1;
    const totalRows = Math.max(minRows, rows.length);
    while (rows.length < totalRows) rows.push([]);
    return {
      title: block.title,
      limit: block.limit,
      count,
      over: block.limit ? count > block.limit : false,
      rows,
    };
  }

  function layout(input, opts) {
    return parse(input, opts).map((b) => layoutBlock(b, opts));
  }

  // ---------------------------------------------------------------- docx

  function esc(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  const CELL = 480; // twips (≒8.47mm)
  const NUM_COL = 300; // 右側の字数目盛り列

  function rPr(font, sz, extra) {
    return (
      '<w:rPr><w:rFonts w:ascii="' + esc(font) + '" w:hAnsi="' + esc(font) + '" w:eastAsia="' + esc(font) +
      '" w:cs="' + esc(font) + '"/>' + (extra || '') + '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/></w:rPr>'
    );
  }

  function run(text, font, sz, extra) {
    return '<w:r>' + rPr(font, sz, extra) + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }

  const P_TIGHT = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';

  function cellXml(cell, font, sz) {
    let runs = '';
    if (cell && cell.text) {
      runs += run(cell.text, font, sz);
      // ぶら下げた句読点は字幅50%に詰めて同じマスに収める
      if (cell.hang) runs += run(cell.hang, font, sz, '<w:w w:val="50"/>');
    }
    return (
      '<w:tc><w:tcPr><w:tcW w:w="' + CELL + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' +
      '<w:p><w:pPr>' + P_TIGHT + '<w:jc w:val="center"/><w:snapToGrid w:val="0"/></w:pPr>' + runs + '</w:p></w:tc>'
    );
  }

  function numCellXml(label, font) {
    const none = '<w:top w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>';
    return (
      '<w:tc><w:tcPr><w:tcW w:w="' + NUM_COL + '" w:type="dxa"/><w:tcBorders>' + none +
      '</w:tcBorders><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:pPr>' + P_TIGHT +
      '<w:ind w:left="40"/></w:pPr>' + (label ? run(label, font, 12, '<w:color w:val="808080"/>') : '') + '</w:p></w:tc>'
    );
  }

  function tableXml(block, font, sz) {
    const border = (name) => '<w:' + name + ' w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>';
    let xml =
      '<w:tbl><w:tblPr><w:tblW w:w="' + (CELL * COLS + NUM_COL) + '" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
      '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') + '</w:tblBorders>' +
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
      '<w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>';
    for (let i = 0; i < COLS; i++) xml += '<w:gridCol w:w="' + CELL + '"/>';
    xml += '<w:gridCol w:w="' + NUM_COL + '"/></w:tblGrid>';
    block.rows.forEach((row, r) => {
      xml += '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="' + CELL + '" w:hRule="exact"/></w:trPr>';
      for (let c = 0; c < COLS; c++) xml += cellXml(row[c], font, sz);
      const n = (r + 1) * COLS;
      xml += numCellXml((r + 1) % 5 === 0 ? String(n) : '', font);
      xml += '</w:tr>';
    });
    return xml + '</w:tbl>';
  }

  function paragraph(text, font, sz, bold, after) {
    return (
      '<w:p><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="' + (after || 0) + '"/></w:pPr>' +
      (text ? run(text, font, sz, bold ? '<w:b/>' : '') : '') + '</w:p>'
    );
  }

  function documentXml(blocks, opts) {
    const font = opts.font || 'ＭＳ 明朝';
    const sz = opts.size || 24;
    let body = '';
    if (opts.title) body += paragraph(opts.title, font, 28, true, 160);
    blocks.forEach((b, i) => {
      const head = [b.title, b.limit ? '（' + b.count + '／' + b.limit + '字）' : '（' + b.count + '字）']
        .filter(Boolean)
        .join(' ');
      body += paragraph(head, font, 21, true, 60);
      body += tableXml(b, font, sz);
      if (i < blocks.length - 1) body += paragraph('', font, 21, false, 120);
    });
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1134" w:right="1000" w:bottom="1134" w:left="1000" w:header="567" w:footer="567" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>'
    );
  }

  const CONTENT_TYPES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';

  const RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  // ---------------------------------------------------------------- zip (無圧縮)

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  function zip(files) {
    const enc = new TextEncoder();
    const chunks = [];
    const central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const crc = crc32(data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // UTF-8
      local.setUint16(8, 0, true); // stored
      local.setUint16(10, 0, true);
      local.setUint16(12, 0x21, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true);
      local.setUint32(22, data.length, true);
      local.setUint16(26, name.length, true);
      local.setUint16(28, 0, true);
      chunks.push(new Uint8Array(local.buffer), name, data);

      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true);
      cd.setUint16(4, 20, true);
      cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, 0, true);
      cd.setUint16(12, 0, true);
      cd.setUint16(14, 0x21, true);
      cd.setUint32(16, crc, true);
      cd.setUint32(20, data.length, true);
      cd.setUint32(24, data.length, true);
      cd.setUint16(28, name.length, true);
      cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((n, a) => n + a.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    const all = chunks.concat(central, [new Uint8Array(end.buffer)]);
    const out = new Uint8Array(all.reduce((n, a) => n + a.length, 0));
    let p = 0;
    for (const a of all) {
      out.set(a, p);
      p += a.length;
    }
    return out;
  }

  /** レイアウト済みブロックから .docx のバイト列を作る */
  function buildDocx(blocks, opts) {
    return zip([
      { name: '[Content_Types].xml', data: CONTENT_TYPES },
      { name: '_rels/.rels', data: RELS },
      { name: 'word/document.xml', data: documentXml(blocks, opts || {}) },
    ]);
  }

  // ---------------------------------------------------------------- pdf
  // 日本語は PDF 標準の非埋め込み CID フォント（Adobe-Japan1 / UniJIS-UCS2-H）を使う。
  // 1文字ずつ座標を指定して描くので、閲覧側の代替フォントの字幅に左右されない。

  const PDF_FONTS = {
    mincho: {
      name: 'HeiseiMin-W3',
      desc: '/Flags 6 /FontBBox [-123 -257 1001 910] /ItalicAngle 0 /Ascent 859 /Descent -141 /CapHeight 709 /StemV 69',
    },
    gothic: {
      name: 'HeiseiKakuGo-W5',
      desc: '/Flags 4 /FontBBox [-92 -250 1010 922] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 737 /StemV 114',
    },
  };

  function pdfHex(s) {
    let h = '';
    for (const ch of Array.from(s)) {
      let c = ch.codePointAt(0);
      if (c > 0xffff) c = 0x3013; // BMP外は〓
      h += c.toString(16).padStart(4, '0');
    }
    return '<' + h.toUpperCase() + '>';
  }

  const f2 = (n) => (Math.round(n * 100) / 100).toString();

  function buildPdf(blocks, opts) {
    opts = opts || {};
    const font = PDF_FONTS[/ゴシック/.test(opts.font || '') ? 'gothic' : 'mincho'];
    const PW = 595.28, PH = 841.89;
    const S = 24; // マスの一辺(pt) ≒ 8.47mm
    const FS = 12; // 本文の字の大きさ
    const LEFT = (PW - S * COLS) / 2 - 8;
    const TOP = 60, BOTTOM = 50;

    const pages = [];
    let ops = null;
    let y = 0; // 現在の描画位置（PDF座標、下から上）
    const newPage = () => {
      ops = [];
      pages.push(ops);
      y = PH - TOP;
    };
    const glyph = (s, x, base, size, scale) => {
      ops.push('BT /F1 ' + size + ' Tf ' + (scale || 100) + ' Tz 1 0 0 1 ' + f2(x) + ' ' + f2(base) + ' Tm ' + pdfHex(s) + ' Tj ET');
    };
    const text = (s, x, base, size, gray) => {
      if (gray) ops.push('0.5 g');
      glyph(s, x, base, size);
      if (gray) ops.push('0 g');
    };

    newPage();
    if (opts.title) {
      text(opts.title, LEFT, y - 14, 14);
      y -= 30;
    }

    blocks.forEach((b, bi) => {
      const head = [b.title, b.limit ? '（' + b.count + '／' + b.limit + '字）' : '（' + b.count + '字）']
        .filter(Boolean)
        .join(' ');
      // 設問が丸ごと入らなければ改ページ（1ページに収まらない長さなら、最低1行入る所から書く）
      const need = 16 + b.rows.length * S;
      const fitsPage = need <= PH - TOP - BOTTOM;
      if ((fitsPage && y - need < BOTTOM) || y - 16 - S < BOTTOM) newPage();
      text(head, LEFT, y - 11, 10.5);
      y -= 16;
      b.rows.forEach((row, r) => {
        if (y - S < BOTTOM) newPage();
        const bottom = y - S;
        ops.push('0.5 G 0.5 w');
        for (let c = 0; c < COLS; c++) ops.push(f2(LEFT + c * S) + ' ' + f2(bottom) + ' ' + S + ' ' + S + ' re');
        ops.push('S');
        const base = bottom + (S - FS) / 2 + FS * 0.13;
        for (let c = 0; c < COLS; c++) {
          const cell = row[c];
          if (!cell || !cell.text) continue;
          const cx = LEFT + c * S + S / 2;
          const chars = Array.from(cell.text);
          if (cell.half) {
            const w = FS / 2;
            let x = cx - (chars.length * w) / 2;
            for (const ch of chars) {
              glyph(ch, x, base, FS);
              x += w;
            }
          } else if (cell.hang) {
            glyph(cell.text, cx - FS * 0.75, base, FS);
            let x = cx + FS * 0.25;
            for (const ch of Array.from(cell.hang)) {
              glyph(ch, x, base, FS, 50);
              x += FS / 2;
            }
          } else {
            glyph(cell.text, cx - FS / 2, base, FS);
          }
        }
        if ((r + 1) % 5 === 0) text(String((r + 1) * COLS), LEFT + S * COLS + 3, bottom + 1, 6, true);
        y = bottom;
      });
      if (bi < blocks.length - 1) y -= 18;
    });

    // ---- オブジェクトを組み立てる
    const objs = []; // index+1 がオブジェクト番号
    const add = (body) => {
      objs.push(body);
      return objs.length;
    };
    const catalogId = add(null);
    const pagesId = add(null);
    const descId = add('<< /Type /FontDescriptor /FontName /' + font.name + ' ' + font.desc + ' >>');
    const cidId = add(
      '<< /Type /Font /Subtype /CIDFontType0 /BaseFont /' + font.name +
        ' /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 2 >> /FontDescriptor ' + descId +
        ' 0 R /DW 1000 /W [1 95 500] >>'
    );
    const fontId = add(
      '<< /Type /Font /Subtype /Type0 /BaseFont /' + font.name + '-UniJIS-UCS2-H /Encoding /UniJIS-UCS2-H /DescendantFonts [' +
        cidId + ' 0 R] >>'
    );
    const pageIds = pages.map((p) => {
      const stream = p.join('\n');
      const contentId = add('<< /Length ' + stream.length + ' >>\nstream\n' + stream + '\nendstream');
      return add(
        '<< /Type /Page /Parent ' + pagesId + ' 0 R /MediaBox [0 0 ' + PW + ' ' + PH + '] /Resources << /Font << /F1 ' +
          fontId + ' 0 R >> >> /Contents ' + contentId + ' 0 R >>'
      );
    });
    objs[catalogId - 1] = '<< /Type /Catalog /Pages ' + pagesId + ' 0 R >>';
    objs[pagesId - 1] = '<< /Type /Pages /Kids [' + pageIds.map((i) => i + ' 0 R').join(' ') + '] /Count ' + pageIds.length + ' >>';

    // 本文はすべて ASCII なので文字数＝バイト数
    let out = '%PDF-1.4\n';
    const offsets = [];
    objs.forEach((body, i) => {
      offsets.push(out.length);
      out += i + 1 + ' 0 obj\n' + body + '\nendobj\n';
    });
    const xref = out.length;
    out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n';
    for (const o of offsets) out += String(o).padStart(10, '0') + ' 00000 n \n';
    out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root ' + catalogId + ' 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
    return new TextEncoder().encode(out);
  }

  const api = { COLS, isHalfWidth, toCells, toRows, parse, layout, buildDocx, documentXml, buildPdf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Genkou = api;
})(typeof window !== 'undefined' ? window : globalThis);
