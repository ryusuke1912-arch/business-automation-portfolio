// 日本語ビジネス文書（Word）ビルダー
// ＭＳ明朝 10.5pt／黒＋薄グレー／「。」ごとに改行／記・以上の体裁
// 使い方は SKILL.md と examples/sample.js を参照
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, BorderStyle, VerticalAlign, Footer, PageNumber, PageBreak,
} = require('docx');

const F = 'ＭＳ 明朝';
const S = 21;            // 10.5pt（半ポイント単位）
const W = 9354;          // 本文幅（A4・左右余白 22.5mm）
const GRAY = 'D9D9D9';   // 表の見出し行
const LIGHT = 'F2F2F2';  // 表の左端列

// 「。」の直後で区切る（「。」で終わらない末尾もそのまま残す）
const sentences = (text) => String(text).split(/(?<=。)/).map(s => s.trim()).filter(Boolean);

const run = (text, o = {}) => new TextRun({
  text, font: { ascii: F, eastAsia: F, hAnsi: F, cs: F },
  size: o.size || S, bold: o.bold, underline: o.u ? {} : undefined, break: o.br ? 1 : undefined,
});
// 「。」ごとに「段落」を分ける（Shift+Enter の改行は使わない。
// 日本語Wordは両端揃えのため、行内改行だと字間が引き伸ばされる）
// first: 1文目の段落設定 / rest: 2文目以降の段落設定
const paras = (text, ro, first, rest) =>
  sentences(text).map((s, i) => P([run(s, ro)], i === 0 ? first : { ...first, ...rest }));

const P = (children, o = {}) => new Paragraph({
  children, alignment: o.align || AlignmentType.LEFT, indent: o.indent, keepNext: o.keepNext,
  spacing: { before: o.before || 0, after: o.after || 0 },
  heading: o.heading,
});

// ---------- 文書の部品 ----------
const blank = () => P([run('')]);
const right = (text, o = {}) => P([run(text, o)], { align: AlignmentType.RIGHT });
const left = (text, o = {}) => P([run(text, o)]);
const center = (text, o = {}) => P([run(text, o)], { align: AlignmentType.CENTER });
const title = (text, subtitle) => [
  P([run(text, { size: 28, bold: true })], { align: AlignmentType.CENTER, before: 120 }),
  ...(subtitle ? [center(subtitle)] : []),
];
const body = (text) => paras(text, {}, { indent: { left: 210 } }, {});                // 本文（「。」で段落を分ける）
const head = (text) => P([run(text, { bold: true })], { before: 240, after: 60, keepNext: true, heading: 'Heading1' });   // １．見出し
const sub = (text) => P([run(text, { bold: true })], { before: 120, after: 40, keepNext: true, indent: { left: 210 }, heading: 'Heading2' }); // （１）小見出し
const item = (text, mark = '・') => paras(mark + text, {}, { indent: { left: 420, hanging: 210 } }, { indent: { left: 420 } });  // 箇条書き
const check = (text) => item(text, '□');
const note = (text) => paras(text, { size: 18 }, { before: 40, indent: { left: 420, hanging: 210 } }, { before: 0, indent: { left: 420 } }); // ※注記（9pt）
const ki = () => center('記');
const ijo = () => right('以上');
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// 宛名・日付・差出人ブロック
function letterHead({ date, to = [], from = [] }) {
  return [right(date), ...to.map(t => left(t)), ...from.map(f => right(f)), blank()];
}

// ---------- 表 ----------
const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: line, bottom: line, left: line, right: line };
/**
 * widths: 列幅（DXA）。合計は W - indent 以内に。
 * rows  : 1行目は見出し。セルは文字列（\n で改行、「。」でも自動改行）か {v, bold, align}
 * o     : { size, firstCol, mid:[中央揃えの列], indent, rowH, noHead }
 */
function table(widths, rows, o = {}) {
  const fsz = o.size || S;
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    indent: { size: o.indent ?? 210, type: WidthType.DXA },
    rows: rows.map((row, ri) => new TableRow({
      tableHeader: ri === 0 && !o.noHead,
      cantSplit: true,
      height: o.rowH && ri > 0 ? { value: o.rowH, rule: 'atLeast' } : undefined,
      children: row.map((c, ci) => {
        const cell = (c && typeof c === 'object') ? c : { v: c };
        const isHead = ri === 0 && !o.noHead;
        const side = o.firstCol && ci === 0 && !isHead;
        const lines = String(cell.v ?? '').split('\n').flatMap(l => l === '' ? [''] : sentences(l));
        return new TableCell({
          width: { size: widths[ci], type: WidthType.DXA },
          borders, verticalAlign: VerticalAlign.CENTER,
          margins: { top: 40, bottom: 40, left: 100, right: 100 },
          shading: isHead ? { type: ShadingType.CLEAR, fill: GRAY, color: 'auto' }
            : side ? { type: ShadingType.CLEAR, fill: LIGHT, color: 'auto' } : undefined,
          children: lines.map(l => new Paragraph({
            alignment: isHead ? AlignmentType.CENTER
              : (cell.align || ((o.mid || []).includes(ci) ? AlignmentType.CENTER : AlignmentType.LEFT)),
            children: [run(l, { size: fsz, bold: cell.bold })],
          })),
        });
      }),
    })),
  });
}

// 罫線で囲んだ所見・ポイント欄（1行目は太字の見出し）
function frame(lines) {
  return new Table({
    width: { size: W - 210, type: WidthType.DXA }, columnWidths: [W - 210],
    indent: { size: 210, type: WidthType.DXA },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: W - 210, type: WidthType.DXA }, borders,
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: lines.flatMap((l, i) => i === 0
        ? [P([run(l, { bold: true })])]
        : paras(l, {}, { indent: { left: 210, hanging: 210 } }, { indent: { left: 210 } })),
    })] })],
  });
}

// ---------- 出力 ----------
function makeDoc(children) {
  children = children.flat(Infinity);
  const font = { ascii: F, eastAsia: F, hAnsi: F, cs: F };
  return new Document({
    compatibility: { doNotExpandShiftReturn: true },
    styles: {
      default: { document: { run: { font, size: S } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font, size: S, bold: true, color: '000000' }, paragraph: { outlineLevel: 0 } },
        { id: 'Heading2', name: 'heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font, size: S, bold: true, color: '000000' }, paragraph: { outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: { size: { width: 11906, height: 16838 },
          margin: { top: 1701, bottom: 1418, left: 1276, right: 1276, header: 851, footer: 851 } },
        grid: { linePitch: 360, type: 'lines' },
      },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ['- ', PageNumber.CURRENT, ' -'], font: F, size: 18 })] })] }) },
      children,
    }],
  });
}
async function save(children, path) {
  fs.writeFileSync(path, await Packer.toBuffer(makeDoc(children)));
  return path;
}

module.exports = {
  W, S, sentences, run, paras, blank, right, left, center, title, body, head, sub, item, check, note,
  ki, ijo, pageBreak, letterHead, table, frame, makeDoc, save, AlignmentType,
};
