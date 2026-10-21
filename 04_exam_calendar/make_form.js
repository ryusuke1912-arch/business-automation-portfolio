// 入試日程 記入用紙（生徒配布用）を作る
// 使い方:
//   node make_form.js                → 名前欄が空欄の用紙（1枚。必要枚数を印刷）
//   node make_form.js names.csv      → 生徒ごとに名字入りの用紙（1人1ページ）
// names.csv の形式（1行1人・見出し行なし）: 学年,名字   例) 中3,石田
const fs = require('fs');
const path = require('path');
const J = require(path.join(__dirname, '..', '.claude', 'skills', 'word-template', 'lib', 'jdoc.js'));

const namesFile = process.argv[2];
const out = process.argv[3] || path.join(__dirname, namesFile ? '入試日程_記入用紙_名前入り.docx' : '入試日程_記入用紙.docx');

const GRADES = ['小6', '中3', '高3'];
const zen = (s) => s.replace(/[0-9]/g, (d) => String.fromCharCode(d.charCodeAt(0) + 0xFEE0));

function readNames(file) {
  return fs.readFileSync(file, 'utf8').replace(/^﻿/, '').split(/\r?\n/)
    .map((l) => l.trim()).filter((l) => l && !l.startsWith('#'))
    .map((l) => { const [g, n] = l.split(/[,，\t]/).map((s) => s.trim()); return { grade: g, name: n || '' }; });
}

function page({ grade, name }) {
  // 学年：名前入りなら該当に■、空欄なら□を並べる
  const gradeLine = GRADES.map((g) => (grade === g ? '■' : '□') + zen(g)).join('　');
  const nameText = name ? `${name}　　　　` : '　　　　　　　　　　';
  const rows = [['', '受験日', '学校名', '区分', '合格発表日']];
  for (let i = 1; i <= 7; i++) rows.push([zen(String(i)), '　　月　　日（　　）', '', '', '　　月　　日']);
  return [
    ...J.letterHead({ date: '提出日　　　月　　日', to: ['受験生のみなさん'] }),
    ...J.title('入試日程 記入用紙'),
    J.blank(),
    J.body('受験する学校と日にちを、日にちの早い順に書いてください。教室の入試カレンダーにまとめて、直前の対策と当日の応援に使います。まだ決まっていない学校は、区分の欄に「予定」と書いてください。'),
    J.table([2400, J.W - 2610], [
      ['学年', gradeLine],
      ['氏名', { v: `${nameText}` }],
    ], { firstCol: true, noHead: true, rowH: 520, mid: [] }),
    J.head('１．受験する学校'),
    J.table([520, 2350, 3000, 1600, 1674], rows, { mid: [0, 1, 4], rowH: 570 }),
    J.note('※ 区分の例：専願、併願、推薦、一般、前期、後期、Ａ日程など。'),
    J.note('※ 日程が変わったときや、受験校を追加・取りやめたときは、すぐに教室長へ伝えてください。'),
    J.head('２．先生に伝えておきたいこと（不安な科目、面接の有無など）'),
    J.table([J.W - 210], [['\n\n\n']], { noHead: true }),
    J.right('保護者確認　＿＿＿＿＿＿＿＿＿＿'),
  ];
}

const list = namesFile ? readNames(namesFile) : [{ grade: '', name: '' }];
const children = [];
list.forEach((s, i) => { if (i > 0) children.push(J.pageBreak()); children.push(...page(s)); });
J.save(children, out).then((p) => console.log(`作成: ${p}（${list.length}ページ）`));
