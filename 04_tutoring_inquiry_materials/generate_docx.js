const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
  ShadingType, LineRuleType,
} = require('docx');
const fs = require('fs');

const MIN = { ascii: '游明朝', eastAsia: '游明朝', hAnsi: '游明朝', cs: '游明朝' };
const GO  = { ascii: '游ゴシック', eastAsia: '游ゴシック', hAnsi: '游ゴシック', cs: '游ゴシック' };

const kids = [];

function title(t) {
  kids.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [new TextRun({ text: t, font: GO, bold: true, color: '2E74B5', size: 30 })],
  }));
}

function sub(t) {
  kids.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
    children: [new TextRun({ text: t, font: MIN, color: '595959', size: 20 })],
  }));
}

function sec(t) {
  kids.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'DEEAF6' },
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text: '◆ ' + t, font: GO, bold: true, size: 24 })],
  }));
}

// 本文行（「。」ごとに1行。「、」では改行しない）
function ln(t, opts = {}) {
  kids.push(new Paragraph({
    indent: { left: 200 },
    spacing: { after: 50, line: 285, lineRule: LineRuleType.AUTO, before: opts.before },
    children: [new TextRun({ text: t, font: MIN, size: 21, bold: opts.bold })],
  }));
}

// 小見出し（セクション内）
function sm(t) {
  kids.push(new Paragraph({
    indent: { left: 200 },
    spacing: { before: 160, after: 60, line: 285, lineRule: LineRuleType.AUTO },
    children: [new TextRun({ text: t, font: GO, bold: true, size: 21 })],
  }));
}

function blank() {
  kids.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
}

// ============================================================

title('ご案内資料');
sub('保護者様へ　／　2026年9月4日　【塾名】【教室名】');

sec('はじめに、お約束');
ln('・こちらからお電話をおかけすることは、一切ございません。');
ln('ご連絡はすべてメールのみで承ります。');
ln('・資料送付後の確認のお電話も、体験のご案内のお電話もいたしません。');
ln('・いただいたご要望は教室内で共有し、記録として残しております。');
ln('・ご検討の途中でお気持ちが変わった場合も、お断りのご連絡は不要です。');
ln('こちらから追いかけることはいたしません。');

sec('「通っているのに身につかない」について');
ln('・はじめに申し上げます。');
ln('プリント教材による自学自習の学習は、とてもよくできた仕組みです。');
ln('毎日机に向かう習慣がつき、計算の速さと正確さが育ちます。');
ln('・そのうえで、手ごたえが薄くなる理由が三つあります。');
ln('教材や先生の質ではなく、学習方式の得意・不得意の話です。');

sm('（1）「できた」と「わかった」が、答案では見分けられない');
ln('・○がついた問題が、意味を理解して解いた○なのか、手順を覚えていただけの○なのか。');
ln('答案の上では、まったく同じ形をしています。');
ln('・低学年のうちは、この差が点数に出ません。');
ln('差が出るのは、問題のパターンが増える3年生からです。');

sm('（2）つまずきの一つ手前まで、戻りにくい');
ln('・繰り下がりが苦手な原因は、引き算そのものではないことが多いです。');
ln('「10といくつ」という数の見方が、まだ定まっていないだけの場合があります。');
ln('・自習型の教材は、間違えた単元を繰り返す設計です。');
ln('ただ、その手前まで戻る判断は人がします。');
ln('・結果として、苦手なまま繰り返すことになります。');
ln('まじめなお子様ほど、時間だけが積み上がっていきます。');

sm('（3）「わからない」と、言い出せない');
ln('・自分で進める学習は、止まったときに誰も気づきません。');
ln('・小学2年生の女の子は、まじめで気を遣うお子様が多い時期です。');
ln('わからなくても手を挙げない。');
ln('「わかった？」と聞かれれば、うなずく。');
ln('・性格の問題ではありません。');
ln('質問することにも、少し訓練が必要なのだと考えています。');

sec('個別指導で変わること');
ln('・大事にしているのは、たくさん教えることではなく、気づくことです。');
ln('・解けた問題も、お子様のほうから説明していただきます。');
ln('「23－8は、どうやって計算した？」');
ln('・「なんとなく」と返ってくれば、手順の記憶による○だと判断します。');
ln('「20から8はひけないから」と言葉が出れば、理解に届いています。');
ln('・答案の上には現れません。');
ln('問い返して、はじめて見えるものです。');
ln('・鉛筆の止まり方も見ています。');
ln('問題文で止まるのか、式を書いたあとで止まるのか。');
ln('止まる場所が違えば、必要な手当ても変わります。');

sm('以前お預かりした、小学2年の女の子の例');
ln('・漢字テストは、いつも満点に近いお子様でした。');
ln('けれど、文章題になると点数が安定しませんでした。');
ln('・音読はとてもきれいでした。');
ln('ただ、「この子はどうして泣いたと思う？」と尋ねると、黙ってしまわれました。');
ln('・文字は読めている。');
ln('けれど、場面が頭に浮かんでいない。');
ln('・練習量を増やすのではなく、音読のあとに二つだけ質問することから始めました。');
ln('「だれが出てきた？」「どこの話？」');
ln('・三か月ほどで、聞く前にご自分から話し出すようになりました。');
ln('文章題が上がったのは、そのあとです。');

sec('いま、切り替える意味');
ln('・3年生から理科と社会が加わり、算数では割り算や分数が出てきます。');
ln('・3年生でつまずくお子様の多くは、3年生の内容が難しいのではありません。');
ln('土台の一部が抜けたまま、量が増えた状態です。');
ln('・2年生の取りこぼしは、数週間から数か月で埋められます。');
ln('同じ内容を5年生で埋め直すと、負担は何倍にもなります。');
ln('・ただ、今始めなければ手遅れ、ということはありません。');
ln('早いほど労力が少なくて済む。');

sec('見学・体験の流れ（約90分）');
ln('・教室のご案内（10分）');
ln('・いまの学習状況をおうかがい（15分）');
ln('・お子様に、簡単な問題に取り組んでいただきます（15分）');
ln('テストではありません。点数も順位もつけません。');
ln('見ているのは、どこで手が止まるかです。');
ln('・体験授業（40分）');
ln('ご見学いただいても、いったんお帰りいただいても構いません。');
ln('・気づいた点を、その場でご報告します（15分）');

ln('・その場でご入塾を決めていただくことは、一切ございません。', { before: 140 });
ln('申込書へのご記入もお願いしません。');
ln('・必ず一度、ご家庭に持ち帰ってご検討ください。');
ln('・体験の翌日に、ご報告書をメールでお送りします。');
ln('そのあとは、保護者様からのご連絡をお待ちします。');
ln('催促はいたしません。');

sec('9月体験・10月入塾の場合の目安');
ln('・9月中旬　　　　メールで体験日を決定');
ln('・9月中旬〜下旬　見学・体験');
ln('・9月末まで　　　ご家庭でご検討');
ln('・10月上旬　　　ご入塾・授業開始');

sec('よくあるご質問');
ln('・体験だけで、入塾しなくても構いませんか。');
ln('　→　もちろんです。');
ln('　　　お断りのご連絡すら不要です。');
ln('・いまの教室と、しばらく併用できますか。');
ln('　→　できます。');
ln('　　　内容が重ならないよう、こちらで調整いたします。');
ln('・女性講師を希望できますか。');
ln('　→　承ります。');
ln('　　　体験の段階から女性講師が担当します。');
ln('・宿題はどのくらい出ますか。');
ln('　→　1回15〜20分程度が目安です。');

sec('ご連絡について');
ln('・ご連絡はメールのみで承ります。');
ln('　【メールアドレス】');
ln('・24時間以内にお返事いたします。');
ln('・体験をご希望の場合は、お子様のお名前と、ご希望日を三つお書きいただくだけで結構です。');

ln('※補足', { before: 240, bold: true });
ln('この資料をご覧になって当教室が合わないとご判断された場合は、そのままお読み捨てください。');

kids.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { before: 160, after: 0 },
  children: [new TextRun({ text: '【塾名】【教室名】　教室長　【教室長名】', font: MIN, size: 21 })],
}));
kids.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { after: 0 },
  children: [new TextRun({ text: '【住所】　【メールアドレス】', font: MIN, size: 21 })],
}));

const doc = new Document({
  styles: { default: { document: { run: { font: MIN, size: 21 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
      },
    },
    children: kids,
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('ご案内資料.docx', b);
  console.log('ok', b.length);
});
