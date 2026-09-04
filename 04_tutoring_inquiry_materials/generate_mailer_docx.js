const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, Footer, PageNumber, LevelFormat, convertMillimetersToTwip,
} = require('docx');
const fs = require('fs');

const MINCHO = { ascii: 'Yu Mincho', eastAsia: 'Yu Mincho', hAnsi: 'Yu Mincho' };
const GOTHIC = { ascii: 'Yu Gothic', eastAsia: 'Yu Gothic', hAnsi: 'Yu Gothic' };
const BODY_SIZE = 21; // 10.5pt in half-points

// helper: parse **bold** segments
function runs(text, opts = {}) {
  const out = [];
  const parts = String(text).split('**');
  parts.forEach((p, i) => {
    if (p === '') return;
    out.push(new TextRun({ text: p, bold: i % 2 === 1 || opts.bold, font: opts.font || MINCHO, size: opts.size || BODY_SIZE, color: opts.color }));
  });
  return out;
}

function body(text, opts = {}) {
  return new Paragraph({
    children: runs(text, opts),
    spacing: { after: opts.after ?? 120, line: 320 },
    alignment: opts.align,
    indent: opts.indent,
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, font: GOTHIC, size: 28, color: '1F3864' })],
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: '1F3864', space: 4 } },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, font: GOTHIC, size: 24, color: '1F3864' })],
    spacing: { before: 240, after: 120 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, font: GOTHIC, size: BODY_SIZE })],
    spacing: { before: 200, after: 100 },
  });
}

function bullet(text) {
  return new Paragraph({
    children: runs(text),
    numbering: { reference: 'bullets', level: 0 },
    spacing: { after: 60, line: 320 },
  });
}

// callout box: single-cell table with light shading
function callout(paragraphTexts, opts = {}) {
  return new Table({
    width: { size: 9070, type: WidthType.DXA },
    columnWidths: [9070],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 8, color: '1F3864' },
      bottom: { style: BorderStyle.SINGLE, size: 8, color: '1F3864' },
      left: { style: BorderStyle.SINGLE, size: 8, color: '1F3864' },
      right: { style: BorderStyle.SINGLE, size: 8, color: '1F3864' },
    },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: 9070, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: opts.fill || 'F2F5FA' },
        margins: { top: 160, bottom: 160, left: 240, right: 240 },
        children: paragraphTexts.map(t => body(t, { after: 60 })),
      })],
    })],
  });
}

function spacer(h = 120) { return new Paragraph({ children: [], spacing: { after: h } }); }

// ---------- 送付状 ----------
const coverLetter = [
  body('2026年9月4日', { align: AlignmentType.RIGHT, after: 200 }),
  body('保護者様', { size: 24, after: 200 }),
  body('【塾名】【教室名】', { align: AlignmentType.RIGHT, after: 0 }),
  body('教室長　【教室長名】', { align: AlignmentType.RIGHT, after: 0 }),
  body('【住所】', { align: AlignmentType.RIGHT, after: 0 }),
  body('メール: 【メールアドレス】', { align: AlignmentType.RIGHT, after: 240 }),
  new Paragraph({
    children: [new TextRun({ text: '資料のご送付につきまして', bold: true, font: GOTHIC, size: 28 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
  }),
  body('拝啓　残暑の候、保護者様におかれましてはお健やかにお過ごしのことと存じます。このたびは数ある教室の中から当教室の資料をご請求いただき、誠にありがとうございます。'),

  h2('はじめに、お約束をひとつ'),
  callout([
    '当教室から保護者様へ、**お電話をおかけすることは一切ございません。**ご連絡はすべてメールのみで承ります。',
    '資料送付後の確認のお電話も、体験のご案内のお電話もいたしません。お問い合わせフォームにいただいたご要望（お電話不可・メール可）は、教室内で共有し、記録として残しております。どうぞご安心ください。',
  ]),
  spacer(),
  body('塾選びの主導権は、保護者様のお手元にあるべきものと考えております。資料をご覧いただき、比べていただき、ご納得のいくタイミングでご連絡をいただければ十分です。ご検討の途中でお気持ちが変わった場合も、お断りのご連絡すら不要です。当教室から追いかけるようなことはいたしません。'),

  h2('同封の資料について'),
  body('お問い合わせの際に「なかなか身についておらず」とお書きくださっていました。その一行に、これまでお子様の様子をきちんと見てこられた時間があるのだろうと感じております。まじめに通っていて、宿題もやっている。それなのに手ごたえが薄い――このご相談は、当教室でもたいへん多くいただきます。'),
  body('そこで本資料では、いわゆる「入塾のご案内」よりも先に、**なぜそれが起こるのか**を最初のページに置きました。原因のご説明にご納得いただけなければ、そのあとの指導方針のお話にも意味がないと考えるためです。'),
  bullet('「まじめに通っているのに、身につかない」が起こる3つの理由'),
  bullet('個別指導で変わること'),
  bullet('低学年のいま、個別指導に切り替える意味'),
  bullet('見学・体験の流れ／よくあるご質問／ご連絡方法'),
  body('料金につきましては、別紙の料金表に月謝・教材費・その他費用のすべてを記載しております。追加でご請求するものはございません。', { after: 160 }),

  h2('9月のご見学・ご体験について'),
  body('9月中にご見学・ご体験、10月にご入塾をお考えとのこと、承知いたしました。日程のご調整はメールのみで完結いたします。ご希望日を第3希望までお書き添えいただければ、空き状況を確認のうえ**24時間以内**にご返信いたします。お電話でのやり取りは発生いたしません。'),
  body('なお、資料をご覧になった結果、当教室が合わないとご判断されることも当然あろうかと存じます。それも含めて、比較のための一資料としてお使いいただければ幸いです。'),
  body('末筆ながら、お子様の毎日が実り多いものでありますよう、心よりお祈り申し上げます。', { after: 200 }),
  body('敬具', { align: AlignmentType.RIGHT }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---------- 本資料 ----------
const main = [];

main.push(new Paragraph({
  children: [new TextRun({ text: 'ご案内資料', bold: true, font: GOTHIC, size: 36, color: '1F3864' })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 120, after: 60 },
}));
main.push(new Paragraph({
  children: [new TextRun({ text: '― まじめに通っているのに、身につかない。その理由と、個別指導にできること ―', font: GOTHIC, size: 22, color: '404040' })],
  alignment: AlignmentType.CENTER,
  spacing: { after: 320 },
}));

main.push(h1('１．「まじめに通っているのに、身につかない」が起こる3つの理由'));
main.push(body('はじめに申し上げておきたいことがございます。プリント教材による自学自習型の学習は、非常によくできた仕組みです。反復によって計算の速度と正確さを育て、毎日机に向かう習慣をつくる力があります。実際、当教室に来られるお子様の中でも、そうした学習を経験されたお子様は、**姿勢がよく、集中して手を動かせる**という共通点をお持ちです。これは簡単に身につくものではありません。'));
main.push(body('そのうえで、「まじめに続けているのに手ごたえが薄い」という状態には、いくつか構造的な理由があります。教材や指導者の質の問題ではなく、**学習方式そのものが持つ得意分野と不得意分野**のお話です。'));

main.push(h3('理由1　「できた」と「わかった」が、答案の上では区別できない'));
main.push(body('プリントは正解か不正解かを記録します。しかし、○がついた問題が、'));
main.push(bullet('意味を理解して解いた○なのか'));
main.push(bullet('手順を覚えていて解いた○なのか'));
main.push(bullet('前の問題と同じやり方だと推測して解いた○なのか'));
main.push(body('――この3つは、答案用紙の上では完全に同じ形をしています。'));
main.push(body('低学年のうちは、この3種類の○の差が成績に出ません。問題のパターンが少なく、手順の記憶だけでかなりのところまで正解できてしまうためです。差が表面化するのは、パターンが急に増える3年生以降になります。'));
main.push(body('「身についていない感じがする」というご感覚は、多くの場合、**この3種類の○を、保護者様が日々のご様子から感じ取っておられる**ということだと私どもは考えております。'));

main.push(h3('理由2　つまずきの「一つ手前」まで戻る仕組みが働きにくい'));
main.push(body('たとえば繰り下がりのある引き算が苦手なお子様がいるとします。原因は引き算そのものではなく、多くの場合その手前にあります。'));
main.push(bullet('10のまとまりを崩して考えることに慣れていない'));
main.push(bullet('「10といくつ」という数の見方が定着していない'));
main.push(bullet('そもそも数字を書く位置がずれて、位を見失っている'));
main.push(body('自学自習型の教材は、間違えた単元を繰り返す設計にはなっていますが、**間違えた単元より前に戻る**判断は、基本的に人がしなければなりません。ご家庭でこれを見抜くのは、専門的な訓練を受けていない限り、率直に申し上げてかなり難しいことです。'));
main.push(body('結果として、苦手な単元を、苦手なまま何度も繰り返すことになります。まじめなお子様ほどきちんと繰り返すため、時間だけが積み上がっていきます。'));

main.push(h3('理由3　「わからない」と言い出せないまま、時間が過ぎる'));
main.push(body('自分で進める学習は、裏を返せば「止まったときに誰も気づかない」学習でもあります。'));
main.push(body('とくに小学2年生の女の子は、真面目で、周囲に気を遣うお子様が多い時期です。わからなくても手を挙げない。とりあえず何か書いておく。「わかった？」と聞かれれば、うなずく。これはお子様の性格の問題ではなく、**質問することにもある程度の訓練が必要**なのだと私どもは考えております。'));
main.push(body('まじめさ・素直さ・我慢強さといった長所が、そのまま「つまずきが見えにくい」という方向に働いてしまう。これが3つ目の理由です。'));

main.push(h1('２．個別指導で変わること'));
main.push(body('個別指導が担うのは、上の3つの理由をひっくり返す役割です。「たくさん教える」ことではなく、**気づく**ことにあります。'));

main.push(h3('「わかったつもり」を、対話で見つける'));
main.push(body('当教室では、解き終わった問題について、しばしばお子様のほうに説明していただきます。'));
main.push(body('たとえば「23−8」を正解したお子様に「どうやって計算したか、先生に教えてくれる？」と尋ねます。'));
main.push(body('「わかんない、なんとなく」――この答えが返ってきたとき、その○は手順の記憶による○である可能性が高い、と判断します。逆に「20から8はひけないから、13から8をひいて…」と言葉が出てくれば、理解に届いています。'));
main.push(body('説明できるかどうかは、答案の上には現れません。**問い返して、初めて見えるもの**です。この一手間を毎回入れられることが、個別指導の中身です。'));

main.push(h3('鉛筆の止まり方で気づく'));
main.push(body('講師が見ているのは答案だけではありません。'));
main.push(bullet('問題文を読んだ直後に止まるのか（読み取りでつまずいている）'));
main.push(bullet('式を書いたあとに止まるのか（計算でつまずいている）'));
main.push(bullet('書いては消し、を繰り返すのか（自信がない）'));
main.push(bullet('消しゴムを使わずすらすら進むのに、答えが違うのか（手順を誤って覚えている）'));
main.push(body('止まる場所が違えば、必要な手当ても違います。小学2年生のお子様は、まだ「どこがわからないか」を言葉にできません。ですから、**言葉にならないものを、様子から読む**のが講師の仕事になります。'));

main.push(h3('実際にあったご様子（小学2年・女の子の例）'));
main.push(body('ご入会当初、漢字テストはいつも満点に近いお子様でした。ところが国語の文章題になると、点数が安定しません。'));
main.push(body('体験時に音読していただいたところ、句読点でほとんど息を切らずに、一定の速さで最後まで読み切られました。読み方はとてもきれいでした。ただ、「この女の子は、どうして泣いたんだと思う？」とお尋ねすると、しばらく黙ってしまわれました。'));
main.push(body('**文字は読めている。けれど、場面が頭に浮かんでいない。**漢字を「書ける」ことと、その言葉の意味を「使える」ことが、別々に育っていた状態でした。'));
main.push(body('そこで、読解の練習量を増やすのではなく、音読のあとに毎回2つだけ質問することから始めました。「だれが出てきた？」「どこの話？」。この2問だけです。'));
main.push(body('3か月ほどで、質問する前にご自分から「この子ね、たぶん…」と話し出されるようになりました。文章題の正答率が上がったのは、そのあとです。'));
main.push(body('算数の文章題についても同様で、当教室では**「ぜんぶで」「のこりは」「ちがいは」といった言葉に、まず鉛筆で印をつける**ところから指導します。式を教える前に、読み方の型をつくります。'));

main.push(h3('「できた」を、お子様自身に自覚してもらう'));
main.push(body('低学年のお子様にとって最も効くのは、点数ではなく「先生が気づいてくれた」という実感です。'));
main.push(body('「先週は、ここで止まってたよね。今日は止まらなかったね」'));
main.push(body('同じ講師が続けて担当するからこそ言える一言です。当教室では、原則として**担当講師を固定**しております（詳細は同封の入会案内・料金表をご覧ください）。'));

main.push(h1('３．低学年のいま、個別指導に切り替える意味'));
main.push(body('「まだ2年生なのに、早すぎないか」というご相談もよくいただきます。事実を三つだけ、お伝えいたします。'));

main.push(h3('（1）小学3年生で、学習量と種類が明確に増える'));
main.push(body('学習指導要領上、3年生からは理科と社会が加わり、外国語活動も始まります。国語では説明文が長くなり、算数では割り算、小数、分数、そして「重さ」「時間」といった単位の扱いが登場します。'));
main.push(body('授業のコマ数が増え、扱う概念の抽象度も上がります。「3年生の壁」「9歳の壁」と呼ばれることがあるのは、この量と質の変化によるものです。'));

main.push(h3('（2）1・2年生の内容は、そのまま3年生以降の土台になる'));
main.push(bullet('繰り上がり・繰り下がり → 筆算、割り算の基礎'));
main.push(bullet('九九 → 割り算、分数、面積'));
main.push(bullet('音読と語彙 → 説明文・物語文の読解'));
main.push(bullet('漢字の意味理解 → 全教科の教科書を読む力'));
main.push(body('3年生でつまずくお子様の多くは、3年生の内容が難しいのではなく、**土台の一部が抜けたまま量が増えた**状態にあります。'));

main.push(h3('（3）低学年のうちは、修正に必要な時間が短い'));
main.push(body('これが最も現実的な理由です。2年生の内容の取りこぼしは、多くの場合、数週間から数か月で埋められます。同じ内容を5年生で埋め直そうとすると、当時より抽象度の高い説明を受け入れる必要があり、また「今の学年の勉強」と並行しなければならないため、負担は数倍になります。'));
main.push(body('さらに、低学年は**学習習慣そのものをつくり直せる時期**でもあります。学年が上がるほど、すでにある習慣（やり方の癖、姿勢、集中の切れ方）を書き換えるのは難しくなります。'));
main.push(body('とはいえ、「今始めなければ手遅れ」といったことは一切ございません。時間の余裕があるうちに手を打てば、必要な労力が少なくて済む――お伝えしたいのは、それだけです。'));

main.push(h1('４．見学・体験の流れ'));
main.push(h3('当日の流れ（所要時間 約90分）'));
main.push(body('**1. 受付・教室のご案内（約10分）**　教室内をご覧いただきます。座席の配置、静かさ、他の生徒さんの様子など、写真ではわからない部分をご確認ください。'));
main.push(body('**2. 現在の学習状況のヒアリング（約15分）**　保護者様からお話をうかがいます。学校でのご様子、ご家庭での宿題の進み方、気になっておられる点など。現在お使いの教材やプリントをお持ちいただけますと、より具体的にお話しできますが、必須ではございません。'));
main.push(body('**3. 学習の確認（約15分）**　お子様には、算数と国語の簡単な問題に取り組んでいただきます。**テストではございません。**点数はお出ししませんし、順位もつけません。講師が見ているのは、正誤よりも「どこで手が止まるか」です。'));
main.push(body('**4. 体験授業（約40分）**　実際の授業と同じ形で行います。保護者様は教室内でご見学いただいても、いったんお帰りいただいても構いません。お子様の緊張度合いによって、その場でお決めください。'));
main.push(body('**5. ご報告（約15分）**　体験中に講師が気づいた点を、その場で口頭にてお伝えします。「ここは十分できている」「ここは少し戻ったほうがよい」という具体的な内容です。'));

main.push(h3('お約束'));
main.push(callout([
  '**その場でご入塾をお決めいただくことは、一切ございません。**',
  '当日、申込書へのご記入をお願いすることはありません。「本日中にお決めいただければ」といった条件付きのご案内も一切ございません。体験当日は資料と口頭でのご報告をお持ち帰りいただき、必ず一度ご家庭でご検討くださいとお願いしております。',
  '体験後のご連絡もメールのみです。体験翌日にご報告書をメールでお送りし、そのあとは保護者様からのご連絡をお待ちいたします。催促のご連絡はいたしません。',
]));
main.push(spacer());

main.push(h3('モデルスケジュール（9月体験・10月入塾をご希望の場合）'));
const schedRows = [
  ['時期', '内容'],
  ['9月上旬〜中旬', '資料をご覧いただき、ご検討'],
  ['9月中旬', 'メールにて体験日をご調整（ご希望日を3つお送りいただくだけで完了）'],
  ['9月中旬〜下旬', '見学・体験授業（約90分）'],
  ['体験翌日', '当教室より、ご報告書と学習プランをメールで送付'],
  ['9月下旬〜月末', 'ご家庭でご検討。ご質問はメールでいつでも承ります'],
  ['9月末頃', 'ご入塾の可否をメールでご連絡'],
  ['10月上旬', 'ご入塾・授業開始'],
];
main.push(new Table({
  width: { size: 9070, type: WidthType.DXA },
  columnWidths: [2470, 6600],
  rows: schedRows.map((r, i) => new TableRow({
    children: r.map((c, j) => new TableCell({
      width: { size: j === 0 ? 2470 : 6600, type: WidthType.DXA },
      shading: i === 0 ? { type: ShadingType.CLEAR, fill: '1F3864' } : (i % 2 === 0 ? { type: ShadingType.CLEAR, fill: 'F2F5FA' } : undefined),
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [new Paragraph({
        children: [new TextRun({ text: c, bold: i === 0, font: i === 0 ? GOTHIC : MINCHO, size: BODY_SIZE, color: i === 0 ? 'FFFFFF' : undefined })],
        spacing: { after: 0 },
      })],
    })),
  })),
}));
main.push(spacer());
main.push(body('体験のご予約は、ご希望日の**3日前まで**にお申し込みいただければ調整可能です。9月中の平日夕方および土曜日に空き枠がございます。'));

main.push(h1('５．よくあるご質問'));
const faqs = [
  ['Q1. 体験だけ受けて、入塾しないという選択でも構いませんか。',
   'もちろんです。比較検討のために複数の教室を体験されるのは、当然のことと考えております。体験後にご入塾されなかった場合も、当教室からのご連絡は一切いたしません。お断りのご連絡すら不要です。なお、体験時にお伝えする学習面の所見は、他塾を選ばれる場合でもそのままお役立ていただけるよう、具体的にお伝えしております。'],
  ['Q2. しつこい勧誘や、電話はありませんか。',
   'ございません。今回のお問い合わせは「メールのみ・電話不可」として教室内で共有・記録しております。また、当教室では入塾をお決めいただくまでの間に、こちらからご連絡する回数を**原則2回まで**（資料送付のご連絡と、体験後のご報告）と定めております。それ以降は保護者様からのご連絡をお待ちする形になります。DMやキャンペーンのご案内も、ご希望がない限りお送りいたしません。'],
  ['Q3. 現在通っている教室と、しばらく併用することはできますか。',
   '可能です。実際に、数か月間併用されてからどちらかに絞られるご家庭も、そのまま長く併用されるご家庭もございます。併用される場合は、内容が重複しないように当教室側で調整いたします。たとえば計算の反復は現在の教材にお任せし、当教室では読解と文章題、および理解の確認に時間を使う、といった形です。お子様のご負担が過剰になっていると判断した場合は、その旨を率直にお伝えします。「両方続けてください」と申し上げることはありません。'],
  ['Q4. 女性講師を希望できますか。',
   '承ります。ご希望をお伝えいただければ、体験授業の段階から女性講師が担当いたします。低学年の女の子の場合、こうしたご希望は多くいただいており、特別なご要望ではございませんので、どうぞご遠慮なくお申し付けください。ご入塾後に担当変更をご希望される場合も、メールでご一報いただければ対応いたします。理由をお聞かせいただく必要はございません。'],
  ['Q5. 宿題はどのくらい出ますか。',
   '小学2年生の場合、**1回あたり15〜20分程度**を目安としております。学校の宿題と合わせても、机に向かう時間が長くなりすぎないよう設計します。量よりも「毎日同じ時間に開く」ことを重視しますので、ご家庭の生活リズム（習い事、就寝時刻など）を体験時にうかがったうえで決定します。宿題が終わらないことでお子様が叱られる状況を作らないことを、方針としております。'],
  ['Q6. 渦森台から通わせる場合、時間帯はどのようになりますか。',
   '低学年のお子様は、**16時台〜18時台**の枠をお選びいただくことが多くなっております。バスのお時間や送迎のご都合に合わせて、開始時刻を調整することも可能です。ご希望の時間帯を体験のお申し込み時にお書き添えください。授業の前後にお子様が教室内でお待ちいただくことも可能です。'],
];
for (const [q, a] of faqs) {
  main.push(new Paragraph({
    children: [new TextRun({ text: q, bold: true, font: GOTHIC, size: BODY_SIZE, color: '1F3864' })],
    spacing: { before: 200, after: 80 },
  }));
  main.push(body(a));
}

main.push(h1('６．ご連絡について'));
main.push(body('ここまでお読みいただき、ありがとうございました。'));
main.push(body('ご連絡は、下記のメールアドレスにて承ります。**お電話でのご連絡はいたしませんので、ご安心ください。**'));
main.push(callout([
  '**ご連絡先**　【メールアドレス】',
  '**担当**　教室長 【教室長名】',
  'いただいたメールには**24時間以内**にご返信いたします（日曜・祝日にいただいた場合は翌営業日となります）。ご返信は、確認のための最小限の回数にとどめます。',
]));
main.push(spacer());
main.push(h3('体験をご希望の場合、お書きいただくのは3つだけです'));
main.push(bullet('お子様のお名前（ふりがな）と学年'));
main.push(bullet('ご希望日を**第3希望まで**（例: 9月17日(木)17時頃／9月19日(土)午前／9月24日(木)17時頃）'));
main.push(bullet('ご希望があれば、女性講師希望・ご質問など'));
main.push(body('これだけで、日程調整は完了いたします。以下をそのままコピーしてお使いいただいても構いません。'));
main.push(callout([
  'お子様のお名前:　　　　　　　　　　（　　年生）',
  '第1希望:　　月　　日（　）　　時頃',
  '第2希望:　　月　　日（　）　　時頃',
  '第3希望:　　月　　日（　）　　時頃',
  'ご要望・ご質問:',
], { fill: 'FFFFFF' }));
main.push(spacer());
main.push(h3('体験のお申し込み以外でも、どうぞご遠慮なく'));
main.push(body('「料金についてもう少し詳しく」「他塾と比べて迷っている点がある」「今回は見送るが、来年また相談したい」――どのようなご用件でも、メールにて承ります。ご質問だけで終わっていただいて構いません。営業のご連絡につながることはございません。'));
main.push(body('そして、この資料をご覧になったうえで当教室が合わないとご判断された場合は、**そのままお読み捨てください。**ご連絡は不要です。'));
main.push(body('お子様にとって最も良い場所が見つかりますことを、願っております。', { after: 240 }));
main.push(body('【塾名】【教室名】', { align: AlignmentType.RIGHT, after: 0 }));
main.push(body('教室長 【教室長名】', { align: AlignmentType.RIGHT, after: 0 }));
main.push(body('【住所】', { align: AlignmentType.RIGHT, after: 0 }));
main.push(body('【メールアドレス】', { align: AlignmentType.RIGHT, after: 0 }));

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: MINCHO, size: BODY_SIZE } },
    },
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '・',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 240 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertMillimetersToTwip(22), bottom: convertMillimetersToTwip(20),
          left: convertMillimetersToTwip(23), right: convertMillimetersToTwip(23),
        },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: GOTHIC, size: 18, color: '808080' })],
        })],
      }),
    },
    children: [...coverLetter, ...main],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('資料送付_送付状・ご案内資料.docx', buf);
  console.log('written', buf.length);
});
