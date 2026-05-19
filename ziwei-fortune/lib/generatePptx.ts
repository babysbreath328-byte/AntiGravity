import pptxgen from 'pptxgenjs';
import type { FortuneResult, AnimalResult } from './types';

// 小春日和カラーパレット
const C = {
  cream:     'FFF8F0',
  gold:      'C8956B',
  goldLight: 'F0D5B5',
  rose:      'D4899A',
  roseLight: 'F7DDE4',
  lavender:  'A896C8',
  lavLight:  'EAE4F5',
  brown:     '4A3020',
  brownMid:  '8B6550',
  white:     'FFFFFF',
  border:    'E0C8A8',
  green:     'F0F5EC',
  pinkLight: 'FDF0F3',
};

const FONT_FALLBACK = 'MS Mincho';

function addSlide1_Title(prs: pptxgen, clientName: string) {
  const slide = prs.addSlide();
  slide.background = { color: C.cream };

  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.18, fill: { color: C.gold }, line: { type: 'none' } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 7.32, w: '100%', h: 0.18, fill: { color: C.gold }, line: { type: 'none' } });
  slide.addShape(prs.ShapeType.rect, { x: 2.5, y: 2.6, w: 5, h: 0.02, fill: { color: C.gold }, line: { type: 'none' } });
  slide.addShape(prs.ShapeType.rect, { x: 2.5, y: 5.1, w: 5, h: 0.02, fill: { color: C.gold }, line: { type: 'none' } });

  slide.addText('✦ 星と動物が伝えるメッセージ ✦', {
    x: 1, y: 2.1, w: 8, h: 0.4,
    align: 'center', fontSize: 11, color: C.brownMid, fontFace: FONT_FALLBACK, charSpacing: 3,
  });
  slide.addText(`${clientName}さんの\n成功ロードマップ`, {
    x: 0.8, y: 2.75, w: 8.4, h: 2.1,
    align: 'center', fontSize: 34, bold: true, color: C.brown, fontFace: FONT_FALLBACK, lineSpacingMultiple: 1.4,
  });
  slide.addText('紫微斗数 × 動物占い（個性心理学）', {
    x: 1, y: 5.2, w: 8, h: 0.4,
    align: 'center', fontSize: 12, color: C.gold, fontFace: FONT_FALLBACK, charSpacing: 2,
  });

  const today = new Date();
  const dateStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
  slide.addText(dateStr, {
    x: 1, y: 6.6, w: 8, h: 0.35,
    align: 'center', fontSize: 10, color: C.brownMid, fontFace: FONT_FALLBACK,
  });
}

function addContentBlock(
  slide: pptxgen.Slide,
  prs: pptxgen,
  label: string,
  content: string,
  x: number, y: number, w: number, h: number,
  bgColor: string,
) {
  slide.addShape(prs.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: bgColor },
    line: { color: C.border, width: 0.5 },
    rectRadius: 0.08,
  });
  slide.addText(label, {
    x: x + 0.18, y: y + 0.1, w: w - 0.36, h: 0.28,
    fontSize: 8.5, bold: true, color: C.brownMid, fontFace: FONT_FALLBACK, charSpacing: 1,
  });
  slide.addText(content, {
    x: x + 0.18, y: y + 0.38, w: w - 0.36, h: h - 0.48,
    fontSize: 9, color: C.brown, fontFace: FONT_FALLBACK, lineSpacingMultiple: 1.45, wrap: true,
  });
}

function addSlide2_Combined(prs: pptxgen, ziwei: FortuneResult, animal: AnimalResult) {
  const slide = prs.addSlide();
  slide.background = { color: C.cream };

  // ヘッダーバー（左=紫微、右=動物）
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 5, h: 0.65, fill: { color: C.lavender }, line: { type: 'none' } });
  slide.addShape(prs.ShapeType.rect, { x: 5, y: 0, w: 5, h: 0.65, fill: { color: C.rose }, line: { type: 'none' } });

  // ヘッダー: 紫微斗数の結果
  slide.addText('紫微斗数の結果', {
    x: 0.25, y: 0.08, w: 4.5, h: 0.25,
    fontSize: 8.5, color: C.white, fontFace: FONT_FALLBACK, charSpacing: 2,
  });
  const starLabel = ziwei.isEmptyPalace ? '空　宮' : (ziwei.mainStar ?? '');
  slide.addText(starLabel, {
    x: 0.25, y: 0.32, w: 4.5, h: 0.28,
    fontSize: 16, bold: true, color: C.white, fontFace: FONT_FALLBACK,
  });

  // ヘッダー: 動物占いの結果
  slide.addText('動物占いの結果', {
    x: 5.1, y: 0.08, w: 4.65, h: 0.25,
    fontSize: 8.5, color: C.white, fontFace: FONT_FALLBACK, charSpacing: 2,
  });
  slide.addText(animal.fullName, {
    x: 5.1, y: 0.32, w: 4.65, h: 0.28,
    fontSize: 14, bold: true, color: C.white, fontFace: FONT_FALLBACK,
  });

  const COL_W = 4.65;
  const COL_L = 0.25;
  const COL_R = COL_L + COL_W + 0.2;

  // 上段: 性格 / 行動スタイル
  const TOP_Y = 0.78;
  const TOP_H = 2.3;
  addContentBlock(slide, prs,
    `◈ 命宮 ${ziwei.mingGongStem}${ziwei.mingGongBranch}　本質・性格`, ziwei.personality,
    COL_L, TOP_Y, COL_W, TOP_H, C.lavLight);
  addContentBlock(slide, prs,
    `◈ No.${animal.characterNo}　${animal.group}　${animal.starType}　行動スタイル`, animal.personality,
    COL_R, TOP_Y, COL_W, TOP_H, C.roseLight);

  // 区切り線
  slide.addShape(prs.ShapeType.rect, {
    x: COL_L, y: TOP_Y + TOP_H + 0.1, w: 9.5, h: 0.015,
    fill: { color: C.border }, line: { type: 'none' },
  });

  // 下段: 得意 / 苦手
  const BOT_Y = TOP_Y + TOP_H + 0.23;
  const BOT_H = 7.15 - BOT_Y;

  const strengthsContent = `【紫微斗数】\n${ziwei.strengths}\n\n【動物占い】\n${animal.strengths}`;
  const weaknessesContent = `【紫微斗数】\n${ziwei.weaknesses}\n\n【動物占い】\n${animal.weaknesses}`;

  addContentBlock(slide, prs, '◉ 得意なことの種類と取り組み方', strengthsContent,
    COL_L, BOT_Y, COL_W, BOT_H, C.green);
  addContentBlock(slide, prs, '◇ 苦手なことの種類と避け方', weaknessesContent,
    COL_R, BOT_Y, COL_W, BOT_H, C.pinkLight);

  // フッター
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 7.32, w: '100%', h: 0.18, fill: { color: C.gold }, line: { type: 'none' } });
  slide.addText('紫微斗数 × 動物占い ─ 成功ロードマップ', {
    x: 0, y: 7.3, w: '100%', h: 0.2,
    align: 'center', fontSize: 7, color: C.white, fontFace: FONT_FALLBACK,
  });
}

function addSlide3_Comment(prs: pptxgen, clientName: string, comment: string) {
  const slide = prs.addSlide();
  slide.background = { color: C.cream };

  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.65, fill: { color: C.gold }, line: { type: 'none' } });
  slide.addText('私からのコメント', {
    x: 0.3, y: 0.08, w: 7, h: 0.25,
    fontSize: 9, color: C.white, fontFace: FONT_FALLBACK, charSpacing: 2,
  });
  slide.addText(`${clientName}さんへ`, {
    x: 0.3, y: 0.32, w: 7, h: 0.28,
    fontSize: 16, bold: true, color: C.white, fontFace: FONT_FALLBACK,
  });

  // コメントエリア
  slide.addShape(prs.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 9.4, h: 6.3,
    fill: { color: C.white },
    line: { color: C.border, width: 1 },
    rectRadius: 0.1,
  });

  if (comment.trim()) {
    slide.addText(comment, {
      x: 0.5, y: 0.95, w: 9, h: 6.1,
      fontSize: 11, color: C.brown, fontFace: FONT_FALLBACK,
      lineSpacingMultiple: 1.6, wrap: true, valign: 'top',
    });
  } else {
    slide.addText('ここにコメントを記入してください', {
      x: 0.5, y: 1.1, w: 9, h: 0.4,
      fontSize: 10, color: 'D0B898', fontFace: FONT_FALLBACK, italic: true,
    });
    for (let i = 0; i < 10; i++) {
      slide.addShape(prs.ShapeType.rect, {
        x: 0.5, y: 1.65 + i * 0.46, w: 8.9, h: 0.01,
        fill: { color: C.goldLight }, line: { type: 'none' },
      });
    }
  }

  slide.addShape(prs.ShapeType.rect, { x: 0, y: 7.32, w: '100%', h: 0.18, fill: { color: C.gold }, line: { type: 'none' } });
  slide.addText('紫微斗数 × 動物占い ─ 成功ロードマップ', {
    x: 0, y: 7.3, w: '100%', h: 0.2,
    align: 'center', fontSize: 7, color: C.white, fontFace: FONT_FALLBACK,
  });
}

export async function generatePptx(
  clientName: string,
  ziweiResult: FortuneResult,
  animalResult: AnimalResult,
  comment = '',
) {
  const prs = new pptxgen();
  prs.layout = 'LAYOUT_WIDE';

  addSlide1_Title(prs, clientName);
  addSlide2_Combined(prs, ziweiResult, animalResult);
  addSlide3_Comment(prs, clientName, comment);

  await prs.writeFile({ fileName: `${clientName}様_成功のロードマップ.pptx` });
}
