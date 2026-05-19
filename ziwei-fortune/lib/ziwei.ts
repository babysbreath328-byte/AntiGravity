import { astro } from 'iztro';
import type { BirthInput, FortuneResult } from './types';
import { getStarFortune, EMPTY_PALACE_FORTUNE } from './fortuneText';

export function calculateFortune(input: BirthInput): FortuneResult {
  const { year, month, day, gender } = input;

  const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  const genderStr = gender === 'male' ? 'male' : 'female';

  let astrolabe;
  try {
    // 時刻不明のため子時（index=0）で固定
    astrolabe = astro.bySolar(dateStr, 0, genderStr, true, 'zh-TW');
  } catch {
    throw new Error(`入力された日付（${dateStr}）は対応範囲外か、存在しない日付です。`);
  }

  const palaces = astrolabe.palaces;
  const mingGong = palaces.find((p) => p.name === '命宮');

  if (!mingGong) {
    throw new Error('命宮の計算に失敗しました');
  }

  const majorStars = mingGong.majorStars ?? [];
  const isEmptyPalace = majorStars.length === 0;
  const mainStarName = isEmptyPalace ? null : majorStars[0].name;
  const additionalStars = isEmptyPalace ? [] : majorStars.slice(1).map((s) => s.name);

  const fortune = mainStarName ? getStarFortune(mainStarName) : null;
  const textData = fortune ?? EMPTY_PALACE_FORTUNE;

  return {
    mainStar: mainStarName,
    additionalStars,
    guanluStars: palaces.find((p) => p.name === '官祿')?.majorStars?.map((s) => s.name) ?? [],
    personality: textData.personality,
    strengths: textData.strengths,
    weaknesses: textData.weaknesses,
    isEmptyPalace,
    mingGongBranch: mingGong.earthlyBranch ?? '',
    mingGongStem: mingGong.heavenlyStem ?? '',
  };
}
