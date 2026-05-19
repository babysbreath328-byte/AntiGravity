'use client';

import { useState } from 'react';
import type { FortuneResult, AnimalResult } from '@/lib/types';
import { STAR_FORTUNE, EMPTY_PALACE_FORTUNE } from '@/lib/fortuneText';
import { generatePptx } from '@/lib/generatePptx';

interface Props {
  ziweiResult: FortuneResult;
  animalResult: AnimalResult;
  onBack: () => void;
}

function ExportModal({
  onClose,
  onExport,
}: {
  onClose: () => void;
  onExport: (name: string) => Promise<void>;
}) {
  const [name, setName] = useState('');
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState('');

  const handleExport = async () => {
    if (!name.trim()) {
      setError('お名前を入力してください');
      return;
    }
    setIsExporting(true);
    setError('');
    try {
      await onExport(name.trim());
      onClose();
    } catch {
      setError('出力中にエラーが発生しました');
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center px-4"
      style={{ background: 'rgba(5,0,20,0.75)', backdropFilter: 'blur(4px)' }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        className="w-full max-w-sm rounded-2xl border border-amber-700/40 p-7"
        style={{ background: 'rgba(20,10,40,0.97)', boxShadow: '0 0 60px rgba(200,149,107,0.2)' }}
      >
        <p className="text-amber-400/80 text-xs tracking-[0.3em] mb-2 text-center">✦ PowerPoint 出力 ✦</p>
        <h2 className="text-white text-lg font-bold text-center mb-5 tracking-wider">
          お客様のお名前を入力
        </h2>
        <input
          type="text"
          value={name}
          onChange={(e) => { setName(e.target.value); setError(''); }}
          onKeyDown={(e) => { if (e.key === 'Enter') handleExport(); }}
          placeholder="例：山田 花子"
          autoFocus
          className="w-full bg-[#1a0d35] border border-[#4a2a7a] text-purple-100 rounded-lg px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-transparent placeholder:text-purple-700 mb-1"
        />
        <p className="text-purple-600 text-xs mb-1">
          ファイル名：<span className="text-amber-500/70">{name ? `${name}様_成功のロードマップ.pptx` : '○○様_成功のロードマップ.pptx'}</span>
        </p>
        {error && <p className="text-red-400 text-xs mt-1 mb-2">{error}</p>}
        <div className="flex gap-3 mt-5">
          <button
            onClick={onClose}
            className="flex-1 py-2.5 rounded-xl border border-purple-700/50 text-purple-400 text-sm tracking-wider hover:bg-purple-900/30 transition-all duration-200"
          >
            キャンセル
          </button>
          <button
            onClick={handleExport}
            disabled={isExporting}
            className="flex-1 py-2.5 rounded-xl text-white text-sm font-bold tracking-wider transition-all duration-200 disabled:opacity-60"
            style={{ background: 'linear-gradient(135deg, #C8956B, #D4A07A)', boxShadow: '0 0 15px rgba(200,149,107,0.3)' }}
          >
            {isExporting ? '生成中...' : '出力する'}
          </button>
        </div>
      </div>
    </div>
  );
}

function buildMemoText(ziwei: FortuneResult, animal: AnimalResult): string {
  const starLabel = ziwei.isEmptyPalace ? '空宮' : (ziwei.mainStar ?? '');
  return [
    '【紫微斗数の結果】',
    `主星：${starLabel}`,
    `命宮：${ziwei.mingGongStem}${ziwei.mingGongBranch}${ziwei.additionalStars.length ? '　' + ziwei.additionalStars.join('・') : ''}`,
    '',
    '▼ 性格・本質',
    ziwei.personality,
    '',
    '【動物占いの結果】',
    `タイプ：${animal.fullName}（No.${animal.characterNo}）`,
    `グループ：${animal.group} / ${animal.starType}`,
    '',
    '▼ 行動スタイル・性格',
    animal.personality,
    '',
    '【得意なことの種類と取り組み方】',
    '▼ 紫微斗数',
    ziwei.strengths,
    '',
    '▼ 動物占い',
    animal.strengths,
    '',
    '【苦手なことの種類と避け方】',
    '▼ 紫微斗数',
    ziwei.weaknesses,
    '',
    '▼ 動物占い',
    animal.weaknesses,
  ].join('\n');
}

export default function FortuneResult({ ziweiResult, animalResult, onBack }: Props) {
  const starData = ziweiResult.mainStar ? STAR_FORTUNE[ziweiResult.mainStar] : null;
  const ziweiData = starData ?? EMPTY_PALACE_FORTUNE;
  const [showModal, setShowModal] = useState(false);
  const [comment, setComment] = useState('');
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    await navigator.clipboard.writeText(buildMemoText(ziweiResult, animalResult));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleExport = async (clientName: string) => {
    await generatePptx(clientName, ziweiResult, animalResult, comment);
  };

  return (
    <>
      {showModal && (
        <ExportModal
          onClose={() => setShowModal(false)}
          onExport={handleExport}
        />
      )}

      <div className="space-y-5 animate-fadeIn">

        {/* ═══ 1枚目 ═══ */}
        <div
          className="rounded-2xl border border-purple-700/40 overflow-hidden"
          style={{ background: 'rgba(13, 6, 32, 0.9)', boxShadow: '0 0 30px rgba(124,58,237,0.1)' }}
        >
          <div className="flex items-center justify-between px-5 py-3 border-b border-purple-800/30">
            <span className="text-purple-500 text-[10px] tracking-[0.3em]">READING REPORT</span>
            <span className="text-purple-700 text-[10px]">1 / 2</span>
          </div>

          <div className="p-4 space-y-3">

            {/* 上段: 紫微斗数の結果 ｜ 動物占いの結果 */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">

              {/* 紫微斗数の結果 */}
              <div className="rounded-xl border border-purple-700/40 overflow-hidden">
                <div
                  className="px-4 py-2.5 border-b border-purple-800/30"
                  style={{ background: 'linear-gradient(135deg, rgba(124,58,237,0.25), rgba(88,28,220,0.15))' }}
                >
                  <p className="text-purple-400 text-[9px] tracking-[0.3em] mb-1">紫微斗数の結果</p>
                  <p
                    className="text-xl font-bold tracking-wider"
                    style={{ background: 'linear-gradient(135deg, #c084fc, #a78bfa)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}
                  >
                    {ziweiResult.isEmptyPalace ? '空　宮' : (ziweiResult.mainStar ?? '')}
                  </p>
                  <p className="text-purple-500/70 text-[10px] mt-0.5">
                    命宮 {ziweiResult.mingGongStem}{ziweiResult.mingGongBranch}
                    {ziweiResult.additionalStars.length ? '　' + ziweiResult.additionalStars.join('・') : ''}
                    {ziweiData.keyword ? `　${ziweiData.keyword}` : ''}
                  </p>
                </div>
                <div className="p-3" style={{ background: 'rgba(19, 10, 37, 0.5)' }}>
                  <p className="text-purple-100 text-xs leading-relaxed">{ziweiResult.personality}</p>
                </div>
              </div>

              {/* 動物占いの結果 */}
              <div className="rounded-xl border border-amber-800/40 overflow-hidden">
                <div
                  className="px-4 py-2.5 border-b border-amber-900/30"
                  style={{ background: 'linear-gradient(135deg, rgba(180,120,30,0.22), rgba(140,90,20,0.12))' }}
                >
                  <p className="text-amber-500/70 text-[9px] tracking-[0.3em] mb-1">動物占いの結果</p>
                  <p
                    className="text-xl font-bold tracking-wider"
                    style={{ background: 'linear-gradient(135deg, #fbbf24, #fcd34d)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}
                  >
                    {animalResult.fullName}
                  </p>
                  <p className="text-amber-600/60 text-[10px] mt-0.5">
                    No.{animalResult.characterNo}　{animalResult.group}　{animalResult.starType}
                  </p>
                </div>
                <div className="p-3" style={{ background: 'rgba(19, 10, 37, 0.5)' }}>
                  <p className="text-purple-100 text-xs leading-relaxed">{animalResult.personality}</p>
                </div>
              </div>
            </div>

            {/* 区切り */}
            <div className="flex items-center gap-3 py-1">
              <div className="flex-1 h-px bg-purple-800/20" />
              <span className="text-purple-700/60 text-[9px] tracking-widest">◈ ◈ ◈</span>
              <div className="flex-1 h-px bg-purple-800/20" />
            </div>

            {/* 下段: 得意なことの種類と取り組み方 ｜ 苦手なことの種類と避け方 */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">

              {/* 得意なことの種類と取り組み方 */}
              <div className="rounded-xl border border-emerald-800/30 overflow-hidden">
                <div className="px-4 py-2.5 border-b border-emerald-900/20" style={{ background: 'rgba(16, 50, 30, 0.45)' }}>
                  <p className="text-emerald-400 text-[11px] font-semibold tracking-wider">◉ 得意なことの種類と取り組み方</p>
                </div>
                <div className="p-3 space-y-3" style={{ background: 'rgba(19, 10, 37, 0.5)' }}>
                  <div>
                    <p className="text-emerald-600/80 text-[9px] tracking-[0.2em] mb-1.5">【紫微斗数】</p>
                    <p className="text-purple-100 text-xs leading-relaxed">{ziweiResult.strengths}</p>
                  </div>
                  <div className="h-px bg-emerald-900/25" />
                  <div>
                    <p className="text-emerald-600/80 text-[9px] tracking-[0.2em] mb-1.5">【動物占い】</p>
                    <p className="text-purple-100 text-xs leading-relaxed">{animalResult.strengths}</p>
                  </div>
                </div>
              </div>

              {/* 苦手なことの種類と避け方 */}
              <div className="rounded-xl border border-rose-800/30 overflow-hidden">
                <div className="px-4 py-2.5 border-b border-rose-900/20" style={{ background: 'rgba(50, 16, 20, 0.45)' }}>
                  <p className="text-rose-400 text-[11px] font-semibold tracking-wider">◇ 苦手なことの種類と避け方</p>
                </div>
                <div className="p-3 space-y-3" style={{ background: 'rgba(19, 10, 37, 0.5)' }}>
                  <div>
                    <p className="text-rose-600/80 text-[9px] tracking-[0.2em] mb-1.5">【紫微斗数】</p>
                    <p className="text-purple-100 text-xs leading-relaxed">{ziweiResult.weaknesses}</p>
                  </div>
                  <div className="h-px bg-rose-900/25" />
                  <div>
                    <p className="text-rose-600/80 text-[9px] tracking-[0.2em] mb-1.5">【動物占い】</p>
                    <p className="text-purple-100 text-xs leading-relaxed">{animalResult.weaknesses}</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* ═══ 2枚目 ═══ */}
        <div
          className="rounded-2xl border border-amber-800/30 overflow-hidden"
          style={{ background: 'rgba(13, 6, 32, 0.9)', boxShadow: '0 0 20px rgba(200,149,107,0.07)' }}
        >
          <div className="flex items-center justify-between px-5 py-3 border-b border-amber-900/30">
            <span className="text-amber-600/60 text-[10px] tracking-[0.3em]">COMMENT</span>
            <span className="text-purple-700 text-[10px]">2 / 2</span>
          </div>
          <div className="p-4">
            <p className="text-amber-500/70 text-xs tracking-[0.25em] mb-3">✦ 私からのコメント</p>
            <textarea
              value={comment}
              onChange={(e) => setComment(e.target.value)}
              placeholder="お客様へのコメントを入力してください..."
              rows={10}
              className="w-full bg-[#0d0620]/60 border border-purple-800/30 text-purple-100 rounded-xl px-4 py-3 text-sm leading-relaxed focus:outline-none focus:ring-1 focus:ring-amber-600/40 focus:border-amber-700/40 placeholder:text-purple-800 resize-none"
            />
          </div>
        </div>

        {/* テキストコピーボタン */}
        <button
          onClick={handleCopy}
          className="w-full py-3.5 rounded-xl text-sm font-bold tracking-widest transition-all duration-200 hover:opacity-90 border border-purple-600/50"
          style={{ background: 'rgba(88,28,220,0.15)', color: copied ? '#86efac' : '#c084fc' }}
        >
          {copied ? '✓ コピーしました' : '📋 テキストをコピーする（AI用メモ）'}
        </button>

        {/* PowerPoint出力ボタン */}
        <button
          onClick={() => setShowModal(true)}
          className="w-full py-3.5 rounded-xl text-sm font-bold tracking-widest transition-all duration-200 hover:opacity-90"
          style={{ background: 'linear-gradient(135deg, #C8956B, #B8849A)', color: '#FFF8F0', boxShadow: '0 0 20px rgba(200,149,107,0.25)' }}
        >
          ✦ PowerPointで出力する ✦
        </button>

        {/* 戻るボタン */}
        <button
          onClick={onBack}
          className="w-full py-3 rounded-xl border border-purple-700/50 text-purple-400 text-sm tracking-widest hover:bg-purple-900/30 hover:text-purple-300 hover:border-purple-600 transition-all duration-200"
        >
          ← もう一度鑑定する
        </button>
      </div>
    </>
  );
}
