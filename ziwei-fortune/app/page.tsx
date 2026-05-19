'use client';

import { useState } from 'react';
import FortuneForm from '@/components/FortuneForm';
import FortuneResult from '@/components/FortuneResult';
import type { BirthInput, FortuneResult as FortuneResultType, AnimalResult } from '@/lib/types';
import { calculateFortune } from '@/lib/ziwei';
import { getAnimalResult } from '@/lib/animalFortune';

type Phase = 'form' | 'result';

export default function Home() {
  const [phase, setPhase] = useState<Phase>('form');
  const [ziweiResult, setZiweiResult] = useState<FortuneResultType | null>(null);
  const [animalResult, setAnimalResult] = useState<AnimalResult | null>(null);
  const [error, setError] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleSubmit = async (input: BirthInput) => {
    setIsLoading(true);
    setError('');
    try {
      await new Promise((r) => setTimeout(r, 600));
      const ziwei = calculateFortune(input);
      const animal = getAnimalResult(input.year, input.month, input.day);
      setZiweiResult(ziwei);
      setAnimalResult(animal);
      setPhase('result');
    } catch (e) {
      setError(e instanceof Error ? e.message : '鑑定中にエラーが発生しました。');
    } finally {
      setIsLoading(false);
    }
  };

  const handleBack = () => {
    setPhase('form');
    setZiweiResult(null);
    setAnimalResult(null);
    setError('');
  };

  return (
    <div className="relative min-h-screen flex flex-col items-center justify-start py-12 px-4">
      <div className="stars-layer" aria-hidden="true" />
      <div
        aria-hidden="true"
        className="fixed top-0 left-1/2 -translate-x-1/2 w-[600px] h-[400px] rounded-full pointer-events-none"
        style={{ background: 'radial-gradient(ellipse, rgba(88,28,135,0.25) 0%, transparent 70%)', filter: 'blur(60px)', zIndex: 0 }}
      />

      <div className="relative z-10 w-full max-w-lg">
        {/* ヘッダー */}
        <header className="text-center mb-8">
          <p className="text-purple-500 text-xs tracking-[0.4em] mb-3">✦ ✦ ✦</p>
          <h1
            className="text-3xl font-bold tracking-[0.15em] mb-1"
            style={{ background: 'linear-gradient(135deg, #c084fc, #a78bfa, #818cf8)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}
          >
            星と動物の鑑定
          </h1>
          <p
            className="text-base tracking-widest"
            style={{ background: 'linear-gradient(135deg, #fbbf24, #f59e0b)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}
          >
            紫微斗数 × 動物占い
          </p>
          <p className="text-purple-500 text-xs tracking-wider mt-2">
            {phase === 'form' ? '生年月日から、あなたの本質と行動スタイルを読み解きます' : '星と動物が伝えるメッセージ'}
          </p>
        </header>

        {/* カード */}
        <div
          className="rounded-2xl border border-purple-800/40 p-5 backdrop-blur-sm"
          style={{ background: 'rgba(13,6,32,0.85)', boxShadow: '0 0 60px rgba(88,28,135,0.15), 0 4px 24px rgba(0,0,0,0.5)' }}
        >
          {phase === 'form' && (
            <div className="animate-fadeIn">
              <FortuneForm onSubmit={handleSubmit} isLoading={isLoading} />
              {error && (
                <div className="mt-4 bg-red-900/30 border border-red-700/50 text-red-300 text-sm rounded-lg px-4 py-3">
                  {error}
                </div>
              )}
            </div>
          )}

          {phase === 'result' && ziweiResult && animalResult && (
            <FortuneResult
              ziweiResult={ziweiResult}
              animalResult={animalResult}
              onBack={handleBack}
            />
          )}
        </div>

        <footer className="text-center mt-6">
          <p className="text-purple-700 text-xs tracking-wider">紫微斗数 × 動物占い ― 星の導きを受け取ってください</p>
        </footer>
      </div>
    </div>
  );
}
