'use client';

import { useState } from 'react';
import type { BirthInput } from '@/lib/types';

interface Props {
  onSubmit: (input: BirthInput) => void;
  isLoading: boolean;
}

export default function FortuneForm({ onSubmit, isLoading }: Props) {
  const currentYear = new Date().getFullYear();
  const [year, setYear] = useState(1990);
  const [month, setMonth] = useState(1);
  const [day, setDay] = useState(1);
  const [gender, setGender] = useState<'male' | 'female'>('male');
  const [error, setError] = useState('');

  const getDaysInMonth = (y: number, m: number) => new Date(y, m, 0).getDate();

  const years = Array.from({ length: currentYear - 1899 }, (_, i) => currentYear - i);
  const months = Array.from({ length: 12 }, (_, i) => i + 1);
  const maxDay = getDaysInMonth(year, month);
  const days = Array.from({ length: maxDay }, (_, i) => i + 1);

  const handleYearChange = (y: number) => {
    setYear(y);
    if (day > getDaysInMonth(y, month)) setDay(getDaysInMonth(y, month));
  };

  const handleMonthChange = (m: number) => {
    setMonth(m);
    if (day > getDaysInMonth(year, m)) setDay(getDaysInMonth(year, m));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    const date = new Date(year, month - 1, day);
    if (date.getMonth() !== month - 1) {
      setError('存在しない日付です。日付を確認してください。');
      return;
    }
    onSubmit({ year, month, day, gender });
  };

  const selectClass =
    'w-full bg-[#1a0d35] border border-[#4a2a7a] text-purple-100 rounded-lg px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-purple-500 focus:border-transparent appearance-none cursor-pointer';

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      {/* 生年月日 */}
      <div>
        <label className="block text-purple-300 text-sm font-medium mb-3 tracking-wider">
          ✦ 生年月日
        </label>
        <div className="grid grid-cols-3 gap-3">
          <div>
            <label className="block text-purple-400 text-xs mb-1">年</label>
            <select
              value={year}
              onChange={(e) => handleYearChange(Number(e.target.value))}
              className={selectClass}
            >
              {years.map((y) => (
                <option key={y} value={y}>{y}年</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-purple-400 text-xs mb-1">月</label>
            <select
              value={month}
              onChange={(e) => handleMonthChange(Number(e.target.value))}
              className={selectClass}
            >
              {months.map((m) => (
                <option key={m} value={m}>{m}月</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-purple-400 text-xs mb-1">日</label>
            <select
              value={day}
              onChange={(e) => setDay(Number(e.target.value))}
              className={selectClass}
            >
              {days.map((d) => (
                <option key={d} value={d}>{d}日</option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* 性別 */}
      <div>
        <label className="block text-purple-300 text-sm font-medium mb-3 tracking-wider">
          ✦ 性別
        </label>
        <div className="grid grid-cols-2 gap-3">
          {([['male', '男性 ♂'], ['female', '女性 ♀']] as const).map(([val, label]) => (
            <button
              key={val}
              type="button"
              onClick={() => setGender(val)}
              className={`py-3 rounded-lg border text-sm font-medium transition-all duration-200 ${
                gender === val
                  ? 'bg-purple-700 border-purple-500 text-white shadow-[0_0_15px_rgba(168,85,247,0.4)]'
                  : 'bg-[#1a0d35] border-[#4a2a7a] text-purple-300 hover:border-purple-500 hover:text-purple-200'
              }`}
            >
              {label}
            </button>
          ))}
        </div>
      </div>

      {error && (
        <div className="bg-red-900/30 border border-red-700 text-red-300 text-sm rounded-lg px-4 py-2">
          {error}
        </div>
      )}

      <button
        type="submit"
        disabled={isLoading}
        className="w-full py-4 rounded-xl font-bold text-white tracking-widest text-sm transition-all duration-300 relative overflow-hidden group disabled:opacity-60 disabled:cursor-not-allowed"
        style={{
          background: 'linear-gradient(135deg, #7c3aed, #4f46e5, #7c3aed)',
          backgroundSize: '200% 200%',
          boxShadow: '0 0 20px rgba(124, 58, 237, 0.5)',
        }}
      >
        <span className="relative z-10">
          {isLoading ? '鑑定中...' : '✦ 命盤を開く ✦'}
        </span>
        <div
          className="absolute inset-0 opacity-0 group-hover:opacity-100 transition-opacity duration-300"
          style={{ background: 'linear-gradient(135deg, #9333ea, #6366f1)' }}
        />
      </button>
    </form>
  );
}
