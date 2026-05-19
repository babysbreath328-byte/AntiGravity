import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: '紫微斗数 命盤鑑定 | 星の導き',
  description: '生年月日・出生時刻・性別から命宮の主星を算出し、あなたの本質と目標達成への道を鑑定します。',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja" className="h-full">
      <body className="min-h-full">{children}</body>
    </html>
  );
}
