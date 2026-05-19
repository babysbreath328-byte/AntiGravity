export interface BirthInput {
  year: number;
  month: number;
  day: number;
  gender: 'male' | 'female';
}

export interface FortuneResult {
  mainStar: string | null;
  additionalStars: string[];
  guanluStars: string[];
  personality: string;
  strengths: string;
  weaknesses: string;
  isEmptyPalace: boolean;
  mingGongBranch: string;
  mingGongStem: string;
}

export interface AnimalResult {
  animalName: string;
  fullName: string;
  personality: string;
  strengths: string;
  weaknesses: string;
  characterNo: number;
  group: string;
  starType: string;
}
