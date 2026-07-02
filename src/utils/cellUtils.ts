const DAYS = ['senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu', 'minggu'];
const DAY_SHORT = ['se', 'sl', 'ra', 'ka', 'ju', 'sa', 'mi'];
const HEADER_WORDS = ['jadwal', 'kelas', 'guru', 'ruang', 'istirahat', 'waktu', 'hari', 'jam'];

export function isSubjectCell(value: string): boolean {
  if (!value) return false;
  const val = value.trim().toLowerCase();
  if (/^[\d.\:\-\s]+$/.test(val)) return false;
  if (DAYS.includes(val) || DAY_SHORT.includes(val)) return false;
  if (HEADER_WORDS.some(w => val.includes(w))) return false;
  return true;
}

export function getCellText(v: unknown): string {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

export const DAYS_NAMES = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];

export function dayToIdx(day: string): number {
  return DAYS_NAMES.findIndex(d => d.toLowerCase() === day.toLowerCase());
}
