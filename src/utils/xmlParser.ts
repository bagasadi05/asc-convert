import ExcelJS from 'exceljs';
import { TEACHER_COLORS, SUBJECT_COLORS, SLATE_BORDER } from './colors';

const COL_WIDTH = 12;

function cleanText(text: string): string {
  if (!text) return '';
  return text.replace(/Trial version/gi, '').replace(/asc timetables/gi, '').trim();
}

export async function xmlToExcel(
  xmlFile: File,
  signal?: AbortSignal,
  onProgress?: (progress: number, message: string) => void,
): Promise<Blob> {
  onProgress?.(10, 'Membaca file XML...');
  if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
  const xmlText = await xmlFile.text();

  onProgress?.(20, 'Mem-parsing XML...');
  const xmlDoc = new DOMParser().parseFromString(xmlText, 'text/xml');

  const classes: Record<string, string> = {};
  const classList: string[] = [];
  xmlDoc.querySelectorAll('classes class').forEach(el => {
    const id = el.getAttribute('id') || '';
    const name = cleanText(el.getAttribute('name') || el.getAttribute('short') || '');
    if (id) { classes[id] = name; classList.push(id); }
  });

  const teacherNames: Record<string, string> = {};
  const teacherCodes: Record<string, string> = {};
  const teacherOrder: string[] = [];
  xmlDoc.querySelectorAll('teachers teacher').forEach(el => {
    const id = el.getAttribute('id') || '';
    const name = cleanText(el.getAttribute('name') || '');
    const code = cleanText(el.getAttribute('short') || '') || id;
    if (id) { teacherNames[id] = name || code; teacherCodes[id] = code; teacherOrder.push(id); }
  });

  const subjectNames: Record<string, string> = {};
  const subjectCodes: Record<string, string> = {};
  const subjectOrder: string[] = [];
  xmlDoc.querySelectorAll('subjects subject').forEach(el => {
    const id = el.getAttribute('id') || '';
    const name = cleanText(el.getAttribute('name') || '');
    const code = cleanText(el.getAttribute('short') || '') || name || id;
    if (id) { subjectNames[id] = name || code; subjectCodes[id] = code; subjectOrder.push(id); }
  });

  const lessons: Record<string, { classIds: string[]; teacherIds: string[]; subjectId: string }> = {};
  xmlDoc.querySelectorAll('lessons lesson').forEach(el => {
    const id = el.getAttribute('id') || '';
    lessons[id] = {
      classIds: (el.getAttribute('classids') || '').split(','),
      teacherIds: (el.getAttribute('teacherids') || '').split(','),
      subjectId: el.getAttribute('subjectid') || '',
    };
  });

  let numDays = 5;
  const cards = xmlDoc.querySelectorAll('cards card');
  cards.forEach(card => {
    const ds = card.getAttribute('days') || '';
    if (ds.length > numDays) numDays = ds.length;
  });

  const daysNames = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
  const maxPeriodPerDay: number[] = new Array(numDays).fill(0);

  onProgress?.(40, 'Menyusun grid jadwal...');

  const grid: Record<string, Record<number, Record<number, string[]>>> = {};
  const gridSubject: Record<string, Record<number, Record<number, string>>> = {};
  const classDayMaxPeriod: Record<string, Record<number, number>> = {};
  const usedTeacherIds = new Set<string>();
  const usedSubjectIds = new Set<string>();

  classList.forEach(cls => {
    grid[cls] = {}; gridSubject[cls] = {}; classDayMaxPeriod[cls] = {};
    for (let d = 0; d < numDays; d++) { grid[cls][d] = {}; gridSubject[cls][d] = {}; classDayMaxPeriod[cls][d] = 0; }
  });

  cards.forEach(card => {
    const lessonId = card.getAttribute('lessonid') || '';
    const period = parseInt(card.getAttribute('period') || '1') - 1;
    const daysStr = card.getAttribute('days') || '';
    const day = daysStr.indexOf('1');
    const lesson = lessons[lessonId];
    if (lesson && day >= 0 && day < numDays) {
      if (period + 1 > maxPeriodPerDay[day]) maxPeriodPerDay[day] = period + 1;
      const tids = lesson.teacherIds.filter(t => teacherNames[t]);
      tids.forEach(t => usedTeacherIds.add(t));
      lesson.classIds.forEach(cId => {
        if (grid[cId]?.[day] && tids.length > 0 && !grid[cId][day][period]) grid[cId][day][period] = tids;
        if (gridSubject[cId]?.[day]) {
          if (period + 1 > (classDayMaxPeriod[cId][day] || 0)) classDayMaxPeriod[cId][day] = period + 1;
          if (lesson.subjectId && subjectNames[lesson.subjectId]) {
            usedSubjectIds.add(lesson.subjectId);
            if (!gridSubject[cId][day][period]) gridSubject[cId][day][period] = lesson.subjectId;
          }
        }
      });
    }
  });

  for (let d = 0; d < numDays; d++) if (!maxPeriodPerDay[d]) maxPeriodPerDay[d] = 10;

  const finalTeacherIds = teacherOrder.filter(id => usedTeacherIds.has(id));
  finalTeacherIds.sort((a, b) => {
    const ca = teacherCodes[a] || a, cb = teacherCodes[b] || b;
    const na = parseInt(ca), nb = parseInt(cb);
    const aIsNum = /^\d+$/.test(ca), bIsNum = /^\d+$/.test(cb);
    if (aIsNum && bIsNum) return na - nb;
    if (aIsNum !== bIsNum) return aIsNum ? -1 : 1;
    return ca.localeCompare(cb, undefined, { numeric: true });
  });

  const teacherColor: Record<string, string> = {};
  finalTeacherIds.forEach((id, idx) => { teacherColor[id] = TEACHER_COLORS[idx % TEACHER_COLORS.length]; });

  const finalSubjectIds = subjectOrder.filter(id => usedSubjectIds.has(id));
  const subjectColor: Record<string, string> = {};
  finalSubjectIds.forEach((id, idx) => { subjectColor[id] = SUBJECT_COLORS[idx % SUBJECT_COLORS.length]; });

  onProgress?.(60, 'Membuat file Excel...');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Jadwal Mengajar Guru');

  const totalPeriodCols = maxPeriodPerDay.reduce((a, b) => a + b, 0);
  const totalCols = 2 + totalPeriodCols;

  const titleRow = ws.getRow(1);
  titleRow.getCell(1).value = 'JADWAL MENGAJAR GURU';
  ws.mergeCells(1, 1, 1, totalCols);
  titleRow.getCell(1).font = { name: 'Segoe UI', size: 14, bold: true };
  titleRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };

  const hTop = ws.getRow(3);
  const hBot = ws.getRow(4);
  hTop.getCell(1).value = 'No.';
  hTop.getCell(2).value = 'Kelas';
  ws.mergeCells(3, 1, 4, 1);
  ws.mergeCells(3, 2, 4, 2);

  let ci = 3;
  for (let d = 0; d < numDays; d++) {
    const start = ci;
    hTop.getCell(start).value = daysNames[d] || `Hari ${d + 1}`;
    for (let p = 0; p < maxPeriodPerDay[d]; p++) hBot.getCell(ci++).value = p + 1;
    if (ci - 1 > start) ws.mergeCells(3, start, 3, ci - 1);
  }

  [hTop, hBot].forEach(row => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.border = SLATE_BORDER;
    });
  });

  onProgress?.(80, 'Mewarnai grid...');
  let ri = 5;
  let no = 1;
  classList.forEach(classId => {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const row = ws.getRow(ri);
    row.getCell(1).value = no++;
    row.getCell(2).value = classes[classId];
    [row.getCell(1), row.getCell(2)].forEach(cell => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      cell.border = SLATE_BORDER;
    });

    let cIdx = 3;
    for (let d = 0; d < numDays; d++) {
      for (let p = 0; p < maxPeriodPerDay[d]; p++) {
        const tids = grid[classId]?.[d]?.[p] || [];
        const cell = row.getCell(cIdx);
        cell.value = tids.map(id => teacherCodes[id] || id).join('/');
        cell.font = { name: 'Segoe UI', size: 9, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = SLATE_BORDER;
        if (tids[0] && teacherColor[tids[0]])
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: teacherColor[tids[0]] } };
        cIdx++;
      }
    }
    ws.getRow(ri).height = 30;
    ri++;
  });

  ri += 2;
  ws.getRow(ri).getCell(1).value = 'DAFTAR GURU';
  ws.getRow(ri).getCell(1).font = { name: 'Segoe UI', size: 12, bold: true };
  ws.mergeCells(ri, 1, ri, totalCols);
  ri++;

  const legendStart = ri;
  const NUM_LEGEND_COLUMNS = 5;
  const itemsPerCol = Math.ceil(finalTeacherIds.length / NUM_LEGEND_COLUMNS) || 1;
  finalTeacherIds.forEach((id, i) => {
    const colGroup = Math.floor(i / itemsPerCol);
    const pos = i % itemsPerCol;
    const row = ws.getRow(legendStart + pos);
    const offset = colGroup * 2;
    row.getCell(1 + offset).value = teacherCodes[id] || id;
    row.getCell(1 + offset).font = { name: 'Segoe UI', size: 10, bold: true };
    row.getCell(1 + offset).alignment = { horizontal: 'right', vertical: 'middle' };
    row.getCell(2 + offset).value = teacherNames[id] || '';
    row.getCell(2 + offset).font = { name: 'Segoe UI', size: 10 };
    row.getCell(2 + offset).alignment = { horizontal: 'left', vertical: 'middle' };
  });

  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 14;
  for (let c = 3; c <= totalCols; c++) ws.getColumn(c).width = COL_WIDTH;
  ws.getRow(3).height = 20;
  ws.getRow(4).height = 18;

  onProgress?.(85, 'Membuat sheet Jadwal Kelas...');
  const ws2 = wb.addWorksheet('Jadwal Kelas (Mapel)');
  const totalCols2 = 1 + totalPeriodCols;

  const s2h1 = ws2.getRow(1);
  const s2h2 = ws2.getRow(2);
  s2h1.getCell(1).value = 'Kelas';
  ws2.mergeCells(1, 1, 2, 1);
  let c2 = 2;
  for (let d = 0; d < numDays; d++) {
    const start = c2;
    s2h1.getCell(start).value = daysNames[d] || `Hari ${d + 1}`;
    for (let p = 0; p < maxPeriodPerDay[d]; p++) s2h2.getCell(c2++).value = p + 1;
    if (c2 - 1 > start) ws2.mergeCells(1, start, 1, c2 - 1);
  }

  [s2h1, s2h2].forEach(row => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB0B0B0' } };
      cell.border = SLATE_BORDER;
    });
  });

  let r2 = 3;
  classList.forEach(classId => {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const row = ws2.getRow(r2);
    row.getCell(1).value = classes[classId];
    row.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true };
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB0B0B0' } };
    row.getCell(1).border = SLATE_BORDER;

    let cIdx = 2;
    for (let d = 0; d < numDays; d++) {
      const ownMax = classDayMaxPeriod[classId]?.[d] || 0;
      for (let p = 0; p < maxPeriodPerDay[d]; p++) {
        const cell = row.getCell(cIdx);
        const subjId = gridSubject[classId]?.[d]?.[p] || '';
        if (subjId) {
          cell.value = subjectCodes[subjId] || subjId;
          cell.font = { name: 'Segoe UI', size: 9, bold: true };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: subjectColor[subjId] || 'FFFFFFFF' } };
        } else if (p >= ownMax) {
          cell.value = 'X';
          cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF808080' } };
          cell.fill = { type: 'pattern', pattern: 'lightUp', fgColor: { argb: 'FFAFAFAF' }, bgColor: { argb: 'FFE5E5E5' } };
        }
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = SLATE_BORDER;
        cIdx++;
      }
    }
    ws2.getRow(r2).height = 16;
    r2++;
  });

  ws2.getColumn(1).width = 8;
  for (let c = 2; c <= totalCols2; c++) ws2.getColumn(c).width = COL_WIDTH;
  ws2.getRow(1).height = 16;
  ws2.getRow(2).height = 14;

  onProgress?.(95, 'Menyimpan file...');
  const buf = await wb.xlsx.writeBuffer();
  onProgress?.(100, 'Selesai!');
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
