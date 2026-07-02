import ExcelJS from 'exceljs';
import { PASTEL_COLORS, THIN_BORDER } from './colors';
import { DAYS_NAMES, dayToIdx } from './cellUtils';

const COL_WIDTH = 12;

export async function subjectExcelToExcel(
  fileBuffer: ArrayBuffer,
  signal?: AbortSignal,
  onProgress?: (progress: number, message: string) => void,
): Promise<Blob> {
  onProgress?.(10, 'Membaca struktur file Excel...');
  if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);

  onProgress?.(30, 'Mengekstrak data mapel...');

  const subjectMap: Record<string, string> = {};
  const reverseMap: Record<string, string> = {};

  const mapelSheet = workbook.getWorksheet('Mata pelajaran');
  if (mapelSheet) {
    let codeCol = -1, nameCol = -1;
    mapelSheet.eachRow((row, rn) => {
      if (codeCol === -1) {
        row.eachCell((cell, cn) => {
          const v = (cell.value?.toString() || '').toLowerCase();
          if (v === 'kode') codeCol = cn;
          if (v === 'mata pelajaran') nameCol = cn;
        });
      } else if (rn > 1) {
        const code = row.getCell(codeCol).value?.toString()?.trim() || '';
        const name = row.getCell(nameCol).value?.toString()?.trim() || '';
        if (code && name) { subjectMap[code] = name; reverseMap[name.toLowerCase()] = code; }
      }
    });
  }

  onProgress?.(45, 'Mengekstrak kontrak guru...');
  const classSubjectTeacher: Record<string, Record<string, string>> = {};
  const classList = new Set<string>();

  const pelajaranSheet = workbook.getWorksheet('Pelajaran');
  if (pelajaranSheet) {
    let guruCol = -1, kelasCol = -1, mapelCol = -1;
    pelajaranSheet.eachRow((row, rn) => {
      if (guruCol === -1) {
        row.eachCell((cell, cn) => {
          const v = (cell.value?.toString() || '').toLowerCase();
          if (v === 'guru') guruCol = cn;
          if (v === 'kelas') kelasCol = cn;
          if (v === 'mata pelajaran') mapelCol = cn;
        });
      } else if (rn > 1) {
        const guru = row.getCell(guruCol).value?.toString()?.trim() || '';
        const kelasRaw = row.getCell(kelasCol).value?.toString()?.trim() || '';
        const mapelCode = row.getCell(mapelCol).value?.toString()?.trim() || '';
        if (guru && kelasRaw && mapelCode) {
          kelasRaw.split(',').map(k => k.trim()).forEach(k => {
            classList.add(k);
            if (!classSubjectTeacher[mapelCode]) classSubjectTeacher[mapelCode] = {};
            const prev = classSubjectTeacher[mapelCode][k];
            classSubjectTeacher[mapelCode][k] = prev?.includes(guru) ? prev : prev ? `${prev}, ${guru}` : guru;
          });
        }
      }
    });
  }

  const getCode = (sheetName: string) => {
    const clean = sheetName.replace('.', '').trim();
    if (reverseMap[clean.toLowerCase()]) return reverseMap[clean.toLowerCase()];
    for (const c in subjectMap) if (subjectMap[c].toLowerCase() === clean.toLowerCase()) return c;
    return clean;
  };

  onProgress?.(60, 'Membaca jadwal setiap mapel...');
  const grid: Record<string, Record<number, Record<number, string>>> = {};
  const sortedClasses = Array.from(classList).sort();
  sortedClasses.forEach(cls => { grid[cls] = {}; });

  let maxPeriod = 10;
  let maxDay = 4;

  workbook.eachSheet(sheet => {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    if (!sheet.name.endsWith('.')) return;
    const subjectName = sheet.name.replace('.', '').trim();
    const subjectCode = getCode(sheet.name);
    let currentDay = -1;

    sheet.eachRow(row => {
      const colB = (row.getCell(2).value?.toString() || '').trim();
      const colC = (row.getCell(3).value?.toString() || '').trim();
      if (colB && DAYS_NAMES.some(d => d.toLowerCase() === colB.toLowerCase())) {
        currentDay = dayToIdx(colB);
        if (currentDay > maxDay) maxDay = currentDay;
      }
      const period = parseInt(colC);
      if (!isNaN(period) && currentDay >= 0) {
        if (period > maxPeriod) maxPeriod = period;
        row.eachCell((cell, cn) => {
          if (cn >= 4) {
            const className = (cell.value?.toString() || '').trim();
            if (className && classList.has(className)) {
              const teacher = classSubjectTeacher[subjectCode]?.[className] || classSubjectTeacher[subjectName]?.[className] || '';
              const text = teacher ? `${subjectName}\n${teacher}` : subjectName;
              if (!grid[className]) grid[className] = {};
              if (!grid[className][currentDay]) grid[className][currentDay] = {};
              grid[className][currentDay][period] = grid[className][currentDay][period]
                ? `${grid[className][currentDay][period]} / ${text}` : text;
            }
          }
        });
      }
    });
  });

  onProgress?.(80, 'Merakit Excel Grid...');
  const out = new ExcelJS.Workbook();
  const ws = out.addWorksheet('Jadwal Kelas');

  const hr1 = ws.getRow(1);
  const hr2 = ws.getRow(2);
  hr1.getCell(1).value = 'Hari';
  hr2.getCell(1).value = 'Kelas';
  ws.mergeCells(1, 1, 2, 1);

  let colIdx = 2;
  for (let d = 0; d <= maxDay; d++) {
    const start = colIdx;
    hr1.getCell(start).value = DAYS_NAMES[d] || `Hari ${d + 1}`;
    for (let p = 1; p <= maxPeriod; p++) hr2.getCell(colIdx++).value = p;
    if (colIdx - 1 > start) ws.mergeCells(1, start, 1, colIdx - 1);
  }

  [hr1, hr2].forEach(row => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font = { name: 'Segoe UI', size: 11, bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.border = THIN_BORDER;
    });
  });

  let ri = 3;
  const subjectColors: Record<string, string> = {};
  let cc = 0;

  sortedClasses.forEach(cn => {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const row = ws.getRow(ri);
    const cls = row.getCell(1);
    cls.value = cn;
    cls.font = { name: 'Segoe UI', size: 11, bold: true };
    cls.alignment = { vertical: 'middle', horizontal: 'center' };
    cls.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    cls.border = THIN_BORDER;

    let ci = 2;
    for (let d = 0; d <= maxDay; d++) {
      for (let p = 1; p <= maxPeriod; p++) {
        const text = grid[cn]?.[d]?.[p] || '';
        const cell = row.getCell(ci);
        cell.value = text;
        cell.font = { name: 'Segoe UI', size: 10 };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = THIN_BORDER;
        if (text) {
          const sn = text.split('\n')[0].trim();
          if (!subjectColors[sn]) subjectColors[sn] = PASTEL_COLORS[cc++ % PASTEL_COLORS.length];
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: subjectColors[sn] } };
        }
        ci++;
      }
    }
    ri++;
  });

  ws.getColumn(1).width = 15;
  for (let c = 2; c < colIdx; c++) ws.getColumn(c).width = COL_WIDTH;

  onProgress?.(95, 'Menyimpan file...');
  const buf = await out.xlsx.writeBuffer();
  onProgress?.(100, 'Selesai!');
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
