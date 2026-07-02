import ExcelJS from 'exceljs';
import { PASTEL_COLORS, THIN_BORDER } from './colors';
import { isSubjectCell } from './cellUtils';

const COL_WIDTH = 12;

export async function htmlToExcel(
  files: File[],
  signal?: AbortSignal,
  onProgress?: (progress: number, message: string) => void,
): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  const subjectColors: Record<string, string> = {};
  let colorIndex = 0;
  let tableCount = 0;

  for (let f = 0; f < files.length; f++) {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const file = files[f];
    onProgress?.(10 + (f / files.length) * 40, `Membaca file ${file.name}...`);

    const htmlText = await file.text();
    const doc = new DOMParser().parseFromString(htmlText, 'text/html');
    const tables = doc.querySelectorAll('table');

    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      tableCount++;

      let sheetName = `Jadwal ${tableCount}`;
      let prev = table.previousElementSibling;
      while (prev && prev.tagName !== 'TABLE') {
        const txt = prev.textContent?.trim();
        if (txt && txt.length < 50) {
          sheetName = txt.replace(/[\\\/\?\*\[\]]/g, '').substring(0, 31);
          break;
        }
        prev = prev.previousElementSibling;
      }
      if (workbook.getWorksheet(sheetName)) sheetName = `${sheetName} (${tableCount})`.substring(0, 31);

      const sheet = workbook.addWorksheet(sheetName);
      const occupied = new Set<string>();
      let maxCol = 1;
      let maxRow = 1;
      const rows = table.querySelectorAll('tr');

      for (let r = 0; r < rows.length; r++) {
        if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
        const tr = rows[r];
        const cells = tr.querySelectorAll('td, th');
        let c = 0;
        for (const cell of cells) {
          while (occupied.has(`${r},${c}`)) c++;
          const rowSpan = parseInt(cell.getAttribute('rowspan') || '1');
          const colSpan = parseInt(cell.getAttribute('colspan') || '1');
          const clone = cell.cloneNode(true) as HTMLElement;
          clone.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
          let text = (clone.textContent?.trim() || '').replace(/Trial version/gi, '').replace(/asc timetables/gi, '').trim();

          const trgRow = r + 1;
          const trgCol = c + 1;
          const tc = sheet.getCell(trgRow, trgCol);
          tc.value = text;
          tc.font = { name: 'Segoe UI', size: 11, bold: cell.tagName === 'TH' };
          tc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
          tc.border = THIN_BORDER;

          if (isSubjectCell(text)) {
            const sn = text.split('\n')[0].trim();
            if (!subjectColors[sn]) subjectColors[sn] = PASTEL_COLORS[colorIndex++ % PASTEL_COLORS.length];
            tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: subjectColors[sn] } };
          } else if (text.toLowerCase().includes('istirahat') || text.toLowerCase().includes('rehat')) {
            tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
            tc.font = { ...tc.font, italic: true, color: { argb: 'FF64748B' } };
          } else if (tc.font?.bold || r === 0 || c === 0) {
            tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            tc.font = { ...tc.font, bold: true };
          }

          if (rowSpan > 1 || colSpan > 1) {
            sheet.mergeCells(trgRow, trgCol, trgRow + rowSpan - 1, trgCol + colSpan - 1);
            for (let i = 0; i < rowSpan; i++)
              for (let j = 0; j < colSpan; j++)
                occupied.add(`${r + i},${c + j}`);
          } else {
            occupied.add(`${r},${c}`);
          }
          if (trgCol + colSpan - 1 > maxCol) maxCol = trgCol + colSpan - 1;
          c += colSpan;
        }
        if (r + 1 > maxRow) maxRow = r + 1;
      }

      for (let c = 1; c <= maxCol; c++) {
        sheet.getColumn(c).width = COL_WIDTH;
      }
    }
  }

  if (!tableCount) throw new Error('Tidak ditemukan tabel di file HTML.');
  onProgress?.(90, 'Menyimpan file...');
  const buf = await workbook.xlsx.writeBuffer();
  onProgress?.(100, 'Selesai!');
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
