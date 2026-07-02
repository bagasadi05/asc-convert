import ExcelJS from 'exceljs';
import { PASTEL_COLORS, THIN_BORDER } from './colors';
import { isSubjectCell } from './cellUtils';

const COL_WIDTH = 12;

export async function beautifyExcel(
  fileBuffer: ArrayBuffer,
  signal?: AbortSignal,
  onProgress?: (progress: number, message: string) => void
): Promise<Blob> {
  onProgress?.(5, 'Membaca file Excel...');
  if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);

  const subjectColors: Record<string, string> = {};
  let colorIndex = 0;
  const totalSheets = workbook.worksheets.length;

  for (let si = 0; si < totalSheets; si++) {
    if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    const sheet = workbook.worksheets[si];
    onProgress?.(10 + (si / totalSheets) * 10, `Menganalisa sheet ${si + 1} dari ${totalSheets}...`);

    let actualMaxRow = 0;
    let actualMaxCol = 0;
    const rowData: { rowNum: number; cells: { col: number; val: string }[] }[] = [];

    sheet.eachRow((row, rn) => {
      const cells: { col: number; val: string }[] = [];
      row.eachCell((cell, cn) => {
        const v = cell.value?.toString() || '';
        if (v) {
          cells.push({ col: cn, val: v });
          if (cn > actualMaxCol) actualMaxCol = cn;
        }
      });
      if (cells.length) {
        rowData.push({ rowNum: rn, cells });
        if (rn > actualMaxRow) actualMaxRow = rn;
      }
    });

    if (!actualMaxRow || !actualMaxCol) continue;

    for (const rd of rowData) {
      for (const { col: cn, val: strValue } of rd.cells) {
        const cell = sheet.getCell(rd.rowNum, cn);
        cell.font = { name: 'Segoe UI', size: 11, bold: cell.font?.bold || false };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = THIN_BORDER;

        if (isSubjectCell(strValue)) {
          const sn = strValue.split('\n')[0].trim();
          if (!subjectColors[sn]) subjectColors[sn] = PASTEL_COLORS[colorIndex++ % PASTEL_COLORS.length];
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: subjectColors[sn] } };
        } else if (strValue.toLowerCase().includes('istirahat')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
          cell.font = { ...cell.font, italic: true, color: { argb: 'FF64748B' } };
        } else if (cell.font?.bold) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        }
      }
    }

    onProgress?.(70 + (si / totalSheets) * 10, 'Menyesuaikan lebar kolom...');

    for (let c = 1; c <= actualMaxCol; c++) {
      sheet.getColumn(c).width = COL_WIDTH;
    }
  }

  onProgress?.(90, 'Menyimpan file...');
  const buf = await workbook.xlsx.writeBuffer();
  onProgress?.(100, 'Selesai!');
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
