import ExcelJS from 'exceljs';
import type { ASCData } from '../types/asc';

export async function generateExcel(data: ASCData): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ASC Converter';
  
  // Create Class Timetable Sheet
  const classSheet = workbook.addWorksheet('Jadwal Kelas');
  
  // Sort days and periods
  const days = data.days.sort((a, b) => a.day.indexOf('1') - b.day.indexOf('1'));
  const periods = data.periods.sort((a, b) => parseInt(a.period) - parseInt(b.period));
  
  let currentRow = 1;
  
  for (const classId of Object.keys(data.classes)) {
    const cls = data.classes[classId];
    
    // Class Header
    const titleCell = classSheet.getCell(`A${currentRow}`);
    titleCell.value = `Jadwal Kelas: ${cls.name}`;
    titleCell.font = { bold: true, size: 14 };
    currentRow++;
    
    // Periods Header
    let colIndex = 2;
    for (const p of periods) {
      const cell = classSheet.getCell(currentRow, colIndex);
      cell.value = `Jam ${p.period}\n${p.starttime} - ${p.endtime}`;
      cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      classSheet.getColumn(colIndex).width = 15;
      colIndex++;
    }
    currentRow++;
    
    // Days Rows
    for (const day of days) {
      const dayCell = classSheet.getCell(currentRow, 1);
      dayCell.value = day.name;
      dayCell.font = { bold: true };
      dayCell.alignment = { horizontal: 'center', vertical: 'middle' };
      dayCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      dayCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      classSheet.getColumn(1).width = 12;
      
      for (let i = 0; i < periods.length; i++) {
        const p = periods[i];
        const cell = classSheet.getCell(currentRow, i + 2);
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        // Find card for this class, day, and period
        const cards = data.cards.filter(c => c.days === day.day && c.period === p.period);
        const cellCards = cards.filter(c => {
          const lesson = data.lessons[c.lessonid];
          return lesson && lesson.classids.includes(classId);
        });
        
        if (cellCards.length > 0) {
          const c = cellCards[0]; // Take first if multiple
          const lesson = data.lessons[c.lessonid];
          const subject = data.subjects[lesson.subjectids[0]];
          const teacher = data.teachers[lesson.teacherids[0]];
          
          cell.value = `${subject?.short || ''}\n${teacher?.short || ''}`;
          cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
          
          if (subject && subject.color) {
            let color = subject.color.replace('#', '');
            if (color.length === 6) color = 'FF' + color;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
          }
        }
      }
      currentRow++;
    }
    
    currentRow += 2; // Space between classes
  }
  
  // Teacher Timetable Sheet
  const teacherSheet = workbook.addWorksheet('Jadwal Guru');
  currentRow = 1;
  
  for (const teacherId of Object.keys(data.teachers)) {
    const teacher = data.teachers[teacherId];
    
    // Teacher Header
    const titleCell = teacherSheet.getCell(`A${currentRow}`);
    titleCell.value = `Jadwal Guru: ${teacher.name}`;
    titleCell.font = { bold: true, size: 14 };
    currentRow++;
    
    // Periods Header
    let colIndex = 2;
    for (const p of periods) {
      const cell = teacherSheet.getCell(currentRow, colIndex);
      cell.value = `Jam ${p.period}\n${p.starttime} - ${p.endtime}`;
      cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      teacherSheet.getColumn(colIndex).width = 15;
      colIndex++;
    }
    currentRow++;
    
    // Days Rows
    for (const day of days) {
      const dayCell = teacherSheet.getCell(currentRow, 1);
      dayCell.value = day.name;
      dayCell.font = { bold: true };
      dayCell.alignment = { horizontal: 'center', vertical: 'middle' };
      dayCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      dayCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      teacherSheet.getColumn(1).width = 12;
      
      for (let i = 0; i < periods.length; i++) {
        const p = periods[i];
        const cell = teacherSheet.getCell(currentRow, i + 2);
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        
        // Find card for this teacher, day, and period
        const cards = data.cards.filter(c => c.days === day.day && c.period === p.period);
        const cellCards = cards.filter(c => {
          const lesson = data.lessons[c.lessonid];
          return lesson && lesson.teacherids.includes(teacherId);
        });
        
        if (cellCards.length > 0) {
          const c = cellCards[0];
          const lesson = data.lessons[c.lessonid];
          const subject = data.subjects[lesson.subjectids[0]];
          const cls = data.classes[lesson.classids[0]];
          
          cell.value = `${subject?.short || ''}\n${cls?.name || ''}`;
          cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
          
          if (subject && subject.color) {
            let color = subject.color.replace('#', '');
            if (color.length === 6) color = 'FF' + color;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
          }
        }
      }
      currentRow++;
    }
    
    currentRow += 2; // Space between teachers
  }
  
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
