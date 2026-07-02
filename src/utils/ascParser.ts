import { XMLParser } from 'fast-xml-parser';
import type { ASCData, Period, Day, Subject, Teacher, Clazz, ClassRoom, Lesson, Card } from '../types/asc';

export function parseASCXML(xmlData: string): ASCData {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
  const root = parser.parse(xmlData);
  const data = root.timetable || root;

  const periods: Period[] = [];
  if (data.periods?.period) {
    const list = Array.isArray(data.periods.period) ? data.periods.period : [data.periods.period];
    list.forEach((p: any) => periods.push({ period: p.period, name: p.name, starttime: p.starttime, endtime: p.endtime }));
  }

  const days: Day[] = [];
  if (data.daysdefs?.daysdef) {
    const list = Array.isArray(data.daysdefs.daysdef) ? data.daysdefs.daysdef : [data.daysdefs.daysdef];
    list.forEach((d: any) => {
      if (d.days && d.days.length > 0 && !d.days.includes('0') && d.days.split('1').length === 2)
        days.push({ day: d.days, name: d.name, short: d.short });
    });
  }

  const toRecord = <T>(key: string, fn: (x: any) => T): Record<string, T> => {
    const r: Record<string, T> = {};
    const items = data[key]?.[key.slice(0, -1)];
    if (items) {
      const list = Array.isArray(items) ? items : [items];
      list.forEach((x: any) => { r[x.id] = fn(x); });
    }
    return r;
  };

  const subjects = toRecord<Subject>('subjects', s => ({ id: s.id, name: s.name, short: s.short, color: s.color }));
  const teachers = toRecord<Teacher>('teachers', t => ({ id: t.id, name: t.name, short: t.short }));
  const classes = toRecord<Clazz>('classes', c => ({ id: c.id, name: c.name, short: c.short }));
  const classrooms = toRecord<ClassRoom>('classrooms', cr => ({ id: cr.id, name: cr.name, short: cr.short }));

  const lessons: Record<string, Lesson> = {};
  if (data.lessons?.lesson) {
    const list = Array.isArray(data.lessons.lesson) ? data.lessons.lesson : [data.lessons.lesson];
    list.forEach((l: any) => {
      lessons[l.id] = {
        id: l.id,
        subjectids: l.subjectid ? l.subjectid.split(',') : [],
        classids: l.classids ? l.classids.split(',') : [],
        teacherids: l.teacherids ? l.teacherids.split(',') : [],
        classroomids: l.classroomids ? l.classroomids.split(',') : [],
        periodspercard: parseInt(l.periodspercard || '1'),
        periodsperweek: parseFloat(l.periodsperweek || '1'),
      };
    });
  }

  const cards: Card[] = [];
  if (data.cards?.card) {
    const list = Array.isArray(data.cards.card) ? data.cards.card : [data.cards.card];
    list.forEach((c: any) => cards.push({ lessonid: c.lessonid, period: c.period, days: c.days }));
  }

  return { periods, days, subjects, teachers, classes, classrooms, lessons, cards };
}
