export interface Period {
  period: string;
  name: string;
  starttime: string;
  endtime: string;
}

export interface Day {
  day: string;
  name: string;
  short: string;
}

export interface Subject {
  id: string;
  name: string;
  short: string;
  color: string;
}

export interface Teacher {
  id: string;
  name: string;
  short: string;
}

export interface ClassRoom {
  id: string;
  name: string;
  short: string;
}

export interface Clazz {
  id: string;
  name: string;
  short: string;
}

export interface Lesson {
  id: string;
  subjectids: string[];
  classids: string[];
  teacherids: string[];
  classroomids: string[];
  periodspercard: number;
  periodsperweek: number;
}

export interface Card {
  lessonid: string;
  period: string;
  days: string;
}

export interface ASCData {
  periods: Period[];
  days: Day[];
  subjects: Record<string, Subject>;
  teachers: Record<string, Teacher>;
  classes: Record<string, Clazz>;
  classrooms: Record<string, ClassRoom>;
  lessons: Record<string, Lesson>;
  cards: Card[];
}
