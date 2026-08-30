/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  GraduationCap,
  LayoutGrid,
  Building2,
  Fingerprint,
  BarChart3,
  Users,
  FileText,
  ClipboardList,
  Settings2,
  Plus, 
  Pencil,
  Trash2, 
  Calendar, 
  Check, 
  Save,
  AlertCircle,
  AlertTriangle,
  Info,
  User as UserIcon,
  ArrowRight,
  Eye,
  EyeOff,
  LogOut,
  X,
  Camera,
  Key,
  Lock,
  Download,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CheckCircle,
  Cloud,
  Database,
  Activity,
  RefreshCw,
  Copy,
  UserX,
  Power,
  Loader2,
  Send,
  Sparkles,
  Clock,
  Timer
} from 'lucide-react';
import { 
  format, 
  addMonths, 
  subMonths, 
  startOfMonth, 
  endOfMonth, 
  startOfWeek, 
  endOfWeek, 
  eachDayOfInterval, 
  isSameMonth, 
  isSameDay,
  parseISO 
} from 'date-fns';
import { id as idLocale } from 'date-fns/locale';
import { 
  collection, 
  doc, 
  setDoc, 
  deleteDoc, 
  writeBatch,
  onSnapshot,
  getDocs,
  getDoc,
  query,
  where,
  deleteField,
  Firestore,
  DocumentData
} from 'firebase/firestore';
import { 
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signInAnonymously,
  signOut, 
  onAuthStateChanged,
  User,
  updateProfile,
  Auth
} from 'firebase/auth';
import { dbDefault, auth, OperationType, handleFirestoreError } from './firebase';
import { importExcelHelper } from './importExcelHelper';
import { importTeacherExcelHelper, normalizeDutyDay, getDayIndex } from './importTeacherExcelHelper';

import ReportsView from './ReportsView';

type Status = 'Hadir' | 'Sakit' | 'Izin' | 'Alpa' | 'Dispen' | '';

interface Student {
  id: string;
  nisn: string;
  name: string;
  class: string;
  waliKelas?: string;
  waliKelasNiy?: string;
  userId?: string;
}

export function getWaliKelasForClass(
  className: string,
  classWaliMap: Record<string, string>,
  students: Student[]
): string {
  if (!className) return '';
  if (classWaliMap && classWaliMap[className] && classWaliMap[className].trim() !== '') {
    return classWaliMap[className].trim();
  }
  const studentWithWali = students.find(s => s.class === className && s.waliKelas && s.waliKelas.trim() !== '');
  if (studentWithWali?.waliKelas) {
    return studentWithWali.waliKelas.trim();
  }
  return '';
}

export function getWaliKelasNiyForClass(
  className: string,
  classWaliNiyMap: Record<string, string>,
  students: Student[],
  teachers?: Teacher[],
  waliName?: string
): string {
  if (!className) return '';
  if (classWaliNiyMap && classWaliNiyMap[className] && classWaliNiyMap[className].trim() !== '') {
    return classWaliNiyMap[className].trim();
  }
  const studentWithWaliNiy = students.find(s => s.class === className && s.waliKelasNiy && s.waliKelasNiy.trim() !== '');
  if (studentWithWaliNiy?.waliKelasNiy) {
    return studentWithWaliNiy.waliKelasNiy.trim();
  }
  if (teachers && waliName && waliName.trim()) {
    const matched = teachers.find(t => t.name.trim().toLowerCase() === waliName.trim().toLowerCase());
    if (matched && matched.niy && matched.niy.trim() !== '-' && matched.niy.trim() !== '') {
      return matched.niy.trim();
    }
  }
  return '';
}

export function compareClass(a: string, b: string): number {
  const romanToInt = (s: string) => {
    const rom: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
    let num = 0;
    for (let i = 0; i < s.length; i++) {
      if (rom[s[i].toLowerCase()] < (rom[s[i + 1]?.toLowerCase()] || 0)) {
        num -= rom[s[i].toLowerCase()];
      } else {
        num += rom[s[i].toLowerCase()];
      }
    }
    return num;
  };
  const getParts = (str: string) => {
    const match = (str || '').trim().match(/^([IVXLCDMivxlcdm]+)(.*)$/);
    if (match) {
      return { num: romanToInt(match[1]), suffix: match[2].trim() };
    }
    return { num: 0, suffix: (str || '').trim() };
  };
  const partA = getParts(a);
  const partB = getParts(b);
  if (partA.num !== partB.num) {
    return partA.num - partB.num;
  }
  return partA.suffix.localeCompare(partB.suffix, 'id-ID', { numeric: true });
}

export function sortStudents(students: Student[]): Student[] {
  return students.slice().sort((a, b) => {
    const classCompare = compareClass(a.class, b.class);
    if (classCompare !== 0) return classCompare;
    return (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase(), 'id-ID');
  });
}

interface Teacher {
  id: string;
  name: string;
  niy: string;
  dutyDay?: string;
  userId?: string;
}

interface AttendanceSession {
  id: string;
  date: string;
  className: string;
  meetingNumber: number;
  records: Record<string, Status>;
  userId?: string;
  recordedByRole?: string;
  recordedBySubject?: string;
  recordedByTeacherName?: string;
  recordedByTeacherId?: string;
  piketTeacherId?: string;
  piketTeacherName?: string;
  lastEditedByTeacherName?: string;
  lastEditedAt?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  editLogs?: any[];
}

interface CustomUser {
  id: string;
  fullname: string;
  username: string;
  password?: string;
  createdAt?: string;
}

function AttendanceView({
  classList,
  students,
  teachers,
  attendanceSessions,
  showToast,
  activeDb,
  activeAuth,
  trackOp,
  profileData,
  classWaliMap = {},
  classWaliNiyMap = {},
  currentUser
}: {
  classList: string[];
  students: Student[];
  teachers: Teacher[];
  attendanceSessions: AttendanceSession[];
  showToast: (message: string, type: 'success' | 'info' | 'error') => void;
  activeDb: Firestore;
  activeAuth: Auth;
  trackOp: (type: 'read' | 'write', count?: number) => void;
  profileData: Record<string, unknown>;
  classWaliMap?: Record<string, string>;
  classWaliNiyMap?: Record<string, string>;
  currentUser: User | null;
}) {
  const [date, setDate] = useState(() => {
    return format(new Date(), 'yyyy-MM-dd');
  });
  const [selectedTeacherId, setSelectedTeacherId] = useState('');
  const [isDateModalOpen, setIsDateModalOpen] = useState(false);
  const [calendarMonth, setCalendarMonth] = useState(new Date());
  
  const [selectedClass, setSelectedClass] = useState('');
  const [isClassModalOpen, setIsClassModalOpen] = useState(false);
  
  const [currentRecords, setCurrentRecords] = useState<Record<string, Status>>({});
  const [isEditing, setIsEditing] = useState(false);
  const [existingSession, setExistingSession] = useState<AttendanceSession | null>(null);

  const studentsInClass = useMemo(() => students.filter(s => s.class === selectedClass), [students, selectedClass]);
  const classSessions = useMemo(() => attendanceSessions.filter(s => s.className === selectedClass), [attendanceSessions, selectedClass]);

  const currentClassWaliKelas = useMemo(() => {
    return getWaliKelasForClass(selectedClass, classWaliMap, students);
  }, [selectedClass, classWaliMap, students]);

  const currentClassWaliKelasNiy = useMemo(() => {
    return getWaliKelasNiyForClass(selectedClass, classWaliNiyMap, students, teachers, currentClassWaliKelas);
  }, [selectedClass, classWaliNiyMap, students, teachers, currentClassWaliKelas]);
  
  // Calculate meeting number for next or current session
  const meetingNumber = useMemo(() => {
    if (existingSession) return existingSession.meetingNumber;
    return classSessions.length + 1;
  }, [existingSession, classSessions]);

  const daysInMonth = useMemo(() => {
    return eachDayOfInterval({
      start: startOfWeek(startOfMonth(calendarMonth), { weekStartsOn: 1 }),
      end: endOfWeek(endOfMonth(calendarMonth), { weekStartsOn: 1 })
    });
  }, [calendarMonth]);

  // Warn if leaving page with unsaved changes
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (isEditing && studentsInClass.length > 0) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [isEditing, studentsInClass.length]);

  useEffect(() => {
    if (profileData?.role === 'Wali Kelas' && profileData?.waliKelasClass) {
      if (selectedClass !== profileData.waliKelasClass) {
        setSelectedClass(profileData.waliKelasClass);
      }
    }
  }, [profileData?.role, profileData?.waliKelasClass]);

  // Reset selected teacher if they are not scheduled for the selected date's dutyDay
  useEffect(() => {
    if (selectedTeacherId && teachers.length > 0) {
      const selectedTeacher = teachers.find(t => t.id === selectedTeacherId);
      if (selectedTeacher) {
        const currentDayName = format(new Date(date), 'EEEE', { locale: idLocale });
        const isDutyDay = selectedTeacher.dutyDay && 
                           selectedTeacher.dutyDay.trim() !== '' && 
                           selectedTeacher.dutyDay.trim() !== '-' && 
                           selectedTeacher.dutyDay.toLowerCase().includes(currentDayName.toLowerCase());
        if (!isDutyDay) {
          setSelectedTeacherId('');
        }
      }
    }
  }, [date, teachers, selectedTeacherId]);

  useEffect(() => {
    if (!selectedClass || studentsInClass.length === 0) {
      setCurrentRecords({});
      setExistingSession(null);
      setIsEditing(false);
      return;
    }

    const found = attendanceSessions.find(s => s.className === selectedClass && s.date === date);

    if (found) {
      setExistingSession(found);
      
      // Ensure any newly added students default to 'Hadir' if they are missing from existing records
      const mergedRecords: Record<string, Status> = { ...found.records };
      studentsInClass.forEach(s => {
        if (!mergedRecords[s.id]) {
          mergedRecords[s.id] = 'Hadir';
        }
      });
      setCurrentRecords(mergedRecords);
      
      if ((found as { recordedByTeacherId?: string }).recordedByTeacherId) {
        setSelectedTeacherId((found as { recordedByTeacherId?: string }).recordedByTeacherId || '');
      }
      setIsEditing(false); // Default view mode when loaded
    } else {
      setExistingSession(null);
      setIsEditing(true); // new entry
      // Default to 'Hadir' for all students
      const initial: Record<string, Status> = {};
      studentsInClass.forEach(s => {
        initial[s.id] = 'Hadir';
      });
      setCurrentRecords(initial);
    }
  }, [date, selectedClass, attendanceSessions, studentsInClass]);

  const handleStatusChange = (studentId: string, status: Status) => {
    setCurrentRecords(prev => ({ ...prev, [studentId]: status }));
  };

  const handleDeleteSession = async (sessionId: string) => {
    try {
      trackOp('write', 1);
      await deleteDoc(doc(activeDb, 'attendanceSessions', sessionId));
      showToast('Data presensi berhasil dihapus.', 'info');
      const currentSessionId = existingSession?.id || `${(currentUser?.uid || activeAuth.currentUser?.uid || 'admin')}_${date}_${selectedClass}`;
      if (sessionId === currentSessionId) {
        setExistingSession(null);
        setIsEditing(true);
        const initial: Record<string, Status> = {};
        studentsInClass.forEach(s => { initial[s.id] = 'Hadir'; });
        setCurrentRecords(initial);
      }
    } catch (err) {
      showToast('Gagal menghapus presensi.', 'error');
      handleFirestoreError(err, OperationType.DELETE, 'attendanceSessions');
    }
  };

  const handleShareWA = () => {
    if (!existingSession) return;
    
    const absents = studentsInClass.filter(s => currentRecords[s.id] !== 'Hadir');
    const selectedTeacherObj = teachers.find(t => t.id === selectedTeacherId);
    
    let text = 'Bismillah\n';
    text += 'Assalamualaikum wr wb\n';
    text += 'Berikut laporan kehadiran siswa untuk hari ini : \n\n';
    text += '*Laporan Presensi*\n';
    text += 'Kelas: ' + selectedClass + '\n';
    if (currentClassWaliKelas) {
      text += 'Wali Kelas: ' + currentClassWaliKelas + '\n';
    }
    if (profileData?.role === 'Petugas Piket' && selectedTeacherObj) {
      text += 'Guru yang Mengabsen: ' + selectedTeacherObj.name + '\n';
    }
    text += 'Tanggal: ' + format(new Date(date), 'EEEE, dd MMMM yyyy', { locale: idLocale }) + '\n';
    text += 'Pertemuan ke: ' + meetingNumber + '\n\n';
    
    if (absents.length === 0) {
      text += '_Semua siswa hadir._\n\n';
    } else {
      text += '*Siswa yang tidak hadir:*\n';
      absents.forEach((s, idx) => {
        text += (idx + 1) + '. ' + s.name + ' (' + currentRecords[s.id] + ')\n';
      });
    }
    text += 'Terimakasih 🙏';
    
    const encodedText = encodeURIComponent(text);
    window.open('https://wa.me/?text=' + encodedText, '_blank');
  };

  const handleSave = async () => {
    if (!selectedClass) {
      showToast('Mohon pilih kelas terlebih dahulu.', 'error');
      return;
    }
    const authUser = currentUser || activeAuth.currentUser;
    if (!authUser) {
      showToast('Sesi berakhir. Silakan masuk kembali.', 'error');
      return;
    }

    const isPiket = profileData?.role === 'Petugas Piket';
    const isWaliOrMapel = profileData?.role === 'Wali Kelas' || profileData?.role === 'Guru Mapel';
    const selectedDate = new Date(date);
    
    if (isPiket) {
       if (!selectedTeacherId) {
          showToast('Mohon pilih Guru yang sedang mengabsen terlebih dahulu.', 'error');
          return;
       }

       const piketDay = getDayIndex(profileData?.hariPiket);
       if (piketDay !== -1 && selectedDate.getDay() !== piketDay) {
          const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
          const dayName = days[piketDay] || 'yang ditentukan';
          showToast(`Akses Ditolak: Anda hanya dapat mengisi presensi pada hari Piket Anda (${dayName}).`, 'error');
          return;
       }
    } else if (!isWaliOrMapel) {
       if (!selectedTeacherId) {
          showToast('Mohon pilih Petugas Piket terlebih dahulu.', 'error');
          return;
       }
    }
    
    // Check missing
    const missing = studentsInClass.some(s => !currentRecords[s.id]);
    if (missing) {
      showToast('Mohon lengkapi semua status presensi.', 'error');
      return;
    }

    const sessionIdRole = profileData?.role === 'Guru Mapel' ? `Mapel_${profileData?.mataPelajaran || 'Umum'}` : profileData?.role || 'Umum';
    const sessionId = existingSession ? existingSession.id : `${authUser.uid}_${date}_${selectedClass}_${sessionIdRole}`;
    
    const selectedTeacherObj = teachers.find(t => t.id === selectedTeacherId);
    
    // Determine editor name
    let currentEditor = '';
    if (isPiket) {
      if (selectedTeacherObj) {
        currentEditor = selectedTeacherObj.name;
      } else {
        currentEditor = profileData?.namaGuruMapel || profileData?.fullname || 'Petugas Piket';
      }
    } else {
      currentEditor = profileData?.namaGuruMapel || profileData?.namaGuruWali || profileData?.fullname || 'Guru';
    }

    let piketNameForSession = '';
    if (isPiket) {
      piketNameForSession = profileData?.namaGuruMapel || profileData?.fullname || 'Petugas Piket';
    } else if (selectedTeacherObj) {
      piketNameForSession = selectedTeacherObj.name;
    } else if (existingSession?.piketTeacherName) {
      piketNameForSession = existingSession.piketTeacherName;
    } else {
      try {
        const dayName = format(new Date(date), 'EEEE', { locale: idLocale });
        const dutyTeacher = teachers.find(t => t.dutyDay && t.dutyDay.toLowerCase().includes(dayName.toLowerCase()));
        if (dutyTeacher) {
          piketNameForSession = dutyTeacher.name;
        }
      } catch {
        // fallback
      }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const currentLogs = (existingSession as any)?.editLogs || [];
    const updatedLogs = [...currentLogs];

    if (existingSession) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const changesList: any[] = [];
      Object.keys(currentRecords).forEach(studentId => {
        const oldVal = existingSession.records?.[studentId] || 'Belum diisi';
        const newVal = currentRecords[studentId];
        if (oldVal !== newVal) {
          const studentObj = studentsInClass.find(s => s.id === studentId);
          changesList.push({
            studentName: studentObj ? studentObj.name : 'Siswa',
            oldStatus: oldVal,
            newStatus: newVal
          });
        }
      });

      if (changesList.length > 0) {
        const newLog = {
          timestamp: new Date().toISOString(),
          editorName: currentEditor,
          editorRole: profileData?.role || 'Guru',
          changes: changesList
        };
        updatedLogs.push(newLog);
      }
    }

    const newSessionRaw: AttendanceSession & { 
      recordedByRole?: string;
      recordedBySubject?: string;
      recordedByTeacherName?: string;
      recordedByTeacherId?: string;
      piketTeacherId?: string;
      piketTeacherName?: string;
      lastEditedByTeacherName?: string;
      lastEditedAt?: string;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      editLogs?: any[];
    } = {
      id: sessionId,
      date,
      className: selectedClass,
      meetingNumber: meetingNumber,
      records: currentRecords,
      userId: authUser.uid,
      recordedByRole: profileData?.role || undefined,
      recordedBySubject: profileData?.mataPelajaran || undefined,
      recordedByTeacherName: existingSession?.recordedByTeacherName || currentEditor,
      recordedByTeacherId: isPiket ? selectedTeacherId : undefined,
      piketTeacherId: isPiket ? (profileData?.id || authUser.uid) : (selectedTeacherId || undefined),
      piketTeacherName: piketNameForSession || undefined,
      lastEditedByTeacherName: currentEditor,
      lastEditedAt: new Date().toISOString(),
      editLogs: updatedLogs
    };

    // Clean undefined properties to prevent Firestore error
    const newSession = Object.entries(newSessionRaw).reduce((acc, [key, val]) => {
      if (val !== undefined) {
        acc[key] = val;
      }
      return acc;
    }, {} as Record<string, unknown>) as AttendanceSession;

    try {
      trackOp('write', 1);
      await setDoc(doc(activeDb, 'attendanceSessions', sessionId), newSession);
      setExistingSession(newSession);
      setIsEditing(false);
      showToast(`Presensi ${selectedClass} tanggal ${format(new Date(date), 'dd/MM/yyyy')} berhasil disimpan.`, 'success');
    } catch (err) {
      showToast('Gagal menyimpan presensi.', 'error');
      handleFirestoreError(err, OperationType.WRITE, 'attendanceSessions');
    }
  };

  return (
    <div className="p-4 sm:p-6 lg:p-8 space-y-6 max-w-7xl mx-auto w-full pb-32">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
         <div>
           <h2 className="text-2xl sm:text-3xl font-bold text-slate-800 tracking-tight">Presensi Kelas</h2>
           <p className="text-sm font-medium text-slate-600 mt-1">Catat kehadiran harian siswa</p>
         </div>
      </div>

      <div className="bg-white p-6 sm:p-8 rounded-[1.5rem] border-2 border-slate-300/60 shadow-sm space-y-6">
         <div className={`grid grid-cols-1 sm:grid-cols-2 ${profileData?.role === 'Petugas Piket' ? 'lg:grid-cols-4' : 'lg:grid-cols-3'} gap-6 sm:gap-8`}>
            <div>
              <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-2.5">Tanggal Presensi</label>
              <button 
                onClick={() => setIsDateModalOpen(true)}
                className="w-full px-4 py-3.5 bg-slate-50/50 border-2 border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-600/20 focus:border-emerald-600 transition-all outline-none text-left font-semibold flex items-center justify-between group hover:bg-white hover:border-emerald-400"
              >
                <span className="text-slate-800 font-bold">
                  {format(parseISO(date), 'dd MMM yyyy')}
                </span>
                <ChevronDown className="w-5 h-5 text-slate-500 group-hover:text-emerald-600 transition-colors" />
              </button>
            </div>
            <div className="relative">
              <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-2.5">Pilih Kelas</label>
              <button 
                onClick={() => {
                  if (profileData?.role === 'Wali Kelas' && profileData?.waliKelasClass) {
                     showToast('Sebagai Wali Kelas, presensi otomatis terkunci ke kelas perwalian Anda.', 'info');
                     return;
                  }
                  setIsClassModalOpen(true);
                }}
                className={`w-full px-4 py-3.5 bg-slate-50/50 border-2 border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-600/20 focus:border-emerald-600 transition-all outline-none text-left font-semibold flex items-center justify-between group ${(profileData?.role === 'Wali Kelas' && profileData?.waliKelasClass) ? 'opacity-80 cursor-default' : 'hover:bg-white hover:border-emerald-400'}`}
              >
                 <span className={selectedClass ? "text-slate-800 font-bold" : "text-slate-500"}>
                   {selectedClass || "-- Pilih Kelas --"}
                 </span>
                 <ChevronDown className="w-5 h-5 text-slate-500 group-hover:text-emerald-600 transition-colors" />
              </button>
            </div>

            {!(profileData?.role === 'Guru Mapel' || profileData?.role === 'Wali Kelas' || profileData?.role === 'Petugas Piket') && (
              <div className="relative">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-2.5">Pilih Petugas Piket</label>
                <select
                  value={selectedTeacherId}
                  onChange={(e) => setSelectedTeacherId(e.target.value)}
                  className="w-full px-4 py-3.5 bg-slate-50/50 border-2 border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-600/20 focus:border-emerald-600 transition-all outline-none text-left font-semibold hover:bg-white hover:border-emerald-400 appearance-none cursor-pointer text-slate-800"
                >
                  <option value="">-- Pilih Petugas Piket --</option>
                  {teachers.map(t => {
                    const currentDayName = format(new Date(date), 'EEEE', { locale: idLocale });
                    const isDutyDay = t.dutyDay && 
                                       t.dutyDay.trim() !== '' && 
                                       t.dutyDay.trim() !== '-' && 
                                       t.dutyDay.toLowerCase().includes(currentDayName.toLowerCase());
                    return (
                      <option 
                        key={t.id} 
                        value={t.id} 
                        className={isDutyDay ? "text-[#098f41] font-bold bg-[#098f41]/10" : "text-slate-700"}
                      >
                        {t.name}{t.dutyDay && t.dutyDay !== '-' ? ` (Piket: ${t.dutyDay})` : ' (Belum ada jadwal piket)'} {isDutyDay ? ' - ✨ JADWAL HARI INI ✨' : ''}
                      </option>
                    );
                  })}
                </select>
                <div className="absolute inset-y-0 right-0 flex items-center px-4 pointer-events-none mt-6">
                  <ChevronDown className="w-5 h-5 text-slate-500" />
                </div>
              </div>
            )}

            {profileData?.role === 'Petugas Piket' && (
              <>
                <div className="relative">
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-2.5">Guru yang Mengabsen</label>
                  <select
                    value={selectedTeacherId}
                    onChange={(e) => setSelectedTeacherId(e.target.value)}
                    className="w-full px-4 py-3.5 bg-slate-50/50 border-2 border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-600/20 focus:border-emerald-600 transition-all outline-none text-left font-semibold hover:bg-white hover:border-emerald-400 appearance-none cursor-pointer"
                  >
                    <option value="">-- Pilih Guru --</option>
                    {teachers.map(t => {
                      const currentDayName = format(new Date(date), 'EEEE', { locale: idLocale });
                      const isDutyDay = t.dutyDay && 
                                         t.dutyDay.trim() !== '' && 
                                         t.dutyDay.trim() !== '-' && 
                                         t.dutyDay.toLowerCase().includes(currentDayName.toLowerCase());
                      return (
                        <option 
                          key={t.id} 
                          value={t.id} 
                          className={isDutyDay ? "text-[#098f41] font-bold bg-[#098f41]/10" : "text-slate-700"}
                        >
                          {t.name}{t.dutyDay && t.dutyDay !== '-' ? ` (Piket: ${t.dutyDay})` : ' (Belum ada jadwal piket)'} {isDutyDay ? ' - ✨ JADWAL HARI INI ✨' : ''}
                        </option>
                      );
                    })}
                  </select>
                  <div className="absolute inset-y-0 right-0 flex items-center px-4 pointer-events-none mt-6">
                    <ChevronDown className="w-5 h-5 text-slate-500" />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-2.5">Wali Kelas</label>
                  <div className="w-full px-4 py-2.5 bg-emerald-50/60 border-2 border-emerald-300/80 rounded-xl font-semibold flex items-center justify-between min-h-[52px]">
                    <div className="flex flex-col truncate">
                      <span className={`text-sm truncate ${currentClassWaliKelas ? 'text-emerald-900 font-bold' : 'text-slate-400 italic font-medium'}`}>
                        {selectedClass ? (currentClassWaliKelas || 'Belum diatur') : '-- Pilih Kelas --'}
                      </span>
                      {currentClassWaliKelas && currentClassWaliKelasNiy && currentClassWaliKelasNiy !== '-' && (
                        <span className="text-[10px] text-emerald-700 font-medium">
                          NIY: <span className="font-mono font-bold">{currentClassWaliKelasNiy}</span>
                        </span>
                      )}
                    </div>
                    {currentClassWaliKelas && (
                      <span className="text-[10px] bg-emerald-200/80 text-emerald-800 font-bold px-2 py-0.5 rounded-md ml-2 shrink-0">
                        Wali
                      </span>
                    )}
                  </div>
                </div>
              </>
            )}
         </div>
      </div>
      
      {/* Modern Date Selection Modal */}
      {isDateModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-slate-900/40 backdrop-blur-sm animate-in fade-in duration-200">
           <div className="bg-white rounded-[2rem] w-full max-w-sm shadow-2xl border border-white/20 overflow-hidden flex flex-col animate-in slide-in-from-bottom-8 zoom-in-95 duration-300">
             <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
               <div>
                  <h3 className="text-xl font-black text-slate-800 tracking-tight">Pilih Tanggal</h3>
                  <p className="text-xs font-semibold text-slate-600 mt-0.5">Tentukan tanggal presensi</p>
               </div>
               <button 
                 onClick={() => setIsDateModalOpen(false)}
                 className="p-2 hover:bg-slate-200/50 text-slate-600 hover:text-slate-700 rounded-xl transition-colors shrink-0"
               >
                 <X className="w-5 h-5" />
               </button>
             </div>
             
             <div className="p-6">
                <div className="flex items-center justify-between mb-6">
                   <button 
                     onClick={() => setCalendarMonth(prev => subMonths(prev, 1))}
                     className="p-2 rounded-xl text-slate-600 hover:bg-slate-100 transition-colors"
                   >
                     <ChevronLeft className="w-5 h-5" />
                   </button>
                   <span className="font-bold text-slate-800 tracking-tight">
                     {format(calendarMonth, 'MMMM yyyy')}
                   </span>
                   <button 
                     onClick={() => setCalendarMonth(prev => addMonths(prev, 1))}
                     className="p-2 rounded-xl text-slate-600 hover:bg-slate-100 transition-colors"
                   >
                     <ChevronRight className="w-5 h-5" />
                   </button>
                </div>
                
                <div className="grid grid-cols-7 gap-1 text-center mb-2">
                   {['Sn', 'Sl', 'Rb', 'Km', 'Jm', 'Sb', 'Mg'].map(day => (
                     <div key={day} className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">{day}</div>
                   ))}
                </div>
                
                <div className="grid grid-cols-7 gap-1">
                   {daysInMonth.map(day => {
                     const dayStr = format(day, 'yyyy-MM-dd');
                     const isSelected = dayStr === date;
                     const isCurrentMonth = isSameMonth(day, calendarMonth);
                     const isToday = isSameDay(day, new Date());
                     const hasAttendance = selectedClass && attendanceSessions.some(s => s.className === selectedClass && s.date === dayStr);
                     
                     return (
                       <button
                         key={day.toISOString()}
                         onClick={() => {
                           setDate(dayStr);
                           setIsDateModalOpen(false);
                         }}
                         disabled={!isCurrentMonth}
                         className={`aspect-square flex flex-col items-center justify-center rounded-xl text-sm font-bold transition-all relative ${
                           !isCurrentMonth ? 'text-transparent cursor-default' :
                           isSelected ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/20 scale-105' :
                           hasAttendance ? 'text-emerald-600 bg-emerald-50/50 hover:bg-emerald-100/50 outline outline-1 outline-lime-200/50' : 'text-slate-700 hover:bg-slate-50'
                         }`}
                       >
                         {isCurrentMonth ? format(day, 'd') : ''}
                         <div className="absolute bottom-1.5 flex gap-1 items-center justify-center">
                           {isToday && (
                             <div className={`w-1 h-1 rounded-full ${isSelected ? 'bg-white/70' : 'bg-slate-300'}`}></div>
                           )}
                           {hasAttendance && (
                             <div className={`w-1.5 h-1.5 rounded-full shadow-sm ${isSelected ? 'bg-white' : 'bg-sky-500'}`}></div>
                           )}
                         </div>
                       </button>
                     );
                   })}
                </div>
             </div>
           </div>
        </div>
      )}

      {/* Modern Class Selection Modal */}
      {isClassModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-slate-900/40 backdrop-blur-sm animate-in fade-in duration-200">
           <div className="bg-white rounded-[2rem] w-full max-w-md shadow-2xl border border-white/20 overflow-hidden flex flex-col animate-in slide-in-from-bottom-8 zoom-in-95 duration-300">
             <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
               <div>
                  <h3 className="text-xl font-black text-slate-800 tracking-tight">Pilih Kelas</h3>
                  <p className="text-xs font-semibold text-slate-600 mt-0.5">Tentukan kelas untuk presensi</p>
               </div>
               <button 
                 onClick={() => setIsClassModalOpen(false)}
                 className="p-2 hover:bg-slate-200/50 text-slate-600 hover:text-slate-700 rounded-xl transition-colors shrink-0"
               >
                 <X className="w-5 h-5" />
               </button>
             </div>
             
             <div className="p-6 max-h-[60vh] overflow-y-auto">
                <div className="grid grid-cols-2 gap-3 sm:gap-4">
                  {classList.slice().sort((a,b) => compareClass(a, b)).map(c => {
                    const studentCount = students.filter(s => s.class === c).length;
                    return (
                      <button
                        key={c}
                        onClick={() => {
                          setSelectedClass(c);
                          setIsClassModalOpen(false);
                        }}
                        className={`relative p-3.5 sm:p-4 rounded-2xl border-2 transition-all duration-200 flex flex-col items-center justify-center gap-1.5 group ${
                          selectedClass === c 
                            ? 'border-emerald-600 bg-emerald-50 text-lime-700 shadow-sm' 
                            : 'border-slate-100 bg-white hover:border-emerald-400 hover:bg-emerald-50/30 text-slate-600'
                        }`}
                      >
                        <div className={`w-9 h-9 rounded-xl flex items-center justify-center transition-colors ${
                          selectedClass === c ? 'bg-emerald-600 text-white' : 'bg-slate-100 text-slate-500 group-hover:bg-emerald-100 group-hover:text-emerald-700'
                        }`}>
                           <Building2 className="w-4 h-4" />
                        </div>
                        <span className="font-bold text-sm text-slate-800">{c}</span>
                        <span className="text-[11px] text-slate-500 truncate max-w-full px-1">
                          {studentCount} Siswa
                        </span>
                        {selectedClass === c && (
                          <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-emerald-600"></div>
                        )}
                      </button>
                    );
                  })}
                  {classList.length === 0 && (
                    <div className="col-span-2 text-center py-8">
                       <p className="text-slate-600 font-medium text-sm">Belum ada kelas yang terdaftar.</p>
                    </div>
                  )}
                </div>
             </div>
           </div>
        </div>
      )}

      {selectedClass && studentsInClass.length === 0 && (
         <div className="bg-amber-50 border border-amber-200 text-amber-700 p-6 rounded-[2rem] text-center font-medium">
           Belum ada data siswa di kelas {selectedClass}. <br/> Silakan tambahkan siswa terlebih dahulu di menu <b>Manajemen Siswa</b>.
         </div>
      )}

      {selectedClass && studentsInClass.length > 0 && (
        <div className="bg-white rounded-[1.5rem] border-2 border-slate-300/60 shadow-sm overflow-hidden flex flex-col">
          <div className="p-6 border-b border-slate-100 bg-white flex flex-col sm:flex-row sm:items-center justify-between gap-4">
             <div>
               <h3 className="text-lg font-bold text-slate-800 tracking-tight">Tabel Presensi</h3>
               {existingSession && !isEditing ? (
                  <p className="text-sm font-bold text-emerald-600 mt-1 flex items-center gap-1.5">
                     <Check className="w-4 h-4" /> Data sudah tersimpan
                  </p>
               ) : existingSession && isEditing ? (
                  <p className="text-sm font-bold text-amber-600 mt-1 flex items-center gap-1.5">
                     <Pencil className="w-4 h-4" /> Mode Edit Data Presensi
                  </p>
               ) : null}
             </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse min-w-[700px]">
              <thead>
                <tr className="bg-slate-50/50 border-b border-slate-100">
                  <th className="p-4 w-12 font-semibold text-slate-600 text-xs uppercase tracking-widest text-center">No</th>
                  <th className="p-4 font-semibold text-slate-600 text-xs uppercase tracking-widest">Nama Lengkap</th>
                  <th className="p-4 font-semibold text-slate-600 text-xs uppercase tracking-widest text-center">Keterangan</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {studentsInClass.map((student, index) => (
                  <tr key={student.id} className="hover:bg-slate-50/30 transition-colors group">
                    <td className="p-4 text-center font-medium text-slate-500 text-sm">{index + 1}</td>
                    <td className="p-4 font-semibold text-slate-700 group-hover:text-slate-900 transition-colors">{student.name}</td>
                    <td className="p-4">
                       <div className="flex items-center justify-end gap-2 sm:gap-2.5">
                         {(['Hadir', 'Sakit', 'Izin', 'Alpa', 'Dispen'] as Status[]).map(statusOpt => {
                            const isSelected = currentRecords[student.id] === statusOpt;
                            
                            let colorClass = 'border-slate-200 text-slate-600 hover:border-slate-300 bg-white';
                            if (isSelected) {
                               if (statusOpt === 'Hadir') colorClass = 'bg-emerald-500 text-white border-emerald-500 shadow-md shadow-emerald-500/20';
                               else if (statusOpt === 'Sakit') colorClass = 'bg-sky-500 text-white border-sky-500 shadow-md shadow-sky-500/20';
                               else if (statusOpt === 'Izin') colorClass = 'bg-amber-500 text-white border-amber-500 shadow-md shadow-amber-500/20';
                               else if (statusOpt === 'Alpa') colorClass = 'bg-rose-500 text-white border-rose-500 shadow-md shadow-rose-500/20';
                               else if (statusOpt === 'Dispen') colorClass = 'bg-indigo-500 text-white border-indigo-500 shadow-md shadow-indigo-500/20';
                            } else if (!isEditing) {
                               colorClass = 'border-slate-100 text-slate-300 bg-slate-50/50';
                            }

                            return (
                              <button
                                key={statusOpt}
                                disabled={!isEditing}
                                onClick={() => handleStatusChange(student.id, statusOpt)}
                                className={`px-4 py-2 sm:py-2.5 rounded-lg text-xs sm:text-sm font-semibold border ${colorClass} transition-all ${!isEditing && isSelected ? 'opacity-90 grayscale-[30%]' : 'disabled:opacity-50 disabled:cursor-not-allowed'}`}
                              >
                                {statusOpt}
                              </button>
                            );
                         })}
                       </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {existingSession && (
            <div className="px-6 py-3.5 bg-emerald-50/90 border-t border-emerald-100 space-y-2">
              <div className="flex items-center gap-2 text-xs sm:text-sm font-medium text-emerald-900">
                <CheckCircle className="w-4 h-4 text-emerald-600 shrink-0" />
                <span>Telah di lakukan presensi oleh <strong className="font-bold text-emerald-950">{existingSession.lastEditedByTeacherName || existingSession.recordedByTeacherName || existingSession.piketTeacherName || 'Petugas Piket'}</strong></span>
              </div>
              {existingSession.editLogs && existingSession.editLogs.length > 0 && (
                <div className="mt-2.5 p-3 bg-white/60 border border-emerald-200/60 rounded-xl space-y-2 text-xs max-h-[160px] overflow-y-auto">
                  <span className="font-bold text-slate-700 block mb-1 flex items-center gap-1">📋 Riwayat Perubahan (Audit Trail):</span>
                  {[...existingSession.editLogs].reverse().map((log: any, idx: number) => (
                    <div key={idx} className="border-b border-slate-200/60 last:border-0 pb-2 last:pb-0 space-y-1">
                      <div className="flex items-center justify-between text-[10px] text-slate-500">
                        <span className="font-semibold text-slate-600">{log.editorName} ({log.editorRole})</span>
                        <span>{format(new Date(log.timestamp), 'dd MMM yyyy, HH:mm', { locale: idLocale })}</span>
                      </div>
                      <div className="pl-2 border-l-2 border-amber-400 space-y-0.5">
                        {log.changes.map((ch: any, cIdx: number) => (
                          <div key={cIdx} className="text-[11px] text-slate-700">
                            <strong>{ch.studentName}</strong>: <span className="text-slate-500 line-through">{ch.oldStatus}</span> → <span className="font-semibold text-emerald-700">{ch.newStatus}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          <div className="p-6 border-t border-slate-100 bg-white flex flex-col gap-3">
             <div className="flex items-center justify-end gap-2 w-full sm:w-auto">
               {isEditing ? (
                 <>
                    {existingSession && (
                       <button
                          onClick={() => {
                             setIsEditing(false);
                             setCurrentRecords(existingSession.records);
                          }}
                          className="flex-1 sm:flex-none sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all active:scale-95"
                       >
                          Batal
                       </button>
                    )}
                    <button 
                       onClick={handleSave} 
                       className="flex-1 sm:flex-none sm:w-auto bg-emerald-600 text-white px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-emerald-700 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Save className="w-4 h-4" /> Simpan Presensi
                    </button>
                 </>
               ) : (
                 <>
                    <button
                       onClick={() => setIsEditing(true)}
                       className="flex-1 sm:flex-none sm:w-auto bg-white border-2 border-slate-300 text-slate-700 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-slate-50 transition-all flex items-center justify-center gap-2"
                    >
                       <Pencil className="w-4 h-4" /> Edit
                    </button>
                    <button
                       onClick={() => { if(existingSession) handleDeleteSession(existingSession.id); }}
                       className="flex-1 sm:flex-none sm:w-auto bg-white text-rose-600 border border-rose-200 px-6 py-2.5 rounded-lg font-semibold text-sm hover:bg-rose-50 transition-all active:scale-95 flex items-center justify-center gap-2"
                    >
                       <Trash2 className="w-4 h-4" /> Hapus
                    </button>
                 </>
               )}
             </div>
             
             {!isEditing && existingSession && (
                <div className="flex justify-end pt-2 border-t border-slate-50 mt-1">
                    <button
                       onClick={handleShareWA}
                       className="w-full sm:w-auto bg-[#25D366] text-white px-8 py-3 rounded-xl font-bold text-sm hover:bg-[#128C7E] transition-all active:scale-95 flex items-center justify-center gap-2 shadow-md shadow-[#25D366]/20"
                    >
                       <Send className="w-5 h-5" /> Kirim ke Group WA
                    </button>
                </div>
             )}
          </div>
        </div>
      )}

      {selectedClass && classSessions.length > 0 && (
        <div className="bg-white rounded-[1.5rem] border-2 border-slate-300/60 shadow-sm overflow-hidden flex flex-col mt-8">
          <div className="p-6 sm:p-8 border-b border-slate-100 bg-white flex flex-col items-center justify-center text-center gap-2">
            <h3 className="text-xl font-bold text-slate-800 tracking-tight">Riwayat Presensi</h3>
            <div className="text-center">
                <p className="text-[10px] font-bold text-slate-500 tracking-widest uppercase mb-1">Informasi Pertemuan</p>
                <p className="text-sm font-semibold text-slate-700">Total {classSessions.length} Pertemuan Efektif</p>
            </div>
          </div>
          <div className="grid grid-cols-4 gap-2 sm:gap-4 p-4 sm:p-6 bg-white">
            {[...classSessions].sort((a,b) => new Date(a.date).getTime() - new Date(b.date).getTime()).map((session, idx) => (
              <button 
                key={session.id} 
                onClick={() => {
                  setDate(session.date);
                  window.scrollTo({ top: 0, behavior: 'smooth' });
                }}
                className="py-1.5 flex items-center justify-center gap-2 hover:text-emerald-600 transition-colors cursor-pointer focus:outline-none group"
              >
                <span className="font-medium text-slate-500 text-xs group-hover:text-[#077a37] transition-colors">{idx + 1}.</span>
                <span className="font-medium text-slate-600 text-xs sm:text-sm group-hover:text-emerald-600 transition-colors">{format(new Date(session.date), 'dd MMM yyyy')}</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

const toAuthEmail = (input: string) => {
  const trimmed = input.toLowerCase().trim();
  if (trimmed.includes('@')) {
    return trimmed;
  }
  const clean = trimmed.replace(/[^a-z0-9._-]/g, '');
  return `${clean || 'user'}@fresh.com`;
};

const safeFirebaseAuth = async (authInstance: Auth, inputEmail: string, pass: string, username: string) => {
  const primaryEmail = toAuthEmail(inputEmail);
  const cleanUser = inputEmail.toLowerCase().trim().replace(/[^a-z0-9._-]/g, '');
  const legacyEmail = inputEmail.includes('@') ? inputEmail : `${cleanUser || 'user'}@smadaf.com`;

  try {
    const res = await signInWithEmailAndPassword(authInstance, primaryEmail, pass);
    return res.user;
  } catch {
    try {
      const res = await signInWithEmailAndPassword(authInstance, legacyEmail, pass);
      return res.user;
    } catch {
      try {
        const createRes = await createUserWithEmailAndPassword(authInstance, primaryEmail, pass);
        return createRes.user;
      } catch {
        try {
          const createRes2 = await createUserWithEmailAndPassword(authInstance, legacyEmail, pass);
          return createRes2.user;
        } catch {
          try {
            const anonRes = await signInAnonymously(authInstance);
            return anonRes.user;
          } catch {
            return authInstance.currentUser || { uid: username || 'user', email: primaryEmail } as unknown as import('firebase/auth').User;
          }
        }
      }
    }
  }
};

const safeSetLocalStorage = (key: string, value: string) => {
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    console.warn(`LocalStorage quota exceeded or disabled for key "${key}":`, e);
  }
};

let lastSpokenText = '';
let lastSpokenTime = 0;

const speakText = (text: string) => {
  if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
    const now = Date.now();
    if (lastSpokenText === text && now - lastSpokenTime < 5000) {
      return;
    }
    lastSpokenText = text;
    lastSpokenTime = now;

    window.speechSynthesis.cancel();
    
    const startSpeaking = () => {
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = 'id-ID';
      utterance.rate = 0.95;
      utterance.pitch = 1.0;
      
      const voices = window.speechSynthesis.getVoices();
      
      // Filter all Indonesian voices using multiple language code variations
      const indonesianVoices = voices.filter(v => {
        const langLower = v.lang.toLowerCase();
        return langLower.startsWith('id') || langLower.includes('id') || langLower === 'id_id';
      });

      // Sort to prioritize natural premium voices (like Google Bahasa Indonesia or professional female voices)
      indonesianVoices.sort((a, b) => {
        const score = (v: SpeechSynthesisVoice) => {
          const nameLower = v.name.toLowerCase();
          let pts = 0;
          if (nameLower.includes('google')) pts += 10;
          if (nameLower.includes('premium') || nameLower.includes('natural')) pts += 5;
          if (nameLower.includes('female') || nameLower.includes('wanita') || nameLower.includes('gadis')) pts += 3;
          return pts;
        };
        return score(b) - score(a);
      });

      const selectedVoice = indonesianVoices[0];
      if (selectedVoice) {
        utterance.voice = selectedVoice;
        console.log(`Menggunakan suara Bahasa Indonesia: ${selectedVoice.name} (${selectedVoice.lang})`);
      } else {
        console.warn("Suara Bahasa Indonesia asli tidak ditemukan. Menggunakan suara sistem default.");
      }
      
      window.speechSynthesis.speak(utterance);
    };

    if (window.speechSynthesis.getVoices().length === 0) {
      const handleVoicesChanged = () => {
        startSpeaking();
        window.speechSynthesis.removeEventListener('voiceschanged', handleVoicesChanged);
      };
      window.speechSynthesis.addEventListener('voiceschanged', handleVoicesChanged);
    } else {
      startSpeaking();
    }
  } else {
    console.warn("Speech Synthesis tidak didukung oleh browser ini.");
  }
};

// Helper to calculate 14:00 WIB daily reset cycle and time info
export interface WibCycleInfo {
  cycleKey: string;
  wibTime: Date;
  currentDateFormatted: string;
  currentTimeStr: string;
  currentDayName: string;
  cycleStartDate: Date;
  nextResetDate: Date;
  countdownFormatted: string;
  hoursLeft: number;
  minutesLeft: number;
  secondsLeft: number;
}

export const getWibCycleInfo = (now: Date = new Date()): WibCycleInfo => {
  // Convert current time to WIB (UTC+7)
  const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
  const wibTime = new Date(utc + (7 * 3600000));
  
  const wibHours = wibTime.getHours();
  const wibMinutes = wibTime.getMinutes();
  const wibSeconds = wibTime.getSeconds();

  // If before 14:00 WIB, the current cycle started yesterday at 14:00 WIB
  const cycleStartDate = new Date(wibTime);
  if (wibHours < 14) {
    cycleStartDate.setDate(cycleStartDate.getDate() - 1);
  }
  cycleStartDate.setHours(14, 0, 0, 0);

  // Next reset is at 14:00 WIB on the next day after cycleStartDate
  const nextResetDate = new Date(cycleStartDate);
  nextResetDate.setDate(nextResetDate.getDate() + 1);
  nextResetDate.setHours(14, 0, 0, 0);

  // Cycle key uniquely identifies the 24-hour cycle starting at 14:00 WIB
  const y = cycleStartDate.getFullYear();
  const m = String(cycleStartDate.getMonth() + 1).padStart(2, '0');
  const d = String(cycleStartDate.getDate()).padStart(2, '0');
  const cycleKey = `${y}-${m}-${d}_14:00_WIB`;

  // Indonesian day and month names
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  
  const currentDayName = days[wibTime.getDay()];
  const currentDayNum = wibTime.getDate();
  const currentMonthName = months[wibTime.getMonth()];
  const currentYear = wibTime.getFullYear();
  const currentTimeStr = `${String(wibHours).padStart(2, '0')}:${String(wibMinutes).padStart(2, '0')}:${String(wibSeconds).padStart(2, '0')} WIB`;
  const currentDateFormatted = `${currentDayName}, ${currentDayNum} ${currentMonthName} ${currentYear}`;

  // Time remaining until next reset (in ms)
  const msUntilReset = Math.max(0, nextResetDate.getTime() - wibTime.getTime());
  const hoursLeft = Math.floor(msUntilReset / 3600000);
  const minutesLeft = Math.floor((msUntilReset % 3600000) / 60000);
  const secondsLeft = Math.floor((msUntilReset % 60000) / 1000);
  const countdownFormatted = `${hoursLeft} jam ${minutesLeft} menit ${secondsLeft} detik`;

  return {
    cycleKey,
    wibTime,
    currentDateFormatted,
    currentTimeStr,
    currentDayName,
    cycleStartDate,
    nextResetDate,
    countdownFormatted,
    hoursLeft,
    minutesLeft,
    secondsLeft
  };
};

export default function App() {
  const [activeUserCustomData, setActiveUserCustomData] = useState<{
    fullname: string;
    username: string;
    configText: string;
    password?: string;
    role?: string;
  } | null>(() => {
    const saved = localStorage.getItem('kaguci_active_custom_user');
    return saved ? JSON.parse(saved) : null;
  });

  const { activeAuth, activeDb } = useMemo(() => {
    return { activeAuth: auth, activeDb: dbDefault };
  }, []);

  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    return localStorage.getItem('kaguci_has_logged_in') === 'true';
  });
  const [currentUser, setCurrentUser] = useState<User | null>(() => {
    try {
      const savedUser = localStorage.getItem('kaguci_active_custom_user');
      if (savedUser) {
        const parsed = JSON.parse(savedUser);
        if (parsed && parsed.username) {
          const usernameKey = parsed.username.toLowerCase().trim();
          return {
            uid: usernameKey,
            email: toAuthEmail(usernameKey),
            displayName: parsed.fullname || parsed.username,
            photoURL: localStorage.getItem(`kaguci_avatar_${usernameKey}`) || localStorage.getItem('kaguci_avatar_current') || null
          } as unknown as User;
        }
      }
    } catch {
      // ignore
    }
    return null;
  });
  const [isAuthLoading, setIsAuthLoading] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'students' | 'attendance' | 'reports' | 'homeroom_report' | 'profile'>('dashboard');

  // Live WIB Cycle Information & 14:00 WIB Clock
  const [cycleInfo, setCycleInfo] = useState<WibCycleInfo>(() => getWibCycleInfo());

  // Firestore Usage Tracking & Quota Modal - Resets daily at 14:00 WIB
  const [sessionUsage, setSessionUsage] = useState(() => {
    const info = getWibCycleInfo();
    const saved = localStorage.getItem('kaguci_session_usage');
    const savedCycle = localStorage.getItem('kaguci_session_cycle');
    
    // If we have saved usage for the exact current 14:00 WIB cycle, load it
    if (saved && savedCycle === info.cycleKey) {
      try {
        return JSON.parse(saved);
      } catch {
        // fallback
      }
    }
    
    // Reset to 0 for a new 14:00 WIB cycle
    localStorage.setItem('kaguci_session_cycle', info.cycleKey);
    localStorage.setItem('kaguci_session_date', info.currentDateFormatted);
    return { reads: 0, writes: 0 };
  });
  const [showQuotaModal, setShowQuotaModal] = useState(false);
  const [showLogoutConfirm, setShowLogoutConfirm] = useState(false);
  const [showIdleTip, setShowIdleTip] = useState(false);
  const [showHomeroomRoleAlert, setShowHomeroomRoleAlert] = useState(false);

  // State for PWA installation & iframe detection
  interface PWAInstallPromptEvent extends Event {
    prompt: () => Promise<void>;
    userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }>;
  }

  const [deferredPrompt, setDeferredPrompt] = useState<PWAInstallPromptEvent | null>(null);
  const [isInstallable, setIsInstallable] = useState(false);
  const [isAppInstalled, setIsAppInstalled] = useState(false);
  const isInsideIframe = useMemo(() => {
    try {
      return window.self !== window.top;
    } catch {
      return true;
    }
  }, []);

  useEffect(() => {
    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as PWAInstallPromptEvent);
      setIsInstallable(true);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);

    const handleAppInstalled = () => {
      setIsAppInstalled(true);
      setIsInstallable(false);
      setDeferredPrompt(null);
    };

    window.addEventListener('appinstalled', handleAppInstalled);

    const iosStandalone = (navigator as unknown as { standalone?: boolean }).standalone;
    if (window.matchMedia('(display-mode: standalone)').matches || iosStandalone) {
      setIsAppInstalled(true);
    }

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
      window.removeEventListener('appinstalled', handleAppInstalled);
    };
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) {
      alert("Instruksi instalasi: Silakan klik tombol 'Buka di Tab Baru' di kanan atas preview, lalu pilih opsi 'Instal' atau 'Tambahkan ke Layar Utama' dari menu browser Anda.");
      return;
    }
    await deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    console.log(`User response to the install prompt: ${outcome}`);
    if (outcome === 'accepted') {
      setIsInstallable(false);
      setDeferredPrompt(null);
    }
  };

  useEffect(() => {
    if (!isLoggedIn) return;
    
    let idleTimeout: NodeJS.Timeout;
    let tipTimeout: NodeJS.Timeout;
    
    const resetIdleTimer = () => {
      clearTimeout(idleTimeout);
      clearTimeout(tipTimeout);
      setShowIdleTip(false);
      
      // Muncul setelah 30 detik tidak ada aktivitas (biar gampang ditest) -> ganti jadi 60 detik (1 menit)
      idleTimeout = setTimeout(() => {
        setShowIdleTip(true);
        // Hilang lagi setelah 10 detik
        tipTimeout = setTimeout(() => {
          setShowIdleTip(false);
        }, 10000);
      }, 60000);
    };

    resetIdleTimer();

    window.addEventListener('mousemove', resetIdleTimer);
    window.addEventListener('keydown', resetIdleTimer);
    window.addEventListener('touchstart', resetIdleTimer);
    window.addEventListener('scroll', resetIdleTimer);

    return () => {
      clearTimeout(idleTimeout);
      clearTimeout(tipTimeout);
      window.removeEventListener('mousemove', resetIdleTimer);
      window.removeEventListener('keydown', resetIdleTimer);
      window.removeEventListener('touchstart', resetIdleTimer);
      window.removeEventListener('scroll', resetIdleTimer);
    };
  }, [isLoggedIn]);

  // Save session usage and active cycle
  useEffect(() => {
    const current = getWibCycleInfo();
    localStorage.setItem('kaguci_session_usage', JSON.stringify(sessionUsage));
    localStorage.setItem('kaguci_session_cycle', current.cycleKey);
    localStorage.setItem('kaguci_session_date', current.currentDateFormatted);
  }, [sessionUsage]);

  // Live WIB clock & Automatic 14:00 WIB daily reset monitor
  useEffect(() => {
    const timer = setInterval(() => {
      const newInfo = getWibCycleInfo();
      setCycleInfo(newInfo);

      const savedCycle = localStorage.getItem('kaguci_session_cycle');
      // If 14:00 WIB has just passed, trigger automatic reset to 0
      if (savedCycle && savedCycle !== newInfo.cycleKey) {
        console.log("⏰ 14:00 WIB reached! Resetting usage limits (reads & writes) to 0.");
        localStorage.setItem('kaguci_session_cycle', newInfo.cycleKey);
        localStorage.setItem('kaguci_session_date', newInfo.currentDateFormatted);
        setSessionUsage({ reads: 0, writes: 0 });

        // Update Firestore for current custom account immediately
        if (activeUserCustomData?.username) {
          const usernameKey = activeUserCustomData.username.toLowerCase().trim();
          setDoc(doc(dbDefault, 'custom_accounts', usernameKey), {
            dailyUsageReads: 0,
            dailyUsageWrites: 0,
            dailyUsageCycle: newInfo.cycleKey,
            dailyUsageDate: newInfo.currentDateFormatted,
            dailyUsageTime: newInfo.currentTimeStr,
            dailyUsageLastReset: newInfo.currentTimeStr
          }, { merge: true }).catch(err => console.warn('Reset sync to custom_accounts error:', err));
        }
      }
    }, 1000);
    return () => clearInterval(timer);
  }, [activeUserCustomData?.username]);

  const lastSyncedUsage = useRef({ reads: 0, writes: 0 });
  const autoLoginAttemptsRef = useRef(0);
  useEffect(() => {
    if (!activeUserCustomData?.username) return;
    const interval = setInterval(() => {
      if (sessionUsage.reads !== lastSyncedUsage.current.reads || sessionUsage.writes !== lastSyncedUsage.current.writes) {
        lastSyncedUsage.current = { ...sessionUsage };
        const usernameKey = activeUserCustomData.username.toLowerCase().trim();
        const current = getWibCycleInfo();
        setDoc(doc(dbDefault, 'custom_accounts', usernameKey), {
          dailyUsageReads: sessionUsage.reads || 0,
          dailyUsageWrites: sessionUsage.writes || 0,
          dailyUsageCycle: current.cycleKey,
          dailyUsageDate: current.currentDateFormatted,
          dailyUsageTime: current.currentTimeStr
        }, { merge: true }).catch(err => console.warn('Failed syncing usage', err));
      }
    }, 15000); // 15s flush
    return () => clearInterval(interval);
  }, [sessionUsage, activeUserCustomData?.username]);

  useEffect(() => {
    const handleQuota = () => setShowQuotaModal(true);
    window.addEventListener('firestore-quota-exceeded', handleQuota);
    return () => window.removeEventListener('firestore-quota-exceeded', handleQuota);
  }, []);

  const trackOp = useCallback((type: 'read' | 'write', count: number = 1) => {
    const current = getWibCycleInfo();
    const savedCycle = localStorage.getItem('kaguci_session_cycle');

    setSessionUsage((prev: { reads: number; writes: number }) => {
      // If cycle has shifted past 14:00 WIB, start fresh from 0 for the new cycle
      if (savedCycle && savedCycle !== current.cycleKey) {
        localStorage.setItem('kaguci_session_cycle', current.cycleKey);
        localStorage.setItem('kaguci_session_date', current.currentDateFormatted);
        return {
          reads: type === 'read' ? count : 0,
          writes: type === 'write' ? count : 0
        };
      }
      return {
        ...prev,
        [type === 'read' ? 'reads' : 'writes']: (prev[type === 'read' ? 'reads' : 'writes'] || 0) + count
      };
    });
  }, []);


  const [avatarUrl, setAvatarUrl] = useState<string | null>(() => {
    try {
      const savedUser = localStorage.getItem('kaguci_active_custom_user');
      if (savedUser) {
        const parsed = JSON.parse(savedUser);
        const username = parsed.username || '';
        const savedAvatar = localStorage.getItem(`kaguci_avatar_${username.toLowerCase()}`);
        return savedAvatar || null;
      }
    } catch {
      // Ignore
    }
    return null;
  });

  const [showSplash, setShowSplash] = useState(true);
  const [minSplashTimeElapsed, setMinSplashTimeElapsed] = useState(false);
  const [isFirstSessionLoad] = useState(() => !sessionStorage.getItem('kaguci_has_loaded'));

  useEffect(() => {
    if (isFirstSessionLoad) {
      sessionStorage.setItem('kaguci_has_loaded', 'true');
    }
    const timer = setTimeout(() => {
      setMinSplashTimeElapsed(true);
    }, isFirstSessionLoad ? 2500 : 800); // 2.5s for initial splash, 0.8s for fast reload
    return () => clearTimeout(timer);
  }, [isFirstSessionLoad]);

  useEffect(() => {
    if (minSplashTimeElapsed && !isAuthLoading) {
      setShowSplash(false);
    }
  }, [minSplashTimeElapsed, isAuthLoading]);

  // Safety timeout to ensure that if Firebase Auth takes too long to initialize,
  // we still dismiss the splash loading screen so the application does not hang indefinitely.
  useEffect(() => {
    const safetyTimer = setTimeout(() => {
      if (isAuthLoading) {
        console.warn("Safety timeout triggered: Firebase Auth initialization took too long. Forcing load completion.");
        setIsAuthLoading(false);
        setMinSplashTimeElapsed(true);
      }
    }, 4500); // 4.5 seconds maximum loading limit
    return () => clearTimeout(safetyTimer);
  }, [isAuthLoading]);

  const [profileData, setProfileData] = useState(() => {
    try {
      const savedUser = localStorage.getItem('kaguci_active_custom_user');
      if (savedUser) {
        const parsed = JSON.parse(savedUser);
        const username = parsed.username || '';
        const savedProfile = localStorage.getItem(`kaguci_profile_${username.toLowerCase()}`);
        if (savedProfile) {
          return JSON.parse(savedProfile);
        }
      }
    } catch {
      // Ignore
    }
    return {
      namaGuruMapel: '',
      namaKepalaSekolah: '',
      nipGuruMapel: '',
      nipKepalaSekolah: '',
      namaKurikulum: '',
      nipKurikulum: '',
      jabatanKurikulum: 'Wakasek Kurikulum',
      namaKesiswaan: '',
      nipKesiswaan: '',
      jabatanKesiswaan: 'Wakasek Kesiswaan',
      namaGuruWali: '',
      nipGuruWali: '',
      jabatanGuruWali: 'Guru Wali',
      namaBK: '',
      nipBK: '',
      jabatanBK: 'Guru BK',
      namaHumas: '',
      nipHumas: '',
      jabatanHumas: 'Wakasek Humas',
      semester: 'Ganjil',
      tahunPelajaran: '',
      mataPelajaran: '',
      role: 'Guru Mapel',
      hariPiket: '',
      waliKelasClass: '',
      jumlahSiswaLakiLaki: '',
      jumlahSiswaPerempuan: ''
    };
  });

  const handleOpenHomeroomReport = useCallback(() => {
    if (profileData.role === 'Wali Kelas') {
      setActiveTab('homeroom_report');
    } else {
      setShowHomeroomRoleAlert(true);
    }
  }, [profileData.role]);

  const hasWelcomedRef = useRef(false);

  useEffect(() => {
    if (isLoggedIn && !isAuthLoading) {
      const cleanUser = (activeUserCustomData?.username || 'user').toLowerCase().trim();
      const sessionKey = `kaguci_welcomed_${cleanUser}`;
      const sessionWelcomed = sessionStorage.getItem(sessionKey);
      
      if (!sessionWelcomed && !hasWelcomedRef.current) {
        const name = (profileData?.namaGuruMapel || activeUserCustomData?.fullname || activeUserCustomData?.username || "").trim();
        if (!name) {
          const checkTimer = setTimeout(() => {
            const delayedName = (profileData?.namaGuruMapel || activeUserCustomData?.fullname || activeUserCustomData?.username || "").trim();
            hasWelcomedRef.current = true;
            sessionStorage.setItem(sessionKey, 'true');
            const phrase = delayedName 
              ? `Selamat datang ${delayedName} di aplikasi SMART DF App.`
              : `Selamat datang di aplikasi SMART DF App.`;
            speakText(phrase);
          }, 1200);
          return () => clearTimeout(checkTimer);
        } else {
          hasWelcomedRef.current = true;
          sessionStorage.setItem(sessionKey, 'true');
          const phrase = `Selamat datang ${name} di aplikasi SMART DF App.`;
          const speakTimer = setTimeout(() => {
            speakText(phrase);
          }, 600);
          return () => clearTimeout(speakTimer);
        }
      }
    } else if (!isLoggedIn) {
      hasWelcomedRef.current = false;
    }
  }, [isLoggedIn, isAuthLoading, activeUserCustomData, profileData]);

  useEffect(() => {
    if (isLoggedIn && activeUserCustomData?.username && profileData) {
      const hasContent = Object.values(profileData).some(val => val !== '' && val !== 'Ganjil');
      if (hasContent) {
        safeSetLocalStorage(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`, JSON.stringify(profileData));
      }
    }
  }, [profileData, isLoggedIn, activeUserCustomData]);

  useEffect(() => {
    if (isLoggedIn && activeUserCustomData?.username && avatarUrl) {
      safeSetLocalStorage(`kaguci_avatar_${activeUserCustomData.username.toLowerCase()}`, avatarUrl);
    }
  }, [avatarUrl, isLoggedIn, activeUserCustomData]);

  const [isProfileEditing, setIsProfileEditing] = useState(false);
  const [isProfileSaving, setIsProfileSaving] = useState(false);

  // States & helper for Admin (Kumpulan User) - Custom Requested
  const isAdmin = useMemo(() => {
    const email = (activeAuth.currentUser?.email || '').toLowerCase().trim();
    const username = (activeUserCustomData?.username || '').toLowerCase().trim();
    const fullname = (activeUserCustomData?.fullname || profileData.namaGuruMapel || '').toLowerCase().trim();

    const isMatch = 
      email.includes('agan') || 
      username.includes('agan') || 
      fullname.includes('agan') ||
      username.includes('parta') ||
      fullname.includes('parta') ||
      fullname.includes('s.kom') ||
      username.includes('endang') ||
      fullname.includes('endang') ||
      fullname.includes('sukmaya') ||
      username === 'admin' ||
      email === 'agan.parta@gmail.com';

    return isMatch;
  }, [activeAuth.currentUser, activeUserCustomData, profileData.namaGuruMapel]);

  const [allUsers, setAllUsers] = useState<CustomUser[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState<boolean>(false);
  const [userToDelete, setUserToDelete] = useState<CustomUser | null>(null);
  const [isDeletingUser, setIsDeletingUser] = useState<boolean>(false);
  
  const [userToEdit, setUserToEdit] = useState<CustomUser | null>(null);
  const [editFullname, setEditFullname] = useState<string>('');
  const [editPassword, setEditPassword] = useState<string>('');
  const [editUsername, setEditUsername] = useState<string>('');
  const [isSavingUser, setIsSavingUser] = useState<boolean>(false);

  const handleEditUserClick = (user: CustomUser) => {
    setUserToEdit(user);
    setEditFullname(user.fullname || '');
    setEditPassword(user.password || '');
    setEditUsername(user.username || '');
  };

  const handleSaveEditedUser = async () => {
    if (!userToEdit) return;
    if (!editFullname.trim()) {
      showToast('Nama Lengkap tidak boleh kosong.', 'error');
      return;
    }
    if (!editUsername.trim()) {
      showToast('Nama Pengguna (Username) tidak boleh kosong.', 'error');
      return;
    }

    setIsSavingUser(true);
    const oldUsername = userToEdit.username.toLowerCase().trim();
    const newUsername = editUsername.toLowerCase().trim();
    const path = `custom_accounts/${newUsername}`;
    try {
      trackOp('write', 1);

      let oldDocData: Record<string, any> = {};
      const oldDocSnap = await getDoc(doc(dbDefault, 'custom_accounts', oldUsername));
      if (oldDocSnap.exists()) {
        oldDocData = oldDocSnap.data();
      }

      if (newUsername !== oldUsername) {
        const checkDoc = await getDoc(doc(dbDefault, 'custom_accounts', newUsername));
        if (checkDoc.exists()) {
          showToast(`Username "${newUsername}" sudah digunakan oleh akun lain.`, 'error');
          setIsSavingUser(false);
          return;
        }
      }

      const updatedData = {
        ...oldDocData,
        fullname: editFullname.trim(),
        username: editUsername.trim(),
        password: editPassword.trim() || oldDocData.password || userToEdit.password || '••••••••',
        updatedAt: new Date().toISOString()
      };

      await setDoc(doc(dbDefault, 'custom_accounts', newUsername), updatedData, { merge: true });

      if (newUsername !== oldUsername) {
        await deleteDoc(doc(dbDefault, 'custom_accounts', oldUsername));
      }

      // If active user was edited, update local state & storage
      const activeUn = (activeUserCustomData?.username || '').toLowerCase().trim();
      if (activeUn === oldUsername) {
        const updatedCustomUser = {
          ...activeUserCustomData,
          username: editUsername.trim(),
          fullname: editFullname.trim(),
          password: editPassword.trim() || activeUserCustomData?.password
        };
        setActiveUserCustomData(updatedCustomUser);
        localStorage.setItem('kaguci_active_custom_user', JSON.stringify(updatedCustomUser));
      }

      showToast(`User "${editFullname.trim()}" berhasil diperbarui.`, 'success');
      setUserToEdit(null);
      fetchAllUsers();
    } catch (err) {
      console.error('Error saving edited user:', err);
      handleFirestoreError(err, OperationType.WRITE, path);
      showToast('Gagal memperbarui user: ' + (err instanceof Error ? err.message : 'Server error'), 'error');
    } finally {
      setIsSavingUser(false);
    }
  };

  const fetchAllUsers = useCallback(async () => {
    if (!isAdmin) return;
    setIsLoadingUsers(true);
    try {
      trackOp('read', 1);
      const qSnapshot = await getDocs(collection(dbDefault, 'custom_accounts'));
      const list: CustomUser[] = [];
      qSnapshot.forEach((docSnap) => {
        const u = docSnap.data();
        list.push({
          id: docSnap.id,
          fullname: u.fullname || docSnap.id,
          username: u.username || docSnap.id,
          password: u.password || '••••••••',
          createdAt: u.createdAt || ''
        });
      });

      // Ensure active user is included in the list
      const activeName = activeUserCustomData?.fullname || profileData.namaGuruMapel || 'Agan Parta,S.Kom.,Gr';
      const activeUn = activeUserCustomData?.username || 'agan.parta';
      const exists = list.some(u => u.username.toLowerCase().trim() === activeUn.toLowerCase().trim());
      if (!exists) {
        list.unshift({
          id: activeUn.toLowerCase().trim(),
          fullname: activeName,
          username: activeUn,
          password: '••••••••',
          createdAt: new Date().toISOString()
        });
      }

      list.sort((a, b) => (a.fullname || '').localeCompare(b.fullname || ''));
      setAllUsers(list);
    } catch (err) {
      console.error('Error fetching users:', err);
      const activeName = activeUserCustomData?.fullname || profileData.namaGuruMapel || 'Agan Parta,S.Kom.,Gr';
      const activeUn = activeUserCustomData?.username || 'agan.parta';
      setAllUsers([{
        id: activeUn.toLowerCase().trim(),
        fullname: activeName,
        username: activeUn,
        password: '••••••••',
        createdAt: new Date().toISOString()
      }]);
    } finally {
      setIsLoadingUsers(false);
    }
  }, [isAdmin, activeUserCustomData, profileData.namaGuruMapel]);

  useEffect(() => {
    if (isLoggedIn && isAdmin) {
      fetchAllUsers();
    }
  }, [isLoggedIn, isAdmin, fetchAllUsers]);

  const handleDeleteUserClick = (user: CustomUser) => {
    setUserToDelete(user);
  };

  const handleConfirmDeleteUser = async () => {
    if (!userToDelete) return;

    if (userToDelete.username.toLowerCase().trim() === 'petugaspiket') {
      showToast('Akun Petugas Piket tidak bisa dihapus.', 'error');
      return;
    }

    setIsDeletingUser(true);
    const targetUsername = userToDelete.username.toLowerCase().trim();
    const targetFullname = userToDelete.fullname;
    const path = `custom_accounts/${targetUsername}`;
    try {
      trackOp('write', 1);
      await deleteDoc(doc(dbDefault, 'custom_accounts', targetUsername));
      showToast(`User "${targetFullname}" berhasil dihapus secara permanen.`, 'success');
      setUserToDelete(null);
      fetchAllUsers();
    } catch (err) {
      console.error('Error deleting user:', err);
      handleFirestoreError(err, OperationType.DELETE, path);
      showToast('Gagal menghapus user: ' + (err instanceof Error ? err.message : 'Server error'), 'error');
    } finally {
      setIsDeletingUser(false);
    }
  };



  const [classList, setClassList] = useState<string[]>([]);
  const [classListLoaded, setClassListLoaded] = useState(false);
  const [classWaliMap, setClassWaliMap] = useState<Record<string, string>>({});
  const [classWaliNiyMap, setClassWaliNiyMap] = useState<Record<string, string>>({});

  // Custom states for reset database features
  const [resetModalType, setResetModalType] = useState<'none' | 'new_semester' | 'new_year' | 'everything' | 'clear_all_students' | 'clear_all_teachers'>('none');
  const [resetSuccessModal, setResetSuccessModal] = useState<'none' | 'new_semester' | 'new_year' | 'everything' | 'clear_all_students' | 'clear_all_teachers'>('none');
  const [, setResetConfirmInput] = useState('');
  const [resetPasswordInput, setResetPasswordInput] = useState('');
  const [resetPasswordError, setResetPasswordError] = useState('');
  const [showResetPassword, setShowResetPassword] = useState(false);
  const [isVerifyingPassword, setIsVerifyingPassword] = useState(false);
  const [newSemesterChoice, setNewSemesterChoice] = useState<'Ganjil' | 'Genap'>('Genap');
  const [newTahunPelajaran, setNewTahunPelajaran] = useState('');
  const [isResettingData, setIsResettingData] = useState(false);
  const [resetProgress, setResetProgress] = useState(0);
  const [studentSuccessModal, setStudentSuccessModal] = useState<'none' | 'added' | 'edited' | 'deleted'>('none');

  // States for Student Absences Widget in Dashboard
  const [dashboardActiveStatsTab, setDashboardActiveStatsTab] = useState<'class_summary' | 'top_rankings'>('class_summary');
  const [dashboardSelectedClassDetail, setDashboardSelectedClassDetail] = useState<string | null>(null);

  // State for Online/Offline connectivity monitor
  const [isOnline, setIsOnline] = useState<boolean>(navigator.onLine);
  const [firebaseConnected, setFirebaseConnected] = useState<boolean>(true);
  const [syncStatus, setSyncStatus] = useState<'active' | 'syncing' | 'error'>('active');
  const [lastSyncError, setLastSyncError] = useState<string | null>(null);
  const [cloudLastSync, setCloudLastSync] = useState<string | null>(null);
  const [quotaExceeded, setQuotaExceeded] = useState(false);

  // Connectivity Listeners
  useEffect(() => {
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, []);

  // Monitor Firebase Connection State
  useEffect(() => {
    if (!activeDb) return;
    
    // Using a simple interval check or real-time connection status if metadata allows
    // For Firestore, we can listen to metadata changes on a document or check sync state
    // But a simple reliable way is listening to system events and error states in snapshots
    
    const checkConnection = () => {
      // Periodic check or rely on onSnapshot errors which we will add below
    };
    
    const interval = setInterval(checkConnection, 10000);
    return () => clearInterval(interval);
  }, [activeDb]);

  // Clean any legacy hash/query/saved credentials from URL & storage on load
  useEffect(() => {
    try {
      localStorage.removeItem('kaguci_saved_credentials');
    } catch {
      /* ignore */
    }
    if (window.location.hash || window.location.search) {
      try {
        window.history.replaceState(null, '', window.location.pathname);
      } catch {
        window.location.hash = '';
      }
    }
  }, []);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(activeAuth, (user) => {
      if (user) {
        setCurrentUser(user);
        setIsLoggedIn(true);
        safeSetLocalStorage('kaguci_has_logged_in', 'true');
        
        // Load user photo directly from firebase user auth object
        if (user.photoURL) {
          setAvatarUrl(user.photoURL);
        }
        
        // Fetch central account mapping for metadata lookup (configText, fullname, profileData)
        if (activeUserCustomData?.username) {
          getDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim())).then(centralDoc => {
            if (centralDoc.exists()) {
              const centralData = centralDoc.data();
              if (centralData.photoURL && !avatarUrl) {
                setAvatarUrl(centralData.photoURL);
              }
              if (centralData.profileData) {
                setProfileData((prev: typeof profileData) => ({ ...prev, ...centralData.profileData }));
              } else if (centralData.namaGuruMapel || centralData.namaKepalaSekolah) {
                const flatProfile = {
                  namaGuruMapel: centralData.namaGuruMapel || '',
                  namaKepalaSekolah: centralData.namaKepalaSekolah || '',
                  nipGuruMapel: centralData.nipGuruMapel || '',
                  nipKepalaSekolah: centralData.nipKepalaSekolah || '',
                  namaKurikulum: centralData.namaKurikulum || '',
                  nipKurikulum: centralData.nipKurikulum || '',
                  jabatanKurikulum: centralData.jabatanKurikulum || 'Wakasek Kurikulum',
                  namaKesiswaan: centralData.namaKesiswaan || '',
                  nipKesiswaan: centralData.nipKesiswaan || '',
                  jabatanKesiswaan: centralData.jabatanKesiswaan || 'Wakasek Kesiswaan',
                  namaGuruWali: centralData.namaGuruWali || '',
                  nipGuruWali: centralData.nipGuruWali || '',
                  jabatanGuruWali: centralData.jabatanGuruWali || 'Guru Wali',
                  namaBK: centralData.namaBK || '',
                  nipBK: centralData.nipBK || '',
                  jabatanBK: centralData.jabatanBK || 'Guru BK',
                  namaHumas: centralData.namaHumas || '',
                  nipHumas: centralData.nipHumas || '',
                  jabatanHumas: centralData.jabatanHumas || 'Wakasek Humas',
                  semester: centralData.semester || 'Ganjil',
                  tahunPelajaran: centralData.tahunPelajaran || '',
                  mataPelajaran: centralData.mataPelajaran || '',
                  role: centralData.role || 'Guru Mapel',
                  waliKelasClass: centralData.waliKelasClass || '',
                  jumlahSiswaLakiLaki: centralData.jumlahSiswaLakiLaki || '',
                  jumlahSiswaPerempuan: centralData.jumlahSiswaPerempuan || ''
                };
                setProfileData((prev: typeof profileData) => ({ ...prev, ...flatProfile }));
              }
            }
          }).catch(err => console.warn('Gagal membaca sinkronisasi metadata dari portal pusat:', err));
        }

        // Async fetch avatar and classList from Firestore without blocking auth state resolution
        getDoc(doc(activeDb, 'users', user.uid)).then(userDoc => {
          if (userDoc.exists()) {
            const data = userDoc.data();
            if (data.photoURL) {
               setAvatarUrl(data.photoURL);
            }
            // Support both flat format and nested 'profileData' format for utmost reliability!
            if (data.profileData) {
               setProfileData((prev: typeof profileData) => ({ ...prev, ...data.profileData }));
            } else if (data.namaGuruMapel || data.namaKepalaSekolah) {
               const flatProfile = {
                 namaGuruMapel: data.namaGuruMapel || '',
                 namaKepalaSekolah: data.namaKepalaSekolah || '',
                 nipGuruMapel: data.nipGuruMapel || '',
                 nipKepalaSekolah: data.nipKepalaSekolah || '',
                 namaKurikulum: data.namaKurikulum || '',
                 nipKurikulum: data.nipKurikulum || '',
                 jabatanKurikulum: data.jabatanKurikulum || 'Wakasek Kurikulum',
                 namaKesiswaan: data.namaKesiswaan || '',
                 nipKesiswaan: data.nipKesiswaan || '',
                 jabatanKesiswaan: data.jabatanKesiswaan || 'Wakasek Kesiswaan',
                 namaGuruWali: data.namaGuruWali || '',
                 nipGuruWali: data.nipGuruWali || '',
                 jabatanGuruWali: data.jabatanGuruWali || 'Guru Wali',
                 namaBK: data.namaBK || '',
                 nipBK: data.nipBK || '',
                 jabatanBK: data.jabatanBK || 'Guru BK',
                 namaHumas: data.namaHumas || '',
                 nipHumas: data.nipHumas || '',
                 jabatanHumas: data.jabatanHumas || 'Wakasek Humas',
                 semester: data.semester || 'Ganjil',
                 tahunPelajaran: data.tahunPelajaran || '',
                 mataPelajaran: data.mataPelajaran || '',
                 role: data.role || 'Guru Mapel',
                 waliKelasClass: data.waliKelasClass || '',
                 jumlahSiswaLakiLaki: data.jumlahSiswaLakiLaki || '',
                 jumlahSiswaPerempuan: data.jumlahSiswaPerempuan || ''
               };
               setProfileData((prev: typeof profileData) => ({ ...prev, ...flatProfile }));
            }

            if (data.classList && Array.isArray(data.classList)) {
              setClassList(data.classList);
            }
            if (data.classWaliMap && typeof data.classWaliMap === 'object') {
              setClassWaliMap(data.classWaliMap);
            }
            if (data.classWaliNiyMap && typeof data.classWaliNiyMap === 'object') {
              setClassWaliNiyMap(data.classWaliNiyMap);
            }
          }
          setClassListLoaded(true);
        }).catch(error => {
          console.warn('Failed to fetch user document or profile on auth state change:', error);
          setClassListLoaded(true);
        });
        setIsAuthLoading(false);
      } else {
        // If there's no user in Firebase Auth, check if we have an active custom user session
        const hasLoggedIn = localStorage.getItem('kaguci_has_logged_in') === 'true';
        const savedCustom = localStorage.getItem('kaguci_active_custom_user');
        const savedCredsStr = localStorage.getItem('kaguci_saved_credentials');

        if (hasLoggedIn && savedCustom) {
          try {
            const parsed = JSON.parse(savedCustom);
            if (parsed && parsed.username) {
              const usernameKey = parsed.username.toLowerCase().trim();
              const fbUser = {
                uid: usernameKey,
                email: toAuthEmail(usernameKey),
                displayName: parsed.fullname || parsed.username,
                photoURL: localStorage.getItem(`kaguci_avatar_${usernameKey}`) || localStorage.getItem('kaguci_avatar_current') || null
              } as unknown as User;

              setCurrentUser(fbUser);
              setIsLoggedIn(true);

              // Background fetch profile and classes
              getDoc(doc(dbDefault, 'custom_accounts', usernameKey)).then(centralDoc => {
                if (centralDoc.exists()) {
                  const cData = centralDoc.data();
                  if (cData.photoURL) setAvatarUrl(cData.photoURL);
                  if (cData.profileData) {
                    setProfileData((prev: typeof profileData) => ({ ...prev, ...cData.profileData }));
                  }
                }
              }).catch(() => {});

              // Attempt silent background re-auth if credentials exist without breaking current state
              if (savedCredsStr && autoLoginAttemptsRef.current < 2) {
                autoLoginAttemptsRef.current += 1;
                try {
                  const creds = JSON.parse(savedCredsStr);
                  if (creds && creds.email && creds.password) {
                    safeFirebaseAuth(activeAuth, creds.email, creds.password, usernameKey)
                      .then((authUser) => {
                        if (authUser) setCurrentUser(authUser);
                      })
                      .catch(() => {});
                  }
                } catch {
                  // ignore
                }
              }

              setIsAuthLoading(false);
              return;
            }
          } catch (e) {
            console.warn("Error restoring session:", e);
          }
        }

        // If genuinely not logged in:
        setCurrentUser(null);
        setIsLoggedIn(false);
        localStorage.removeItem('kaguci_has_logged_in');
        setAvatarUrl(null);
        setClassList([]);
        setClassListLoaded(false);
        setClassWaliMap({});
        setClassWaliNiyMap({});
        setStudents([]);
        setStudentsLoaded(false);
        setAttendanceSessions([]);
        setIsAuthLoading(false);
      }
    });
    return () => unsubscribe();
  }, [activeAuth, activeDb, activeUserCustomData, avatarUrl]);

  const lastSavedClassListRef = useRef<string | null>(null);

  useEffect(() => {
    if (isLoggedIn && currentUser && classListLoaded) {
      const classListStr = JSON.stringify(classList);
      if (lastSavedClassListRef.current === classListStr) return;
      
      // We removed the auto-sync here to prevent infinite update loops with onSnapshot.
      // Class updates are now handled explicitly in the UI where they are modified.
      lastSavedClassListRef.current = classListStr;
      console.log("classList updated locally, auto-sync to Firebase disabled to prevent loops.");
    }
  }, [classList, classListLoaded, isLoggedIn, currentUser]);

  const [accountDeletedAlert, setAccountDeletedAlert] = useState(false);

  interface UserUsage {
    username: string;
    fullname: string;
    reads: number;
    writes: number;
    date: string;
    time?: string;
    cycleKey?: string;
    isCurrentCycle: boolean;
  }

  const [allUsersUsage, setAllUsersUsage] = useState<UserUsage[]>([]);
  const [isLoadingUsage, setIsLoadingUsage] = useState(false);

  useEffect(() => {
    if (activeTab === 'profile' && isLoggedIn) {
      setIsLoadingUsage(true);
      const currentCycle = getWibCycleInfo().cycleKey;
      getDocs(collection(dbDefault, 'custom_accounts'))
        .then(snap => {
          trackOp('read', snap.size || 1);
          const usages: UserUsage[] = snap.docs.map(doc => {
            const data = doc.data();
            const userCycle = data.dailyUsageCycle;
            // User limits reset automatically every day at 14:00 WIB.
            // If the user's stored cycle is older than current 14:00 WIB cycle, reads and writes are 0.
            const isCurrentCycle = Boolean(userCycle && userCycle === currentCycle);
            return {
              username: doc.id,
              fullname: data.fullname || doc.id,
              reads: isCurrentCycle ? (data.dailyUsageReads || 0) : 0,
              writes: isCurrentCycle ? (data.dailyUsageWrites || 0) : 0,
              date: isCurrentCycle ? (data.dailyUsageDate || '-') : 'Telah Reset (14:00 WIB)',
              time: isCurrentCycle ? (data.dailyUsageTime || '') : '',
              cycleKey: userCycle,
              isCurrentCycle
            };
          });
          // Sort by highest reads first
          usages.sort((a, b) => b.reads - a.reads);
          setAllUsersUsage(usages);
        })
        .catch(err => console.warn("Failed caching usage", err))
        .finally(() => setIsLoadingUsage(false));
    }
  }, [activeTab, isLoggedIn, cycleInfo.cycleKey]);

  // Monitor account deletion for custom generated accounts only (with safety guards)
  const loginTimestampRef = useRef<number>(Date.now());
  useEffect(() => {
    if (isLoggedIn) {
      loginTimestampRef.current = Date.now();
    }
  }, [isLoggedIn]);

  useEffect(() => {
    if (!isLoggedIn || isAuthLoading) return;
    
    let usernameKey = activeUserCustomData?.username;
    const isCustomSystemEmail = currentUser?.email && currentUser.email.endsWith('@kaguci.admin.system.local');
    
    if (!usernameKey && isCustomSystemEmail) {
       usernameKey = currentUser.email!.split('@')[0];
    }
    
    // If not a custom system account, or default system users, do not attach custom_accounts snapshot
    if (!usernameKey || (!activeUserCustomData && !isCustomSystemEmail)) return;
    
    usernameKey = usernameKey.toLowerCase().trim();
    if (usernameKey === 'admin' || usernameKey === 'petugaspiket') return;

    const centralRef = doc(dbDefault, 'custom_accounts', usernameKey);
    
    // Listen for changes to the user's account document
    const unsubscribe = onSnapshot(centralRef, (docSnap) => {
      if (!docSnap.metadata.fromCache) {
         trackOp('read', 1);
      }
      
      // Only proceed if confirmed by server with no pending local writes, and after at least 5s of session
      const timeSinceLogin = Date.now() - loginTimestampRef.current;
      if (!docSnap.metadata.fromCache && !docSnap.metadata.hasPendingWrites && !docSnap.exists() && timeSinceLogin > 5000) {
         console.warn("User account no longer exists in database! Forcing logout.");
         
         // 1. Cleared all stored credentials
         localStorage.removeItem(`kaguci_profile_${usernameKey}`);
         localStorage.removeItem(`kaguci_avatar_${usernameKey}`);
         sessionStorage.removeItem(`kaguci_welcomed_${usernameKey}`);
         localStorage.removeItem('kaguci_active_custom_user');
         localStorage.removeItem('kaguci_saved_credentials');
         localStorage.removeItem('kaguci_has_logged_in');
         
         try { window.history.replaceState(null, '', window.location.pathname + window.location.search); } 
         catch { window.location.hash = ''; }
         
         setActiveUserCustomData(null);
         setCurrentUser(null);
         setAuthEmail('');
         setAuthPassword('');
         setIsLoggedIn(false);
         
         try { signOut(activeAuth); } catch { /* ignore */ }
         
         setAccountDeletedAlert(true);
      }
    }, (error) => {
      console.warn("Account monitor snapshot warning:", error);
    });

    return () => unsubscribe();
  }, [isLoggedIn, isAuthLoading, activeUserCustomData, currentUser, activeAuth]);

  const [students, setStudents] = useState<Student[]>([]);
  const [studentsLoaded, setStudentsLoaded] = useState(false);
  const [teachers, setTeachers] = useState<Teacher[]>([]);
  const [teachersLoaded, setTeachersLoaded] = useState(false);
  const [attendanceSessions, setAttendanceSessions] = useState<AttendanceSession[]>([]);

  const isResettingDataRef = useRef(false);
  useEffect(() => {
    isResettingDataRef.current = isResettingData;
  }, [isResettingData]);

  useEffect(() => {
    // Wait until auth is fully initialized before deciding if we are logged out
    if (isAuthLoading) return;
    
    // Only clear if NOT logged in after auth has finished loading
    if (!currentUser) {
      console.log("Not logged in, skipping student fetch");
      setStudents([]);
      // Do not set loaded to true here, wait for actual user to be set
      return;
    }

    if (!activeDb) return;
    
    const uid = currentUser.uid;
    const activeProjId = activeDb?.app?.options?.projectId;
    const defaultProjId = dbDefault?.app?.options?.projectId;
    const isDefaultDb = (activeDb === dbDefault) || 
      (!!activeProjId && !!defaultProjId && activeProjId === defaultProjId) ||
      (!activeUserCustomData || !activeUserCustomData.configText);
    console.log("Fetching students. isDefaultDb:", isDefaultDb, "userId:", uid, "activeDb Instance:", activeDb);

    const qStudents = collection(activeDb, 'students');
    const qTeachers = collection(activeDb, 'teachers');
    
    const unsubscribeTeachers = onSnapshot(
      qTeachers,
      (snapshot) => {
        const fetchedTeachers = snapshot.docs.map(doc => {
          const data = doc.data() as Record<string, unknown>;
          const rawDutyDay = data.dutyDay || data.hariPiket || data.jadwalPiket || data.hari || data.piket || data.piketDay || data.hari_piket || '';
          return {
            id: doc.id,
            ...data,
            dutyDay: normalizeDutyDay(rawDutyDay)
          } as Teacher;
        }).filter(t => !t.userId || t.userId === uid);
        fetchedTeachers.sort((a, b) => (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase(), 'id-ID'));
        setTeachers(fetchedTeachers);
        setTeachersLoaded(true);
      },
      (error) => {
        console.error("Error fetching teachers:", error);
        setTeachersLoaded(true);
      }
    );

    const unsubscribeStudents = onSnapshot(
      qStudents, 
      (snapshot) => {
        setFirebaseConnected(true);
        setSyncStatus('active');
        setCloudLastSync(new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }));
        const readCost = snapshot.metadata.fromCache ? 0 : snapshot.docChanges().length;
        if (readCost > 0) trackOp('read', readCost);
        console.log("Firestore snapshots for students received:", snapshot.size, "Cost:", readCost);
        const fetchedStudents = snapshot.docs
          .map(doc => ({ id: doc.id, ...doc.data() } as Student))
          .filter(s => !s.userId || s.userId === uid);
        fetchedStudents.sort((a, b) => (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase(), 'id-ID'));
        setStudents(fetchedStudents);
        setStudentsLoaded(true);

        // Auto-merge classList from fetched students cautiously
        const extractedClasses = Array.from(new Set(fetchedStudents.map(s => String(s.class || '').trim()).filter(Boolean)));
        setClassList(prev => {
           // If we are currently resetting, don't perform auto-merging that could restore deleted classes
           if (isResettingDataRef.current) return [];
           
           if (fetchedStudents.length === 0 && studentsLoaded) {
             return prev; 
           }

           const merged = Array.from(new Set([...prev, ...extractedClasses]));
           merged.sort((a,b) => a.localeCompare(b, 'id-ID', { numeric: true }));
           if (JSON.stringify(prev) === JSON.stringify(merged)) return prev;
           
           // Persist newly discovered classes to Cloud if we are not in the middle of a reset
           if (uid && !isResettingDataRef.current) {
             setDoc(doc(activeDb, 'users', uid), { classList: merged, updatedAt: new Date().toISOString() }, { merge: true })
               .catch(err => console.warn("Failed to auto-sync merged classList to cloud:", err));
           }
           
           return merged;
         });
      },
      (error) => {
        console.error("Error fetching students:", error);
        setStudentsLoaded(true);
        setFirebaseConnected(false);
        setSyncStatus('error');
        setLastSyncError(`Gagal menarik data siswa: ${error.message}`);
        
        if (currentUser) {
          try {
            handleFirestoreError(error, OperationType.LIST, 'students');
          } catch(err) {
             const msg = err instanceof Error ? err.message : String(err);
             if (msg.toLowerCase().includes('permission')) {
               showToast('Firestore Rules Error! Data siswa gagal ditarik dari Cloud. Pastikan rules Firebase database mandiri Anda sudah "allow read, write: if request.auth != null;".', 'error');
             }
          }
        }
      }
    );

    const unsubscribeUser = onSnapshot(doc(activeDb, 'users', uid), (docSnapshot) => {
      setFirebaseConnected(true);
      setSyncStatus('active');
      setCloudLastSync(new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }));
      const readCost = docSnapshot.metadata.fromCache ? 0 : 1;
      if (readCost > 0) trackOp('read', readCost);
      console.log("Real-time update received for user:", uid, "Cost:", readCost);
      if (docSnapshot.exists()) {
        const data = docSnapshot.data();
        if (data.photoURL) {
          setAvatarUrl(data.photoURL);
        }
        
        // Sync Profile Data real-time
        if (data.profileData) {
           setProfileData((prev: typeof profileData) => ({ ...prev, ...data.profileData }));
        } else if (data.namaGuruMapel || data.namaKepalaSekolah) {
           const flatProfile = {
             namaGuruMapel: data.namaGuruMapel || '',
             namaKepalaSekolah: data.namaKepalaSekolah || '',
             nipGuruMapel: data.nipGuruMapel || '',
             nipKepalaSekolah: data.nipKepalaSekolah || '',
             namaKurikulum: data.namaKurikulum || '',
             nipKurikulum: data.nipKurikulum || '',
             jabatanKurikulum: data.jabatanKurikulum || 'Wakasek Kurikulum',
             namaKesiswaan: data.namaKesiswaan || '',
             nipKesiswaan: data.nipKesiswaan || '',
             jabatanKesiswaan: data.jabatanKesiswaan || 'Wakasek Kesiswaan',
             namaGuruWali: data.namaGuruWali || '',
             nipGuruWali: data.nipGuruWali || '',
             jabatanGuruWali: data.jabatanGuruWali || 'Guru Wali',
             namaBK: data.namaBK || '',
             nipBK: data.nipBK || '',
             jabatanBK: data.jabatanBK || 'Guru BK',
             namaHumas: data.namaHumas || '',
             nipHumas: data.nipHumas || '',
             jabatanHumas: data.jabatanHumas || 'Wakasek Humas',
             semester: data.semester || 'Ganjil',
             tahunPelajaran: data.tahunPelajaran || '',
             mataPelajaran: data.mataPelajaran || '',
             role: data.role || 'Guru Mapel',
             hariPiket: data.hariPiket || '',
             waliKelasClass: data.waliKelasClass || '',
             jumlahSiswaLakiLaki: data.jumlahSiswaLakiLaki || '',
             jumlahSiswaPerempuan: data.jumlahSiswaPerempuan || ''
           };
           setProfileData((prev: typeof profileData) => ({ ...prev, ...flatProfile }));
        }

        if (data.classList && Array.isArray(data.classList)) {
          const newListStr = JSON.stringify(data.classList);
          setClassList(prev => {
            if (JSON.stringify(prev) === newListStr) return prev;
            return data.classList;
          });
        }

        if (data.classWaliMap && typeof data.classWaliMap === 'object') {
          const newWaliStr = JSON.stringify(data.classWaliMap);
          setClassWaliMap(prev => {
            if (JSON.stringify(prev) === newWaliStr) return prev;
            return data.classWaliMap;
          });
        }

        if (data.classWaliNiyMap && typeof data.classWaliNiyMap === 'object') {
          const newWaliNiyStr = JSON.stringify(data.classWaliNiyMap);
          setClassWaliNiyMap(prev => {
            if (JSON.stringify(prev) === newWaliNiyStr) return prev;
            return data.classWaliNiyMap;
          });
        }
      }
    }, (error) => {
      console.error("Error fetching user doc:", error);
      setFirebaseConnected(false);
      setSyncStatus('error');
      setLastSyncError(`Gagal sinkronisasi profil: ${error.message}`);
    });

    const qSessions = collection(activeDb, 'attendanceSessions');
    const unsubscribeSessions = onSnapshot(
      qSessions, 
      (snapshot) => {
        setFirebaseConnected(true);
        setSyncStatus('active');
        setCloudLastSync(new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }));
        const readCost = snapshot.metadata.fromCache ? 0 : snapshot.docChanges().length;
        if (readCost > 0) trackOp('read', readCost);
        console.log("Real-time update received for sessions:", snapshot.size, "docs. Cost:", readCost);
        const fetchedSessions = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as AttendanceSession));
        setAttendanceSessions(fetchedSessions);
      },
      (error) => {
        console.error("Error fetching sessions:", error);
        setFirebaseConnected(false);
        setSyncStatus('error');
        setLastSyncError(`Gagal menarik data absensi: ${error.message}`);
        
        if (currentUser) {
          try {
            handleFirestoreError(error, OperationType.LIST, 'attendanceSessions');
          } catch(err) {
             const msg = err instanceof Error ? err.message : String(err);
             if (msg.toLowerCase().includes('permission')) {
               showToast('Firestore Rules Error! Data sesi gagal ditarik dari Cloud. Pastikan rules Firebase database mandiri Anda sudah "allow read, write: if request.auth != null;".', 'error');
             }
          }
        }
      }
    );

    return () => {
      unsubscribeTeachers();
      unsubscribeStudents();
      unsubscribeUser();
      unsubscribeSessions();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isLoggedIn, currentUser, activeDb, isAuthLoading, activeUserCustomData, activeAuth]);

  const [newStudent, setNewStudent] = useState({ name: '', nisn: '', class: '', waliKelas: '', waliKelasNiy: '' });
  const [editingStudentId, setEditingStudentId] = useState<string | null>(null);
  const [confirmationAction, setConfirmationAction] = useState<{ type: 'edit' | 'delete', student: Student } | null>(null);

  const [classToEdit, setClassToEdit] = useState<string | null>(null);
  const [newClassNameInput, setNewClassNameInput] = useState<string>('');
  const [classToDelete, setClassToDelete] = useState<{ name: string; studentCount: number } | null>(null);

  const [newTeacher, setNewTeacher] = useState({ name: '', niy: '', dutyDay: '' });
  const [editingTeacherId, setEditingTeacherId] = useState<string | null>(null);
  const [deletingTeacherId, setDeletingTeacherId] = useState<string | null>(null);

  const [actionPasswordModal, setActionPasswordModal] = useState<{
    title: string;
    description: string;
    onSuccess: () => void;
    expectedPassword?: string;
  } | null>(null);
  const [actionPasswordInput, setActionPasswordInput] = useState('');
  const [actionPasswordError, setActionPasswordError] = useState('');
  const [isVerifyingActionPassword, setIsVerifyingActionPassword] = useState(false);
  const [showActionPassword, setShowActionPassword] = useState(false);

  const requestEditStudent = (student: Student) => {
    setActionPasswordModal({
      title: 'Autentikasi Admin: Edit Siswa',
      description: `Masukkan password admin default untuk mengedit data siswa "${student.name}".`,
      onSuccess: () => {
        handleEditStudent(student);
      }
    });
    setActionPasswordInput('');
    setActionPasswordError('');
  };

  const requestDeleteStudent = (student: Student) => {
    setActionPasswordModal({
      title: 'Autentikasi Admin: Hapus Siswa',
      description: `Masukkan password admin default untuk menghapus data siswa "${student.name}".`,
      onSuccess: () => {
        setConfirmationAction({ type: 'delete', student });
      }
    });
    setActionPasswordInput('');
    setActionPasswordError('');
  };

  const requestEditTeacher = (teacher: Teacher) => {
    setActionPasswordModal({
      title: 'Autentikasi Admin: Edit Guru',
      description: `Masukkan password admin default untuk mengedit data guru "${teacher.name}".`,
      onSuccess: () => {
        handleEditTeacher(teacher);
      }
    });
    setActionPasswordInput('');
    setActionPasswordError('');
  };

  const requestDeleteTeacher = (teacher: Teacher) => {
    setActionPasswordModal({
      title: 'Autentikasi Admin: Hapus Guru',
      description: `Masukkan password admin default untuk menghapus data guru "${teacher.name}".`,
      onSuccess: () => {
        setTeacherConfirmationAction({ type: 'delete', teacher });
      }
    });
    setActionPasswordInput('');
    setActionPasswordError('');
  };
  
  const [isUploadingPhoto, setIsUploadingPhoto] = useState(false);
  const profilePhotoInputRef = useRef<HTMLInputElement>(null);

  const compressAndCropAvatar = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      if (!file.type.startsWith('image/')) {
        reject(new Error('Format file bukan gambar yang valid (gunakan format JPG, PNG, atau WEBP).'));
        return;
      }
      const reader = new FileReader();
      reader.onload = (event) => {
        const rawBase64 = event.target?.result as string;
        if (!rawBase64) {
          reject(new Error('Gagal membaca data gambar.'));
          return;
        }
        const img = new Image();
        img.onload = () => {
          try {
            const canvas = document.createElement('canvas');
            const size = 320; // 320x320 optimal square avatar
            canvas.width = size;
            canvas.height = size;
            const ctx = canvas.getContext('2d');
            if (!ctx) {
              resolve(rawBase64);
              return;
            }
            // Center-crop to perfect square
            const minDim = Math.min(img.width, img.height);
            const sx = (img.width - minDim) / 2;
            const sy = (img.height - minDim) / 2;
            ctx.drawImage(img, sx, sy, minDim, minDim, 0, 0, size, size);
            
            const compressed = canvas.toDataURL('image/jpeg', 0.85);
            resolve(compressed);
          } catch (e) {
            console.warn('Canvas crop fallback:', e);
            resolve(rawBase64);
          }
        };
        img.onerror = () => reject(new Error('Format gambar tidak didukung atau file rusak.'));
        img.src = rawBase64;
      };
      reader.onerror = () => reject(new Error('Gagal membaca file gambar.'));
      reader.readAsDataURL(file);
    });
  };

  const handleProfilePhotoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploadingPhoto(true);
    try {
      const compressedBase64 = await compressAndCropAvatar(file);

      // 1. Optimistically update local state & preview immediately
      setAvatarUrl(compressedBase64);

      // 2. Persist to localStorage immediately
      safeSetLocalStorage('kaguci_avatar_current', compressedBase64);
      const username = (activeUserCustomData?.username || '').toLowerCase().trim();
      if (username) {
        safeSetLocalStorage(`kaguci_avatar_${username}`, compressedBase64);
      }
      const activeUid = currentUser?.uid || activeAuth.currentUser?.uid;
      if (activeUid) {
        safeSetLocalStorage(`kaguci_avatar_${activeUid}`, compressedBase64);
      }

      // 3. Save to Firestore (custom_accounts & users)
      const saveOps: Promise<void>[] = [];
      if (username) {
        trackOp('write', 1);
        saveOps.push(
          setDoc(doc(dbDefault, 'custom_accounts', username), { photoURL: compressedBase64 }, { merge: true })
            .catch(err => console.warn('Gagal menyimpan foto ke custom_accounts:', err))
        );
      }
      if (activeUid) {
        trackOp('write', 1);
        saveOps.push(
          setDoc(doc(activeDb, 'users', activeUid), { photoURL: compressedBase64 }, { merge: true })
            .catch(err => console.warn('Gagal menyimpan foto ke users activeDb:', err))
        );
        if (activeDb !== dbDefault) {
          saveOps.push(
            setDoc(doc(dbDefault, 'users', activeUid), { photoURL: compressedBase64 }, { merge: true })
              .catch(err => console.warn('Gagal menyimpan foto ke users dbDefault:', err))
          );
        }
      }

      // Safe update Firebase Auth profile (ignore URL length limits quietly)
      const authUser = currentUser || activeAuth.currentUser;
      if (authUser) {
        updateProfile(authUser, { photoURL: compressedBase64 }).catch(() => {});
      }

      await Promise.all(saveOps);
      showToast('Foto profil berhasil disimpan dan diperbarui!', 'success');
    } catch (err: unknown) {
      console.error('Error uploading profile photo:', err);
      const msg = err instanceof Error ? err.message : 'Gagal memperbarui foto profil.';
      showToast(msg, 'error');
    } finally {
      setIsUploadingPhoto(false);
      if (profilePhotoInputRef.current) {
        profilePhotoInputRef.current.value = '';
      }
    }
  };

  const handleRemoveProfilePhoto = async () => {
    if (!avatarUrl) return;
    setIsUploadingPhoto(true);
    try {
      setAvatarUrl(null);
      localStorage.removeItem('kaguci_avatar_current');
      const username = (activeUserCustomData?.username || '').toLowerCase().trim();
      if (username) {
        localStorage.removeItem(`kaguci_avatar_${username}`);
      }
      const activeUid = currentUser?.uid || activeAuth.currentUser?.uid;
      if (activeUid) {
        localStorage.removeItem(`kaguci_avatar_${activeUid}`);
      }

      const removeOps: Promise<void>[] = [];
      if (username) {
        trackOp('write', 1);
        removeOps.push(
          setDoc(doc(dbDefault, 'custom_accounts', username), { photoURL: deleteField() }, { merge: true })
            .catch(err => console.warn('Gagal menghapus foto dari custom_accounts:', err))
        );
      }
      if (activeUid) {
        trackOp('write', 1);
        removeOps.push(
          setDoc(doc(activeDb, 'users', activeUid), { photoURL: deleteField() }, { merge: true })
            .catch(err => console.warn('Gagal menghapus foto dari users activeDb:', err))
        );
        if (activeDb !== dbDefault) {
          removeOps.push(
            setDoc(doc(dbDefault, 'users', activeUid), { photoURL: deleteField() }, { merge: true })
              .catch(err => console.warn('Gagal menghapus foto dari users dbDefault:', err))
          );
        }
      }

      await Promise.all(removeOps);
      showToast('Foto profil berhasil dihapus.', 'info');
    } catch (err: unknown) {
      console.error('Gagal menghapus foto profil:', err);
      const msg = err instanceof Error ? err.message : 'Gagal menghapus foto profil.';
      showToast(msg, 'error');
    } finally {
      setIsUploadingPhoto(false);
      if (profilePhotoInputRef.current) {
        profilePhotoInputRef.current.value = '';
      }
    }
  };
  const [importResult, setImportResult] = useState<{
    isOpen: boolean;
    successCount: number;
    skipCount: number;
    failCount: number;
    emptyCount: number;
    totalParsed: number;
    error?: boolean;
    errorMessage?: string;
    details?: string[];
    sheetsProcessed?: { name: string; count: number }[];
  } | null>(null);

  const [isGdriveModalOpen, setIsGdriveModalOpen] = useState(false);
  const [gdriveTarget, setGdriveTarget] = useState<'student' | 'teacher'>('student');
  const [gdriveUrl, setGdriveUrl] = useState('');
  const [isGdriveImporting, setIsGdriveImporting] = useState(false);

  const handleImportFromGdrive = async () => {
    if (!gdriveUrl.trim()) {
      showToast('Mohon masukkan link Google Drive atau Google Sheets yang valid.', 'error');
      return;
    }

    setIsGdriveImporting(true);
    showToast('Sedang menghubungkan ke Google Drive & mengunduh file Excel...', 'info');

    try {
      let targetUrl = gdriveUrl.trim();
      
      // Convert Google Sheets URL to export xlsx
      if (targetUrl.includes('docs.google.com/spreadsheets')) {
        const match = targetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (match && match[1]) {
          targetUrl = `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=xlsx`;
        }
      } 
      // Convert Google Drive file share URL to direct download
      else if (targetUrl.includes('drive.google.com/file/d/')) {
        const match = targetUrl.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
        if (match && match[1]) {
          targetUrl = `https://drive.google.com/uc?export=download&id=${match[1]}`;
        }
      } else if (targetUrl.includes('drive.google.com/open?id=')) {
        const match = targetUrl.match(/id=([a-zA-Z0-9-_]+)/);
        if (match && match[1]) {
          targetUrl = `https://drive.google.com/uc?export=download&id=${match[1]}`;
        }
      }

      let response: Response;
      try {
        response = await fetch(targetUrl);
        if (!response.ok) throw new Error('Direct fetch failed');
      } catch {
        const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(targetUrl)}`;
        response = await fetch(proxyUrl);
        if (!response.ok) throw new Error('CORS proxy fetch failed');
      }

      const blob = await response.blob();
      const fileName = gdriveTarget === 'teacher' ? "gdrive_teachers.xlsx" : "gdrive_students.xlsx";
      const file = new File([blob], fileName, { 
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
      });

      setIsGdriveModalOpen(false);
      setGdriveUrl('');
      setIsGdriveImporting(false);

      if (gdriveTarget === 'teacher') {
        await importTeacherExcelHelper(
          file,
          activeAuth,
          activeDb,
          setImportResult,
          showToast,
          excelTeacherInputRef,
          teachers,
          currentUser?.uid
        );
      } else {
        await importExcelHelper(
          file,
          activeAuth,
          activeDb,
          classList,
          setClassList,
          setImportResult,
          showToast,
          excelInputRef,
          students,
          currentUser?.uid
        );
      }
    } catch (err) {
      console.error(err);
      setIsGdriveImporting(false);
      showToast('Gagal mengunduh file dari Google Drive. Pastikan link dapat diakses publik ("Siapa saja yang memiliki link dapat melihat").', 'error');
    }
  };

  const attendanceStats = useMemo(() => {
    if (attendanceSessions.length === 0) return { rate: '100%', attentionCount: 0 };
    
    // Set ID siswa aktif terdaftar
    const validStudentIds = new Set(students.map(s => s.id));
    
    let totalRecords = 0;
    let totalHadir = 0;
    const attentionSet = new Set<string>();
    
    attendanceSessions.forEach(session => {
      if (session.records) {
        Object.entries(session.records).forEach(([studentId, status]) => {
          if (status) {
            // Hanya hitung rekam absensi dari siswa aktif terdaftar
            if (validStudentIds.size === 0 || validStudentIds.has(studentId)) {
              totalRecords++;
              if (status === 'Hadir' || status === 'Dispen') {
                totalHadir++;
              }
              if (status === 'Alpa' || status === 'Sakit' || status === 'Izin') {
                attentionSet.add(studentId);
              }
            }
          }
        });
      }
    });
    
    const ratePercentage = totalRecords > 0 
      ? Math.round((totalHadir / totalRecords) * 100) 
      : 100;
      
    return {
      rate: `${ratePercentage}%`,
      attentionCount: attentionSet.size
    };
  }, [attendanceSessions, students]);

  const studentAbsenceStats = useMemo(() => {
    // Map student ID to their attendance counts
    const counts: Record<string, { alpa: number; sakit: number; izin: number; totalNonHadir: number; alpaDates: string[]; sakitDates: string[]; izinDates: string[] }> = {};
    
    // Initialize for all students
    students.forEach(s => {
      counts[s.id] = { alpa: 0, sakit: 0, izin: 0, totalNonHadir: 0, alpaDates: [], sakitDates: [], izinDates: [] };
    });
    
    // Process sessions
    attendanceSessions.forEach(session => {
      if (session.records) {
        Object.entries(session.records).forEach(([studentId, status]) => {
          if (!counts[studentId]) {
            counts[studentId] = { alpa: 0, sakit: 0, izin: 0, totalNonHadir: 0, alpaDates: [], sakitDates: [], izinDates: [] };
          }
          if (status === 'Alpa') {
            counts[studentId].alpa++;
            counts[studentId].alpaDates.push(session.date);
            counts[studentId].totalNonHadir++;
          } else if (status === 'Sakit') {
            counts[studentId].sakit++;
            counts[studentId].sakitDates.push(session.date);
            counts[studentId].totalNonHadir++;
          } else if (status === 'Izin') {
            counts[studentId].izin++;
            counts[studentId].izinDates.push(session.date);
            counts[studentId].totalNonHadir++;
          }
        });
      }
    });

    interface StudentAbsenceDetail {
      id: string; 
      name: string; 
      class: string; 
      alpa: number; 
      sakit: number; 
      izin: number; 
      totalNonHadir: number;
      alpaDates: string[];
      sakitDates: string[];
      izinDates: string[];
    }

    // Group students by class
    const classGroups: Record<string, StudentAbsenceDetail[]> = {};

    // Initialize arrays for each class
    classList.forEach(cls => {
      classGroups[cls] = [];
    });

    // Distribute students to classes and attach counts
    students.forEach(s => {
      const cls = s.class || 'Tanpa Kelas';
      if (!classGroups[cls]) {
        classGroups[cls] = [];
      }
      const c = counts[s.id] || { alpa: 0, sakit: 0, izin: 0, totalNonHadir: 0, alpaDates: [], sakitDates: [], izinDates: [] };
      classGroups[cls].push({
        id: s.id,
        name: s.name,
        class: cls,
        ...c
      });
    });

    // For each class, find student with max Alpa and max Sakit+Izin
    const result: Array<{
      className: string;
      totalAlpaStudents: number;
      totalSakitStudents: number;
      totalIzinStudents: number;
      allAbsenceList: Array<{
        id: string;
        name: string;
        alpa: number;
        sakit: number;
        izin: number;
        total: number;
        alpaDates: string[];
        sakitDates: string[];
        izinDates: string[];
      }>;
    }> = [];

    (Object.entries(classGroups) as Array<[string, StudentAbsenceDetail[]]>).forEach(([className, list]) => {


      const absentStudents = list
        .filter(item => item.totalNonHadir > 0)
        .map(item => ({
          id: item.id,
          name: item.name,
          alpa: item.alpa,
          sakit: item.sakit,
          izin: item.izin,
          total: item.totalNonHadir,
          alpaDates: item.alpaDates,
          sakitDates: item.sakitDates,
          izinDates: item.izinDates
        }))
        .sort((a, b) => b.total - a.total);

      const totalAlpaStudents = list.filter(item => item.alpa > 0).length;
      const totalSakitStudents = list.filter(item => item.sakit > 0).length;
      const totalIzinStudents = list.filter(item => item.izin > 0).length;

      result.push({
        className,
        totalAlpaStudents,
        totalSakitStudents,
        totalIzinStudents,
        allAbsenceList: absentStudents
      });
    });

    result.sort((a, b) => a.className.localeCompare(b.className, 'id-ID', { numeric: true }));

    return result;
  }, [students, attendanceSessions, classList]);

  const [toast, setToast] = useState<{ message: string, type: 'success' | 'info' | 'error' } | null>(null);
  const toastTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [isForgotPassword, setIsForgotPassword] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  // States for new school registration modal (Kotak Dialog)
  const [isRegisterModalOpen, setIsRegisterModalOpen] = useState(false);
  const [regFullName, setRegFullName] = useState('');
  const [regUsername, setRegUsername] = useState('');
  const [regPassword, setRegPassword] = useState('');
  const [regToken, setRegToken] = useState('');
  const [generatedToken, setGeneratedToken] = useState('');

  // States for registration result dialog status
  const [registrationResult, setRegistrationResult] = useState<{
    success: boolean;
    title: string;
    message: string;
    fullname?: string;
    username?: string;
    projectId?: string;
    showBypassButton?: boolean;
    configToSave?: {
      fullname: string;
      username: string;
      password?: string;
      configText: string;
    };
  } | null>(null);

  // States for login error dialog (modern message box for unregistered accounts / wrong password)
  const [loginError, setLoginError] = useState<{
    title: string;
    message: string;
    recommendations: string[];
    username?: string;
  } | null>(null);

  // States for password/username recovery
  const [recoverySearchVal, setRecoverySearchVal] = useState('');
  const [recoverySearchType, setRecoverySearchType] = useState<'username' | 'password'>('username');
  const [recoveryResult, setRecoveryResult] = useState<DocumentData[] | null>(null);
  const [isRecoveryLoading, setIsRecoveryLoading] = useState(false);

  const showToast = (message: string, type: 'success' | 'info' | 'error') => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    setToast({ message, type });
    toastTimeoutRef.current = setTimeout(() => setToast(null), 3000);
  };

  useEffect(() => {
    const handleOnline = () => {
      setIsOnline(true);
      showToast('Koneksi internet terhubung kembali! Data disinkronkan ke Cloud.', 'success');
    };
    const handleOffline = () => {
      setIsOnline(false);
      showToast('Koneksi terputus. Menggunakan data lokal (Offline Mode).', 'error');
    };
    const handleQuota = () => {
      setQuotaExceeded(true);
      setShowQuotaModal(true);
    };

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);
    window.addEventListener('firestore-quota-exceeded', handleQuota);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
      window.removeEventListener('firestore-quota-exceeded', handleQuota);
    };
  }, []);

  // Initialize and ensure default admin account exists in Firebase
  useEffect(() => {
    const initDefaultAdmin = async () => {
      try {
        const adminRef = doc(dbDefault, 'custom_accounts', 'admin');
        const snap = await getDoc(adminRef);
        if (!snap.exists() || snap.data()?.fullname === 'Administrator') {
          await setDoc(adminRef, {
            fullname: 'Endang Sukmaya,S.Pd.I., Gr.',
            username: 'admin',
            password: '123456@#',
            role: 'Admin',
            configText: '',
            createdAt: snap.exists() ? (snap.data()?.createdAt || new Date().toISOString()) : new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }, { merge: true });
        }
      } catch (err) {
        console.warn('Init default admin notice:', err);
      }
    };
    initDefaultAdmin();
  }, []);

  const handleForceRegister = async (configData: { fullname: string; username: string; password?: string; configText: string }) => {
    setIsAuthLoading(true);
    setRegistrationResult(null);
    try {
      const resolvedUsername = configData.username.toLowerCase().trim();

      // 1. Create native account inside the main Firebase Cloud Database
      await safeFirebaseAuth(auth, resolvedUsername, configData.password || '', resolvedUsername);

      // 2. Map directory entry centrally to support login routing & forgot password lookups
      await setDoc(doc(dbDefault, 'custom_accounts', resolvedUsername), {
        fullname: configData.fullname,
        username: resolvedUsername,
        password: configData.password,
        configText: configData.configText,
        createdAt: new Date().toISOString()
      });

      // On success: trigger results popup
      setRegistrationResult({
        success: true,
        title: 'Pendaftaran Akun Berhasil!',
        message: 'Akun Administrator baru Anda siap digunakan dengan Cloud Database FRESH secara instan.',
        fullname: configData.fullname,
        username: resolvedUsername,
        projectId: 'Database Pusat Default FRESH'
      });

      // Autofill login credentials for easy access
      setAuthEmail(resolvedUsername);
      if (configData.password) {
        setAuthPassword(configData.password);
      }
      
      // Clean fields and close form
      setIsRegisterModalOpen(false);
      setRegFullName('');
      setRegUsername('');
      setRegPassword('');
    } catch (error) {
      const err = error as Error;
      let IndonesianError = err.message;
      if (err.message.includes('email-already-in-use')) {
        const lookupUsername = configData.username.toLowerCase().trim();
        let centralUserFound: { fullname?: string; username?: string; configText?: string } | null = null;
        try {
          const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', lookupUsername));
          if (lookupDoc.exists()) {
            const data = lookupDoc.data();
            centralUserFound = {
              fullname: data.fullname,
              username: data.username,
              configText: data.configText
            };
          }
        } catch (lookupErr) {
          console.error('Error lookup di catch force:', lookupErr);
        }

        if (centralUserFound) {
          IndonesianError = `Maaf, Username "${lookupUsername}" sudah terdaftar di database sistem pusat.\n\n• Nama Pengguna: ${centralUserFound.fullname || 'N/A'}\n• Akun (Username): ${centralUserFound.username || 'N/A'}\n\nSilakan gunakan menu "Masuk" (Login) dan gunakan akun tersebut beserta kata sandinya untuk login.`;
        } else {
          IndonesianError = `Username "${lookupUsername}" sudah pernah didaftarkan pada database project Firebase Anda, namun kata sandi yang Anda ketik salah.\n\nLangkah Solusi:\n1. Masukkan kata sandi yang tepat jika Anda adalah pemilik akun tersebut.\n2. ATAU, buka Firebase Console Anda -> menu Authentication -> hapus akun "${lookupUsername}@kaguci.com", setelah itu coba daftarkan kembali.\n3. ATAU, silakan daftar dengan memakai Username yang berbeda.`;
        }
      } else if (err.message.includes('weak-password')) {
        IndonesianError = 'Kata sandi minimal berisi 6 karakter.';
      } else if (err.message.includes('invalid-api-key') || err.message.includes('API key')) {
        IndonesianError = 'API Key yang terdapat pada konfigurasi Web Firebase Anda salah atau tidak valid.';
      } else if (err.message.includes('network-request-failed')) {
        IndonesianError = 'Koneksi jaringan gagal atau domain AuthDomain tidak terdaftar di Firebase Anda.';
      } else if (err.message.includes('operation-not-allowed')) {
        IndonesianError = 'Provider "Email/Password" belum aktif di Firebase Console Anda!';
      } else if (err.message.includes('configuration-not-found') || err.message.includes('auth/configuration-not-found')) {
        IndonesianError = 'Layanan Autentikasi belum diinisialisasi di Proyek Firebase Anda! Silakan masuk ke Firebase Console -> klik menu "Authentication" di sebelah kiri -> lalu klik tombol "Get Started" (Mulai) untuk mengaktifkannya.';
      }

      setRegistrationResult({
        success: false,
        title: 'Pendaftaran Gagal!',
        message: `${IndonesianError}\n\n(Detail Teknis: ${err.message})`
      });
    } finally {
      setIsAuthLoading(false);
    }
  };

  const handleEditStudent = (student: Student) => {
    const studentWali = student.waliKelas || getWaliKelasForClass(student.class, classWaliMap, students);
    const studentWaliNiy = student.waliKelasNiy || getWaliKelasNiyForClass(student.class, classWaliNiyMap, students, teachers, studentWali);
    setNewStudent({ 
      name: student.name, 
      nisn: student.nisn, 
      class: student.class,
      waliKelas: studentWali,
      waliKelasNiy: studentWaliNiy
    });
    setEditingStudentId(student.id);
    showToast('Mode edit aktif untuk ' + student.name, 'info');
    
    // Scroll to form if needed
    const formElement = document.getElementById('student-input-container');
    if (formElement) {
      formElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  const handleEditTeacher = (teacher: Teacher) => {
    setNewTeacher({ name: teacher.name, niy: teacher.niy, dutyDay: normalizeDutyDay(teacher.dutyDay || '') });
    setEditingTeacherId(teacher.id);
    showToast('Mode edit aktif untuk ' + teacher.name, 'info');
    
    // Scroll to form if needed
    const formElement = document.getElementById('teacher-input-container');
    if (formElement) {
      formElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  const [deletingStudentId, setDeletingStudentId] = useState<string | null>(null);

  const confirmAction = async () => {
    if (!confirmationAction) return;
    
    if (confirmationAction.type === 'delete') {
      const studentToDelete = confirmationAction.student;
      
      // OPTIMISTIC UI: Close modals and show success feedback immediately
      setConfirmationAction(null);
      setStudentSuccessModal('deleted');
      setDeletingStudentId(null); // Clear deleting state to avoid visual lag
      
      try {
        trackOp('write', 1);
        // Perform deletion in background
        await deleteDoc(doc(activeDb, 'students', studentToDelete.id));
        showToast(`Data ${studentToDelete.name} berhasil dihapus dari cloud.`, 'success');
      } catch (err) {
        console.error("Deletion failed:", err);
        showToast('Sinkronisasi hapus gagal. Cek koneksi internet Anda.', 'error');
      }
    }
  };

  const [teacherConfirmationAction, setTeacherConfirmationAction] = useState<{ type: 'delete', teacher: Teacher } | null>(null);

  const confirmTeacherAction = async () => {
    if (!teacherConfirmationAction) return;
    
    if (teacherConfirmationAction.type === 'delete') {
      const teacherToDelete = teacherConfirmationAction.teacher;
      
      // OPTIMISTIC UI: Close modals and show success feedback immediately
      setTeacherConfirmationAction(null);
      setDeletingTeacherId(null); // Clear deleting state
      
      try {
        trackOp('write', 1);
        // Perform deletion in background
        await deleteDoc(doc(activeDb, 'teachers', teacherToDelete.id));
        showToast(`Data guru ${teacherToDelete.name} berhasil dihapus.`, 'success');
      } catch (err) {
        console.error("Deletion failed:", err);
        showToast('Sinkronisasi hapus gagal. Cek koneksi internet Anda.', 'error');
      }
    }
  };

  const handleResetData = async (type: 'new_semester' | 'new_year' | 'everything' | 'clear_all_students' | 'clear_all_teachers') => {
    if (!currentUser || !activeDb) return;
    setIsResettingData(true);
    setResetProgress(0);
    
    // Start an interval to animate the progress smoothly and extremely fast
    let currentSimulatedProgress = 0;
    const progressInterval = setInterval(() => {
      // Fast, human-like fluid steps
      currentSimulatedProgress += Math.random() > 0.4 ? 6 : 4;
      if (currentSimulatedProgress >= 93) {
        clearInterval(progressInterval);
        currentSimulatedProgress = 93; // hold near 93% until parallel DB writes complete
      }
      setResetProgress(currentSimulatedProgress);
    }, 25);

    const finishUI = async (successType: 'new_semester' | 'new_year' | 'everything' | 'clear_all_students' | 'clear_all_teachers') => {
      clearInterval(progressInterval);
      setResetProgress(100);
      await new Promise(r => setTimeout(r, 450)); // Allow user to see 100%
      setResetModalType('none');
      setResetSuccessModal(successType);
      setResetConfirmInput('');
      setResetPasswordInput('');
      setResetPasswordError('');
      setResetProgress(0);
      setIsResettingData(false);
    };

    try {
      const promises: Promise<void>[] = [];

      if (type === 'new_semester' || type === 'new_year') {
        const sessionsToDelete = [...attendanceSessions];
        if (sessionsToDelete.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < sessionsToDelete.length; i++) {
            batch.delete(doc(activeDb, 'attendanceSessions', sessionsToDelete[i].id));
            count++;
            if (count === 400 || i === sessionsToDelete.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }

        const updatedProfile = { 
          ...profileData, 
          semester: type === 'new_semester' ? newSemesterChoice : profileData.semester,
          ...(type === 'new_semester' && newTahunPelajaran ? { tahunPelajaran: newTahunPelajaran } : {})
        };
        setProfileData(updatedProfile);

        if (activeAuth.currentUser) {
          trackOp('write', 1);
          promises.push(setDoc(doc(activeDb, 'users', activeAuth.currentUser.uid), {
            profileData: updatedProfile,
            semester: updatedProfile.semester,
            ...(updatedProfile.tahunPelajaran ? { tahunPelajaran: updatedProfile.tahunPelajaran } : {})
          }, { merge: true }));
        }
        
        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              attendanceSessions: [],
              profileData: updatedProfile,
              semester: updatedProfile.semester,
              ...(updatedProfile.tahunPelajaran ? { tahunPelajaran: updatedProfile.tahunPelajaran } : {})
            }, { merge: true }));
            localStorage.setItem(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`, JSON.stringify(updatedProfile));
        }

        if (promises.length > 0) {
          try {
            await Promise.all(promises);
          } catch (dbErr) {
            console.error('Reset semester error:', dbErr);
            handleFirestoreError(dbErr, OperationType.DELETE, 'attendanceSessions');
          }
        }
        
        setAttendanceSessions([]);
        if (currentUser) {
          localStorage.removeItem(`kaguci_sessions_${currentUser.uid}`);
        }
        showToast(type === 'new_semester' ? `Semester Baru (${newSemesterChoice}) berhasil disiapkan.` : 'Pembersihan Tahun Ajaran baru telah disiapkan.', 'success');
        await finishUI(type);

      } else if (type === 'everything') {
        const sessionsToDelete = [...attendanceSessions];
        const studentsToDelete = [...students];
        const teachersToDelete = [...teachers];
        const allDeletes = [
          ...sessionsToDelete.map(s => doc(activeDb, 'attendanceSessions', s.id)),
          ...studentsToDelete.map(s => doc(activeDb, 'students', s.id)),
          ...teachersToDelete.map(t => doc(activeDb, 'teachers', t.id))
        ];
        
        if (allDeletes.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < allDeletes.length; i++) {
            batch.delete(allDeletes[i]);
            count++;
            if (count === 400 || i === allDeletes.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }
        
        if (currentUser) {
          trackOp('write', 1);
          promises.push(setDoc(doc(activeDb, 'users', currentUser.uid), { 
            classList: [], 
            photoURL: deleteField(),
            profileData: {
              namaGuruMapel: '', nipGuruMapel: '', mataPelajaran: '',
              namaKepalaSekolah: '', nipKepalaSekolah: '',
              semester: 'Ganjil', tahunPelajaran: ''
            }
          }, { merge: true }));
        }



        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              classList: [], students: [], attendanceSessions: [], teachers: [],
              photoURL: deleteField(),
              profileData: {
                namaGuruMapel: '', nipGuruMapel: '', mataPelajaran: '',
                namaKepalaSekolah: '', nipKepalaSekolah: '',
                semester: 'Ganjil', tahunPelajaran: ''
              }
            }, { merge: true }));
        }
        
        if (promises.length > 0) {
          try {
            await Promise.all(promises);
          } catch (dbErr) {
            console.error('Reset everything error:', dbErr);
            handleFirestoreError(dbErr, OperationType.DELETE, 'everything');
          }
        }

        setClassList([]);
        setStudents([]);
        setTeachers([]);
        setAttendanceSessions([]);
        setAvatarUrl(null);
        setProfileData({
          namaGuruMapel: '', nipGuruMapel: '', mataPelajaran: '',
          namaKepalaSekolah: '', nipKepalaSekolah: '',
          semester: 'Ganjil', tahunPelajaran: '',
          role: 'Guru Mapel',
          hariPiket: '', waliKelasClass: ''
        });
        
        if (currentUser) {
          localStorage.removeItem(`kaguci_students_${currentUser.uid}`);
          localStorage.removeItem(`kaguci_teachers_${currentUser.uid}`);
          localStorage.removeItem(`kaguci_classList_${currentUser.uid}`);
          localStorage.removeItem(`kaguci_avatar_${currentUser.uid}`);
          localStorage.removeItem(`kaguci_sessions_${currentUser.uid}`);
        }

        showToast('Seluruh data berhasil dihapus secara permanen.', 'success');
        await finishUI('everything');

      } else if (type === 'clear_all_students') {
        const studentsToDelete = [...students];
        if (studentsToDelete.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < studentsToDelete.length; i++) {
            batch.delete(doc(activeDb, 'students', studentsToDelete[i].id));
            count++;
            if (count === 400 || i === studentsToDelete.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }
        
        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              students: [] 
            }, { merge: true }));
        }
        
        if (promises.length > 0) {
          try {
            await Promise.all(promises);
          } catch (dbErr) {
            console.error('Clear students error:', dbErr);
            handleFirestoreError(dbErr, OperationType.DELETE, 'students');
          }
        }
        
        setStudents([]);
        if (currentUser) {
          localStorage.removeItem(`kaguci_students_${currentUser.uid}`);
        }
        showToast('Data siswa berhasil dikosongkan.', 'success');
        await finishUI('clear_all_students');

      } else if (type === 'clear_all_teachers') {
        const teachersToDelete = [...teachers];
        if (teachersToDelete.length > 0) {
          let batch = writeBatch(activeDb);
          let count = 0;
          for (let i = 0; i < teachersToDelete.length; i++) {
            batch.delete(doc(activeDb, 'teachers', teachersToDelete[i].id));
            count++;
            if (count === 400 || i === teachersToDelete.length - 1) {
              trackOp('write', count);
              promises.push(batch.commit());
              batch = writeBatch(activeDb);
              count = 0;
            }
          }
        }
        
        if (activeUserCustomData?.username) {
            trackOp('write', 1);
            promises.push(setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), { 
              teachers: [] 
            }, { merge: true }));
        }
        
        if (promises.length > 0) {
          try {
            await Promise.all(promises);
          } catch (dbErr) {
            console.error('Clear teachers error:', dbErr);
            handleFirestoreError(dbErr, OperationType.DELETE, 'teachers');
          }
        }
        
        setTeachers([]);
        if (currentUser) {
          localStorage.removeItem(`kaguci_teachers_${currentUser.uid}`);
        }
        showToast('Data guru berhasil dikosongkan.', 'success');
        await finishUI('clear_all_teachers');
      }
    } catch (err) {
      clearInterval(progressInterval);
      console.error("Error resetting data:", err);
      showToast('Sinkronisasi cloud sedang sibuk. Mohon tunggu sejenak.', 'info');
      setIsResettingData(false);
      setResetProgress(0);
    }
  };

  const handleSaveClassWali = async (className: string, waliName: string, waliNiy: string) => {
    const updatedWaliMap = {
      ...(classWaliMap || {}),
      [className]: waliName.trim()
    };
    const updatedWaliNiyMap = {
      ...(classWaliNiyMap || {}),
      [className]: waliNiy.trim()
    };
    setClassWaliMap(updatedWaliMap);
    setClassWaliNiyMap(updatedWaliNiyMap);

    const uid = currentUser?.uid || activeAuth.currentUser?.uid;
    if (uid) {
      try {
        trackOp('write', 1);
        await setDoc(doc(activeDb, 'users', uid), {
          classWaliMap: updatedWaliMap,
          classWaliNiyMap: updatedWaliNiyMap
        }, { merge: true });
        showToast(`Data Wali Kelas ${className} berhasil disimpan!`, 'success');
      } catch (err) {
        console.error('Error saving class wali:', err);
        showToast('Gagal menyimpan data wali kelas ke database.', 'error');
      }
    } else {
      showToast(`Data Wali Kelas ${className} berhasil disimpan!`, 'success');
    }
  };


  const addOrUpdateStudent = async () => {
    if (!newStudent.name || !newStudent.nisn || !newStudent.class) {
      showToast('Mohon lengkapi semua data siswa (Nama, NISN, Kelas).', 'error');
      return;
    }
    
    // Cek apakah NISN sudah terdaftar (double)
    const isDuplicateNisn = students.some(s => s.nisn === newStudent.nisn && s.id !== editingStudentId);
    if (isDuplicateNisn) {
      showToast('NISN / NIS sudah terdaftar.', 'error');
      return;
    }
    
    const uid = currentUser?.uid || activeAuth.currentUser?.uid || 'admin';

    try {
      let finalWali = newStudent.waliKelas?.trim() || '';
      let finalWaliNiy = newStudent.waliKelasNiy?.trim() || '';

      if (!finalWali) {
        finalWali = getWaliKelasForClass(newStudent.class.trim(), classWaliMap, students) || '';
      }
      if (!finalWaliNiy) {
        finalWaliNiy = getWaliKelasNiyForClass(newStudent.class.trim(), classWaliNiyMap, students, teachers, finalWali) || '';
      }

      const studentPayload: Partial<Student> = {
        name: newStudent.name.trim(),
        nisn: newStudent.nisn.trim(),
        class: newStudent.class.trim(),
        userId: uid
      };

      if (finalWali) {
        studentPayload.waliKelas = finalWali;
        const updatedWaliMap = {
          ...(classWaliMap || {}),
          [newStudent.class.trim()]: finalWali
        };
        setClassWaliMap(updatedWaliMap);
        if (currentUser) {
          setDoc(doc(activeDb, 'users', currentUser.uid), {
            classWaliMap: updatedWaliMap
          }, { merge: true }).catch(err => console.warn('Failed to save classWaliMap:', err));
        }

      }

      if (finalWaliNiy) {
        studentPayload.waliKelasNiy = finalWaliNiy;
        const updatedWaliNiyMap = {
          ...(classWaliNiyMap || {}),
          [newStudent.class.trim()]: finalWaliNiy
        };
        setClassWaliNiyMap(updatedWaliNiyMap);
        if (currentUser) {
          setDoc(doc(activeDb, 'users', currentUser.uid), {
            classWaliNiyMap: updatedWaliNiyMap
          }, { merge: true }).catch(err => console.warn('Failed to save classWaliNiyMap:', err));
        }

      }

      if (editingStudentId) {
        const studentData = { ...studentPayload, id: editingStudentId } as Student;
        // Reset form immediately for fast feel
        setEditingStudentId(null);
        setNewStudent({ name: '', nisn: '', class: '', waliKelas: '', waliKelasNiy: '' });
        setStudentSuccessModal('edited');
        
        await setDoc(doc(activeDb, 'students', editingStudentId), studentData, { merge: true });
        showToast('Data siswa berhasil diperbarui.', 'success');
      } else {
        const newId = Date.now().toString();
        const studentData = { ...studentPayload, id: newId } as Student;
        // Reset form immediately
        setNewStudent(prev => ({ ...prev, name: '', nisn: '', waliKelas: '', waliKelasNiy: '' }));
        setStudentSuccessModal('added');
        
        await setDoc(doc(activeDb, 'students', newId), studentData);
        showToast('Siswa Berhasil Ditambahkan', 'success');
      }
    } catch (err) {
      const errorObj = err as Error;
      showToast('Gagal menyimpan: ' + (errorObj.message || 'Server error'), 'error');
      handleFirestoreError(err, OperationType.WRITE, 'students');
    }
  };

  const addOrUpdateTeacher = async () => {
    if (!newTeacher.name || !newTeacher.name.trim()) {
      showToast('Mohon masukkan nama guru.', 'error');
      return;
    }

    const trimmedName = newTeacher.name.trim();
    const rawNiy = newTeacher.niy ? newTeacher.niy.trim() : '';
    const formattedNiy = rawNiy === '' || rawNiy === '-' ? '-' : rawNiy;
    const formattedDutyDay = normalizeDutyDay(newTeacher.dutyDay || '');

    // Check duplicate NIY only if NIY is not '-'
    if (formattedNiy !== '-') {
      const isDuplicateNiy = teachers.some(t => t.niy && t.niy.trim() !== '-' && t.niy.trim() === formattedNiy && t.id !== editingTeacherId);
      if (isDuplicateNiy) {
        showToast('NIY sudah terdaftar.', 'error');
        return;
      }
    }
    
    const uid = currentUser?.uid || activeAuth.currentUser?.uid || 'admin';

    try {
      if (editingTeacherId) {
        const teacherData: Teacher = { 
          id: editingTeacherId, 
          name: trimmedName, 
          niy: formattedNiy, 
          dutyDay: formattedDutyDay,
          userId: uid 
        };
        setEditingTeacherId(null);
        setNewTeacher({ name: '', niy: '', dutyDay: '' });
        
        await setDoc(doc(activeDb, 'teachers', editingTeacherId), teacherData);
        showToast('Data guru berhasil diperbarui.', 'success');
      } else {
        const newId = Date.now().toString();
        const teacherData: Teacher = { 
          id: newId, 
          name: trimmedName, 
          niy: formattedNiy, 
          dutyDay: formattedDutyDay,
          userId: uid 
        };
        setNewTeacher({ name: '', niy: '', dutyDay: '' });
        
        await setDoc(doc(activeDb, 'teachers', newId), teacherData);
        showToast('Guru Berhasil Ditambahkan', 'success');
      }
    } catch (err) {
      const errorObj = err as Error;
      showToast('Gagal menyimpan: ' + (errorObj.message || 'Server error'), 'error');
      handleFirestoreError(err, OperationType.WRITE, 'teachers');
    }
  };

  const excelInputRef = useRef<HTMLInputElement>(null);

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    await importExcelHelper(
      file,
      activeAuth,
      activeDb,
      classList,
      setClassList,
      setImportResult,
      showToast,
      excelInputRef,
      students,
      currentUser?.uid
    );
  };

  const excelTeacherInputRef = useRef<HTMLInputElement>(null);

  const handleImportTeacherExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    await importTeacherExcelHelper(
      file,
      activeAuth,
      activeDb,
      setImportResult,
      showToast,
      excelTeacherInputRef,
      teachers,
      currentUser?.uid
    );
  };

  const menuItems = [
    { 
      id: 'dashboard', 
      label: 'Utama', 
      icon: LayoutGrid, 
      color: 'text-[#098f41]', 
      activeBg: 'bg-emerald-50/90 text-[#077a37] ring-1 ring-[#098f41]/30',
      iconBg: 'bg-[#098f41] text-white',
      hoverBg: 'hover:bg-emerald-50/50',
      inactiveBg: 'bg-emerald-50/60',
      inactiveColor: 'text-[#098f41]'
    },
    { 
      id: 'attendance', 
      label: 'Presensi', 
      icon: Fingerprint, 
      color: 'text-[#098f41]', 
      activeBg: 'bg-emerald-50/80 text-lime-700 ring-1 ring-emerald-200/50',
      iconBg: 'bg-[#098f41] text-white',
      hoverBg: 'hover:bg-emerald-50/40',
      inactiveBg: 'bg-emerald-50/50',
      inactiveColor: 'text-[#077a37]'
    },
    { 
      id: 'students', 
      label: 'Input Data', 
      icon: Users, 
      color: 'text-sky-500', 
      activeBg: 'bg-sky-50/80 text-sky-700 ring-1 ring-sky-200/50',
      iconBg: 'bg-sky-500 text-white',
      hoverBg: 'hover:bg-sky-50/40',
      inactiveBg: 'bg-sky-50/50',
      inactiveColor: 'text-sky-500'
    },
    { 
      id: 'reports', 
      label: 'Laporan', 
      icon: BarChart3, 
      color: 'text-rose-500', 
      activeBg: 'bg-rose-50/80 text-rose-700 ring-1 ring-rose-200/50',
      iconBg: 'bg-rose-500 text-white',
      hoverBg: 'hover:bg-rose-50/40',
      inactiveBg: 'bg-rose-50/50',
      inactiveColor: 'text-rose-500'
    }
  ];

  const renderContent = () => {
    switch (activeTab) {
      case 'dashboard':
        return (
          <div className="p-4 sm:p-6 lg:p-8 space-y-6 lg:space-y-8 max-w-7xl mx-auto">
            
            <AnimatePresence>
              {showIdleTip && (
                <motion.div
                  initial={{ opacity: 0, y: -20, height: 0 }}
                  animate={{ opacity: 1, y: 0, height: 'auto' }}
                  exit={{ opacity: 0, y: -20, height: 0 }}
                  transition={{ duration: 0.3 }}
                  className="overflow-hidden"
                >
                  <div className="bg-sky-50 border border-sky-100 rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row items-start sm:items-center gap-4 text-left shadow-sm mb-6">
                    <div className="w-10 h-10 shrink-0 bg-sky-100 text-sky-600 rounded-full flex items-center justify-center">
                      <Cloud className="w-5 h-5" />
                    </div>
                    <div className="flex-1">
                      <h3 className="text-sm font-bold text-sky-800 mb-1">Tips Menghemat Kuota Database Harian</h3>
                      <p className="text-xs text-sky-700 leading-relaxed">
                        Apabila aplikasi sedang tidak digunakan dalam waktu lama, harap menekan tombol <strong>Keluar (Logout)</strong> yang berada di sudut kanan atas.
                        Hal ini sangat penting untuk mengurangi aktivitas sinkronisasi di latar belakang sehingga <strong>Limit Kuota Harian (Free Tier)</strong> database Anda tidak cepat habis.
                      </p>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Top grid: Welcome Banner & Quick Action */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              {/* Welcome Card (spans 2 on desktop) */}
              <div className="lg:col-span-2 bg-gradient-to-br from-[#098f41] flex flex-col justify-center to-[#077a37] rounded-[2rem] p-6 sm:p-8 text-white shadow-lg shadow-[#098f41]/25 border-2 border-[#098f41] relative overflow-hidden min-h-[240px]">
                <div className="absolute top-0 right-0 p-8 flex items-center justify-center opacity-10">
                  <GraduationCap className="w-48 h-48 sm:w-64 sm:h-64 -rotate-12 transform" />
                </div>
                
                <div className="relative z-10 w-full mb-4 flex justify-start">
                  <div className="bg-white/20 backdrop-blur-md px-3 py-1.5 rounded-full text-[10px] sm:text-xs font-bold text-white flex items-center gap-1.5 border border-white/20 shadow-sm">
                    <Cloud className="w-3.5 h-3.5" />
                    Database Project: {activeAuth.app.options.projectId || 'N/A'}
                  </div>
                </div>

                <div className="relative z-10 w-full sm:w-3/4">
                  <p className="text-emerald-100 text-xs sm:text-sm font-bold tracking-wider uppercase mb-1 drop-shadow-sm">
                    Halo . {activeUserCustomData?.fullname || 'User'}
                  </p>
                  <h2 className="text-3xl sm:text-4xl font-black mb-3 tracking-tight leading-tight">
                    Selamat Datang di Aplikasi <span className="text-white border-b-2 border-white/40 pb-0.5">SMART DF App</span>
                  </h2>
                  <p className="text-emerald-50 text-base sm:text-lg font-medium opacity-90">Pantau kehadiran, kelola data siswa, dan akses laporan terkini dalam satu tempat.</p>
                </div>
                <div className="relative z-10 mt-8 grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4 max-w-2xl">
                  <button 
                    onClick={() => setActiveTab('attendance')} 
                    className="bg-white text-[#077a37] hover:bg-emerald-50 px-5 py-3.5 rounded-2xl font-extrabold text-sm shadow-md hover:shadow-lg hover:scale-102 active:scale-95 transition-all flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <Fingerprint className="w-5 h-5 shrink-0" /> 
                    <span>Mulai Presensi</span>
                  </button>
                  <button 
                    onClick={handleOpenHomeroomReport} 
                    className="bg-white text-[#077a37] hover:bg-emerald-50 px-5 py-3.5 rounded-2xl font-extrabold text-sm shadow-md hover:shadow-lg hover:scale-102 active:scale-95 transition-all flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <ClipboardList className="w-5 h-5 shrink-0 text-amber-600" /> 
                    <span>Laporan Wali Kelas</span>
                  </button>
                  <button 
                    onClick={() => window.location.reload()} 
                    className="bg-white text-[#077a37] hover:bg-emerald-50 px-5 py-3.5 rounded-2xl font-extrabold text-sm shadow-md hover:shadow-lg hover:scale-102 active:scale-95 transition-all flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <RefreshCw className="w-5 h-5 shrink-0" /> 
                    <span>Refresh Halaman</span>
                  </button>
                </div>
              </div>

              {/* Date & Motivation Widget */}
              <div className="bg-white rounded-[2rem] p-6 sm:p-8 border-2 border-[#098f41]/30 hover:border-[#098f41] shadow-sm flex flex-col justify-center items-center text-center h-full min-h-[240px] transition-all">
                 <div className="w-20 h-20 bg-emerald-50 rounded-full flex items-center justify-center mb-6 text-[#098f41] shrink-0 border border-emerald-100">
                    <Calendar className="w-10 h-10" />
                 </div>
                 <h3 className="text-xl sm:text-2xl font-black text-slate-800 tracking-tight leading-tight">
                   {new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                 </h3>
                 <p className="text-slate-600 font-semibold mt-3 text-sm">Semoga harimu menyenangkan!</p>
              </div>
            </div>

            {/* KPI Summary Grid */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 sm:gap-6">
              {[
                { 
                  label: 'Total Siswa', 
                  value: students.length.toString(), 
                  icon: Users, 
                  color: 'text-indigo-600', 
                  bg: 'bg-indigo-50/70 border-indigo-100', 
                  shadow: 'hover:shadow-indigo-100/60 hover:border-indigo-300' 
                },
                { 
                  label: 'Total Kelas', 
                  value: classList.length.toString(), 
                  icon: Building2, 
                  color: 'text-amber-600', 
                  bg: 'bg-amber-50/70 border-amber-100', 
                  shadow: 'hover:shadow-amber-100/60 hover:border-amber-300' 
                },
                { 
                  label: 'Tingkat Kehadiran', 
                  value: attendanceStats.rate, 
                  icon: Check, 
                  color: 'text-[#098f41]', 
                  bg: 'bg-emerald-50/70 border-emerald-100', 
                  shadow: 'hover:shadow-emerald-100/60 hover:border-emerald-300' 
                },
                { 
                  label: 'Total Siswa Absen', 
                  value: attendanceStats.attentionCount.toString(), 
                  icon: AlertCircle, 
                  color: 'text-rose-600', 
                  bg: 'bg-rose-50/70 border-rose-100', 
                  shadow: 'hover:shadow-rose-100/60 hover:border-rose-300' 
                },
              ].map((item, i) => (
                <div key={i} className={`p-5 sm:p-6 rounded-[2rem] border-2 border-[#098f41]/20 hover:border-[#098f41] bg-white flex flex-col gap-4 sm:gap-6 justify-between items-center sm:items-start text-center sm:text-left transition-all duration-300 hover:-translate-y-1 shadow-sm ${item.shadow}`}>
                  <div className={`p-4 rounded-[1.25rem] w-fit ${item.bg} border`}>
                    <item.icon className={`w-7 h-7 ${item.color}`} />
                  </div>
                  <div className="w-full">
                    <p className="text-3xl sm:text-4xl font-black text-slate-800 tracking-tighter leading-none">{item.value}</p>
                    <p className="text-xs sm:text-sm font-bold text-slate-600 mt-2">{item.label}</p>
                  </div>
                </div>
              ))}
            </div>

            {/* BARU: Menu Informasi & Ringkasan Ketidakhadiran Siswa (Alpa, Sakit, Izin) */}
            <div className="bg-white rounded-[3rem] border-2 border-[#098f41] shadow-xl shadow-[#098f41]/10 flex flex-col overflow-hidden">
              {/* Header Section - Selalu Terlihat */}
              <div className="p-8 sm:p-10 border-b border-slate-100 bg-white shadow-sm z-20 shrink-0">
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-8">
                  <div className="flex-1">
                    <div className="flex items-center gap-4">
                      <div className="p-3 bg-rose-50 rounded-[1.25rem] text-rose-500 shadow-inner">
                        <AlertCircle className="w-7 h-7" />
                      </div>
                      <div>
                        <h2 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight">Informasi Ketidakhadiran Siswa</h2>
                        <p className="text-sm sm:text-base text-slate-600 font-semibold mt-1.5 opacity-80">
                          Analisis riwayat ketidakhadiran terbanyak di setiap kelas.
                        </p>
                      </div>
                    </div>
                  </div>
                  
                  {/* Tab selector */}
                  <div className="flex bg-slate-100 p-1.5 rounded-[1.25rem] border-2 border-slate-300/60 self-start sm:self-auto shrink-0 shadow-inner">
                    <button 
                      onClick={() => setDashboardActiveStatsTab('class_summary')}
                      className={`px-5 py-2.5 rounded-xl text-xs font-black transition-all duration-300 ${
                        dashboardActiveStatsTab === 'class_summary' 
                          ? 'bg-white text-slate-900 shadow-xl shadow-slate-200/50 scale-100' 
                          : 'text-slate-600 hover:text-slate-700 hover:bg-white/40'
                      }`}
                    >
                      Ringkasan Kelas
                    </button>
                    <button 
                      onClick={() => setDashboardActiveStatsTab('top_rankings')}
                      className={`px-5 py-2.5 rounded-xl text-xs font-black transition-all duration-300 ${
                        dashboardActiveStatsTab === 'top_rankings' 
                          ? 'bg-white text-slate-900 shadow-xl shadow-slate-200/50 scale-100' 
                          : 'text-slate-600 hover:text-slate-700 hover:bg-white/40'
                      }`}
                    >
                      Ranking Ketidakhadiran
                    </button>
                  </div>
                </div>
              </div>

              <div className="p-6 sm:p-8 overflow-y-auto max-h-[850px] lg:max-h-[700px] flex-1 custom-scrollbar scroll-smooth">
                {dashboardActiveStatsTab === 'class_summary' ? (
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 pb-4">
                  {studentAbsenceStats.length === 0 ? (
                    <div className="col-span-full py-12 text-center bg-slate-50 rounded-[2rem] border border-dashed border-slate-200">
                      <p className="text-slate-500 font-bold mb-2">Belum ada data kelas yang terdaftar</p>
                      <button onClick={() => setActiveTab('students')} className="text-sm font-bold text-[#077a37] hover:underline">
                        Mulai dengan mengelola kelas & siswa di sini
                      </button>
                    </div>
                  ) : (
                    studentAbsenceStats.map((clsData) => {

                      
                      return (
                        <div key={clsData.className} className="bg-slate-50/50 rounded-2xl p-5 border border-slate-100 flex flex-col justify-between hover:shadow-md transition-all">
                          <div>
                            <div className="flex items-center justify-between mb-4">
                              <span className="bg-[#098f41] text-white px-3.5 py-1 rounded-full text-xs font-black shadow-sm">
                                {clsData.className}
                              </span>
                              <span className="text-xs text-slate-500 font-bold">
                                {clsData.allAbsenceList.length} siswa absen
                              </span>
                            </div>

                            <div className="space-y-3.5">
                              {/* Summary of Absences */}
                              <div className="grid grid-cols-3 gap-2">
                                <div className="bg-white p-3 rounded-xl border border-rose-100 flex flex-col items-center justify-center text-center">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Alpa</span>
                                  <span className="text-lg font-black text-rose-600">{clsData.totalAlpaStudents}</span>
                                  <span className="text-[9px] text-slate-400 font-medium">Siswa</span>
                                </div>
                                <div className="bg-white p-3 rounded-xl border border-amber-100 flex flex-col items-center justify-center text-center">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Sakit</span>
                                  <span className="text-lg font-black text-amber-500">{clsData.totalSakitStudents}</span>
                                  <span className="text-[9px] text-slate-400 font-medium">Siswa</span>
                                </div>
                                <div className="bg-white p-3 rounded-xl border border-indigo-100 flex flex-col items-center justify-center text-center">
                                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Izin</span>
                                  <span className="text-lg font-black text-indigo-500">{clsData.totalIzinStudents}</span>
                                  <span className="text-[9px] text-slate-400 font-medium">Siswa</span>
                                </div>
                              </div>
                            </div>
                          </div>

                          <button 
                            onClick={() => setDashboardSelectedClassDetail(clsData.className)}
                            className="w-full mt-4 bg-white hover:bg-slate-100 border-2 border-slate-300 text-slate-700 py-2 rounded-xl text-xs font-bold transition-colors flex items-center justify-center gap-1.5"
                          >
                            <Info className="w-3.5 h-3.5" /> Detail Absensi Kelas
                          </button>
                        </div>
                      );
                    })
                  )}
                </div>
              ) : (
                <div className="overflow-x-auto rounded-2xl border border-slate-100">
                  <table className="w-full text-left border-collapse bg-slate-50/20">
                    <thead>
                      <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase border-b border-slate-100">
                        <th className="py-4 px-6">Nama Siswa</th>
                        <th className="py-4 px-6">Kelas</th>
                        <th className="py-4 px-6 text-center">Sakit (S)</th>
                        <th className="py-4 px-6 text-center">Izin (I)</th>
                        <th className="py-4 px-6 text-center">Alpa (A)</th>
                        <th className="py-4 px-6 text-center">Total Absen</th>
                        <th className="py-4 px-6 text-right">Tindakan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {(() => {
                        const allAbsentOrdered = studentAbsenceStats
                          .flatMap(c => c.allAbsenceList.map(s => ({ ...s, className: c.className })))
                          .sort((a, b) => b.total - a.total || b.alpa - a.alpa);
                          
                        if (allAbsentOrdered.length === 0) {
                          return (
                            <tr>
                              <td colSpan={7} className="py-12 text-center text-slate-500 font-medium italic bg-white animate-fade">
                                Belum ada siswa berstatus Alpa, Sakit, atau Izin.
                              </td>
                            </tr>
                          );
                        }
                        
                        return allAbsentOrdered.map((student, rank) => (
                          <tr key={student.id} className="hover:bg-slate-50/50 bg-white transition-colors">
                            <td className="py-4 px-6 font-bold text-slate-800 flex items-center gap-2">
                              <span className={`w-5 h-5 flex items-center justify-center rounded-full text-[10px] font-black shrink-0 ${
                                rank === 0 ? 'bg-rose-500 text-white' : 
                                rank === 1 ? 'bg-amber-500 text-white' : 
                                rank === 2 ? 'bg-indigo-500 text-white' : 'bg-slate-200 text-slate-600'
                              }`}>
                                {rank + 1}
                              </span>
                              {student.name}
                            </td>
                            <td className="py-4 px-6 text-slate-600 font-bold">{student.className}</td>
                            <td className="py-4 px-6 text-center">
                              {student.sakit > 0 ? (
                                <span className="bg-amber-50 text-amber-600 px-2 py-1 rounded-md text-xs font-extrabold border border-amber-100">
                                  {student.sakit} kali
                                </span>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                            <td className="py-4 px-6 text-center">
                              {student.izin > 0 ? (
                                <span className="bg-indigo-50 text-indigo-600 px-2 py-1 rounded-md text-xs font-extrabold border border-indigo-100">
                                  {student.izin} kali
                                </span>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                            <td className="py-4 px-6 text-center">
                              {student.alpa > 0 ? (
                                <span className="bg-rose-50 text-rose-600 px-2 py-1 rounded-md text-xs font-extrabold border border-rose-100 animate-pulse">
                                  {student.alpa} kali
                                </span>
                              ) : (
                                <span className="text-slate-300">-</span>
                              )}
                            </td>
                            <td className="py-4 px-6 text-center">
                              <span className="bg-slate-100 text-slate-700 font-black px-2.5 py-1.5 rounded-full text-xs border-2 border-slate-300">
                                {student.total} Hari
                              </span>
                            </td>
                            <td className="py-4 px-6 text-right">
                              <button 
                                onClick={() => {
                                  setActiveTab('reports');
                                }}
                                className="text-xs text-[#077a37] hover:text-lime-700 font-extrabold flex items-center justify-end gap-1 ml-auto"
                              >
                                Lihat Rekap <ArrowRight className="w-3.5 h-3.5" />
                              </button>
                            </td>
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>
              )}
              </div>
            </div>

            {/* MODAL / KOTAK DETAIL ABSENSI PER KELAS */}
            <AnimatePresence>
              {dashboardSelectedClassDetail && (
                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4 overflow-y-auto">
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.95, y: 20 }}
                    animate={{ opacity: 1, scale: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.95, y: 20 }}
                    className="bg-white rounded-[2.5rem] w-full max-w-2xl overflow-hidden shadow-2xl border border-slate-100 flex flex-col my-8"
                  >
                    <div className="bg-gradient-to-r from-lime-600 to-lime-700 p-6 sm:p-8 text-white flex justify-between items-center shrink-0">
                      <div>
                        <span className="bg-white/20 text-white px-3 py-1 rounded-full text-xs font-extrabold border border-white/10 uppercase tracking-widest">
                          Detail Ketidakhadiran
                        </span>
                        <h3 className="text-2xl font-black mt-2">Kelas {dashboardSelectedClassDetail}</h3>
                      </div>
                      <button 
                        onClick={() => setDashboardSelectedClassDetail(null)}
                        className="p-3 bg-white/10 rounded-2xl hover:bg-white/20 transition-all border border-white/10"
                      >
                        <X className="w-6 h-6" />
                      </button>
                    </div>
                    
                    <div className="p-6 sm:p-10 overflow-y-auto max-h-[85vh] md:max-h-[80vh] space-y-6 custom-scrollbar flex-1 bg-slate-50/20">
                      {(() => {
                        const classInfo = studentAbsenceStats.find(c => c.className === dashboardSelectedClassDetail);
                        if (!classInfo || classInfo.allAbsenceList.length === 0) {
                          return (
                            <div className="py-24 text-center text-slate-500 bg-white rounded-[2.5rem] border border-dashed border-slate-200 shadow-sm mx-2">
                              <div className="mx-auto w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mb-4">
                                <CheckCircle className="w-8 h-8 text-[#098f41]" />
                              </div>
                              <p className="font-bold text-slate-600">Semua siswa di kelas ini hadir 100%.</p>
                              <p className="text-xs mt-2 uppercase tracking-widest opacity-60">Tidak ada riwayat ketidakhadiran</p>
                            </div>
                          );
                        }

                        return (
                          <div className="space-y-5">
                            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">
                                Daftar Siswa
                              </p>
                              <span className="text-[10px] bg-slate-100 px-2 py-1 rounded-md text-slate-600 font-black">
                                {classInfo.allAbsenceList.length} SISWA TERCATAT
                              </span>
                            </div>
                            <div className="divide-y divide-slate-100">
                              {classInfo.allAbsenceList.map((stu) => (
                                <div key={stu.id} className="py-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4 first:pt-0 last:pb-0 hover:bg-slate-50/30 transition-colors px-2 rounded-xl">
                                  <div className="flex-1">
                                    <h4 className="font-black text-slate-800 text-lg tracking-tight">{stu.name}</h4>
                                    <div className="flex items-center gap-2 mt-1.5">
                                      <span className="w-1.5 h-1.5 rounded-full bg-slate-300"></span>
                                      <p className="text-xs text-slate-600 font-bold">Total: {stu.total} hari tidak hadir</p>
                                    </div>
                                  </div>
                                  <div className="flex flex-wrap items-center gap-2 sm:shrink-0">
                                    {stu.sakit > 0 && (
                                      <div className="flex flex-col items-start bg-amber-50 px-3.5 py-2 rounded-2xl border border-amber-100 shadow-sm shadow-amber-100/50">
                                        <span className="text-amber-600 text-xs font-black flex items-center gap-1.5">
                                          <span className="w-2 h-2 bg-amber-500 rounded-full animate-pulse"></span>
                                          {stu.sakit} Sakit
                                        </span>
                                        <span className="text-[9px] text-amber-600/70 mt-1 font-semibold leading-tight max-w-[120px]">
                                          {stu.sakitDates.map(d => new Date(d).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric'})).join(', ')}
                                        </span>
                                      </div>
                                    )}
                                    {stu.izin > 0 && (
                                      <div className="flex flex-col items-start bg-indigo-50 px-3.5 py-2 rounded-2xl border border-indigo-100 shadow-sm shadow-indigo-100/50">
                                        <span className="text-indigo-600 text-xs font-black flex items-center gap-1.5">
                                          <span className="w-2 h-2 bg-indigo-500 rounded-full animate-pulse"></span>
                                          {stu.izin} Izin
                                        </span>
                                        <span className="text-[9px] text-indigo-600/70 mt-1 font-semibold leading-tight max-w-[120px]">
                                          {stu.izinDates.map(d => new Date(d).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric'})).join(', ')}
                                        </span>
                                      </div>
                                    )}
                                    {stu.alpa > 0 && (
                                      <div className="flex flex-col items-start bg-rose-50 px-3.5 py-2 rounded-2xl border border-rose-100 shadow-sm shadow-rose-100/50">
                                        <span className="text-rose-600 text-xs font-black flex items-center gap-1.5">
                                          <span className="w-2 h-2 bg-rose-500 rounded-full animate-pulse"></span>
                                          {stu.alpa} Alpa
                                        </span>
                                        <span className="text-[9px] text-rose-600/70 mt-1 font-semibold leading-tight max-w-[120px]">
                                          {stu.alpaDates.map(d => new Date(d).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric'})).join(', ')}
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end gap-3 shrink-0">
                      <button 
                        onClick={() => setDashboardSelectedClassDetail(null)}
                        className="bg-white border-2 border-slate-300 text-slate-700 px-6 py-3 rounded-2xl font-bold text-sm hover:bg-slate-100 transition-all"
                      >
                        Tutup
                      </button>
                      <button 
                        onClick={() => {
                          setDashboardSelectedClassDetail(null);
                          setActiveTab('reports');
                        }}
                        className="bg-[#077a37] text-white px-6 py-3 rounded-2xl font-bold text-sm hover:bg-lime-700 transition-all shadow-sm flex items-center gap-2"
                      >
                        Buka Laporan <ArrowRight className="w-4 h-4" />
                      </button>
                    </div>
                  </motion.div>
                </div>
              )}
            </AnimatePresence>

            {/* Bottom Section - Quick Actions / Shortcuts */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
               {/* Quick Reports Access */}
               <div className="bg-white p-6 sm:p-8 rounded-[2rem] border-2 border-[#098f41]/30 hover:border-[#098f41] shadow-sm transition-all">
                  <div className="flex items-center gap-4 mb-6">
                    <div className="p-3 bg-emerald-50 rounded-2xl text-[#098f41]">
                       <BarChart3 className="w-6 h-6" />
                    </div>
                    <h2 className="text-xl font-bold text-slate-800 tracking-tight">Laporan & Rekap</h2>
                  </div>
                  <div className="space-y-3">
                     <button onClick={() => setActiveTab('reports')} className="w-full p-4 sm:p-5 rounded-2xl border border-slate-100 hover:border-[#098f41]/40 hover:bg-emerald-50/50 transition-colors flex items-center justify-between group">
                        <div className="flex items-center gap-4">
                           <FileText className="w-6 h-6 text-slate-500 group-hover:text-[#098f41] transition-colors" />
                           <span className="font-bold text-slate-600 group-hover:text-[#077a37] transition-colors">Laporan Presensi Harian</span>
                        </div>
                        <ArrowRight className="w-5 h-5 text-slate-300 group-hover:text-[#098f41] group-hover:translate-x-1 transition-all" />
                     </button>
                     <button onClick={() => setActiveTab('reports')} className="w-full p-4 sm:p-5 rounded-2xl border border-slate-100 hover:border-[#098f41]/40 hover:bg-emerald-50/50 transition-colors flex items-center justify-between group">
                        <div className="flex items-center gap-4">
                           <FileText className="w-6 h-6 text-slate-500 group-hover:text-[#098f41] transition-colors" />
                           <span className="font-bold text-slate-600 group-hover:text-[#077a37] transition-colors">Rekapitulasi Bulanan</span>
                        </div>
                        <ArrowRight className="w-5 h-5 text-slate-300 group-hover:text-[#098f41] group-hover:translate-x-1 transition-all" />
                     </button>
                     <button 
                       onClick={handleOpenHomeroomReport} 
                       className="w-full p-4 sm:p-5 rounded-2xl border border-amber-100 hover:border-amber-300 hover:bg-amber-50 transition-colors flex items-center justify-between group cursor-pointer"
                     >
                        <div className="flex items-center gap-4">
                           <ClipboardList className="w-6 h-6 transition-colors text-amber-500 group-hover:text-amber-600" />
                           <span className="font-bold transition-colors text-slate-600 group-hover:text-amber-800">Laporan Bulanan Wali Kelas</span>
                        </div>
                        <ArrowRight className="w-5 h-5 transition-all text-slate-300 group-hover:text-amber-600 group-hover:translate-x-1" />
                     </button>
                  </div>
               </div>

               {/* Hint / Setup Call to Action */}
               <div className="bg-white p-6 sm:p-8 rounded-[2rem] border-2 border-[#098f41]/30 hover:border-[#098f41] shadow-sm flex flex-col items-center justify-center text-center relative overflow-hidden transition-all">
                  <div className="absolute top-0 right-0 p-6 opacity-5">
                    <Settings2 className="w-48 h-48 -rotate-45" />
                  </div>
                  <div className="w-20 h-20 bg-emerald-50 rounded-[1.5rem] flex items-center justify-center text-[#098f41] mb-6 relative z-10 border border-emerald-100">
                     <Plus className="w-10 h-10" />
                  </div>
                  <h3 className="text-xl font-bold text-slate-800 relative z-10 mb-2">Perbarui Data Master</h3>
                  <p className="text-sm font-medium text-slate-600 max-w-xs relative z-10">Data siswa atau kelas ada yang baru? Segera tambahkan untuk keakuratan presensi.</p>
                  <button onClick={() => setActiveTab('students')} className="mt-8 bg-[#098f41] text-white px-6 py-3 rounded-2xl font-bold text-sm shadow-sm hover:bg-[#077a37] hover:shadow-md transition-all flex items-center gap-2 relative z-10">
                     Buka Manajemen Siswa <ArrowRight className="w-4 h-4 ml-1" />
                  </button>
               </div>
            </div>

          </div>
        );
      case 'students':
        return (
          <div className="p-8 space-y-8">
            {/* Kelas */}
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
              <h2 className="text-lg font-bold text-slate-800">Manajemen Kelas</h2>
                <div className="flex flex-col sm:flex-row gap-4">
                  <input type="text" className="flex-1 p-3 border rounded-xl" placeholder="Nama Kelas (Contoh: X-A)" onKeyDown={(e) => {
                    if (e.key === 'Enter') { 
                      const val = e.currentTarget.value.trim(); 
                      if (val) {
                        if (classList.some(c => c.toLowerCase() === val.toLowerCase())) {
                          showToast('Kelas "' + val + '" sudah terdaftar', 'error');
                        } else {
                          const arr = [...classList, val];
                          arr.sort((a,b) => a.localeCompare(b, 'id-ID', { numeric: true }));
                          setClassList(arr); 
                          
                          // Explicit cloud save
                          if (currentUser) {
                            setDoc(doc(activeDb, 'users', currentUser.uid), { classList: arr }, { merge: true })
                              .catch(e => console.error("Error saving new class:", e));
                          }
                          
                          e.currentTarget.value = ''; 
                        }
                      }
                    }
                  }} />
                  <button className="bg-[#098f41] text-white font-bold py-3 px-6 rounded-xl hover:bg-[#077a37] w-full sm:w-auto" onClick={(e) => {
                    const input = e.currentTarget.previousElementSibling as HTMLInputElement;
                    const val = input.value.trim(); 
                    if (val) {
                      if (classList.some(c => c.toLowerCase() === val.toLowerCase())) {
                          showToast('Kelas "' + val + '" sudah terdaftar', 'error');
                      } else {
                        const newList = [...classList, val];
                        setClassList(newList); 
                        
                        // Explicit cloud save
                        if (currentUser) {
                          setDoc(doc(activeDb, 'users', currentUser.uid), { classList: newList }, { merge: true })
                            .catch(e => console.error("Error saving new class:", e));
                        }
                        
                        input.value = ''; 
                      }
                    }
                  }}>Tambah</button>
                </div>
                <div className="flex flex-wrap gap-3">
                  {classList.length === 0 ? <p className="text-slate-600 italic">Belum ada kelas.</p> : classList.slice().sort((a,b) => compareClass(a, b)).map(c => {
                    const count = students.filter(s => s.class === c).length;
                    return (
                      <div key={c} className="flex flex-col gap-1 border border-slate-200 bg-slate-50 rounded-2xl p-3">
                        <div className="flex items-center gap-2 justify-between">
                          <span className="font-bold text-slate-700">{c}</span>
                          <span className="bg-white px-2 py-0.5 rounded-full text-xs font-bold text-slate-500 border shadow-sm">{count} Siswa</span>
                        </div>
                        <div className="flex gap-2 mt-2">
                          <button 
                            className="flex-1 text-xs py-1.5 px-3 rounded-lg bg-blue-50 text-blue-600 font-bold hover:bg-blue-100 flex justify-center items-center gap-1 transition-colors"
                            onClick={() => {
                              setClassToEdit(c);
                              setNewClassNameInput(c);
                            }}
                          >
                            <Pencil className="w-3.5 h-3.5" /> Edit
                          </button>
                          <button 
                            className="flex-1 text-xs py-1.5 px-3 rounded-lg bg-red-50 text-red-600 font-bold hover:bg-red-100 flex justify-center items-center gap-1 transition-colors"
                            onClick={() => {
                              setClassToDelete({ name: c, studentCount: count });
                            }}
                          >
                            <Trash2 className="w-3.5 h-3.5" /> Hapus
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
              
              <div className={`grid grid-cols-1 ${(profileData?.role === 'Guru Mapel' || profileData?.role === 'Wali Kelas') ? '' : 'lg:grid-cols-2'} gap-8 mb-8`}>
                {/* Siswa */}
                <div id="student-input-container" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <h2 className="text-lg font-bold text-slate-800">{editingStudentId ? 'Edit Data Siswa' : 'Tambah Siswa'}</h2>
                    <div className="flex flex-wrap items-center gap-2">
                      <input 
                        type="file" 
                        accept=".xlsx, .xls" 
                        className="hidden" 
                        ref={excelInputRef} 
                        onChange={handleImportExcel} 
                      />
                      <button 
                        onClick={() => excelInputRef.current?.click()}
                        className="bg-[#098f41]/10 text-[#098f41] border border-[#098f41]/20 font-bold py-2 px-3 rounded-xl hover:bg-[#098f41]/20 transition-colors text-xs sm:text-sm flex items-center gap-1.5"
                      >
                         <Download className="w-4 h-4" /> Import Excel
                      </button>
                      <button 
                        onClick={() => {
                          setGdriveTarget('student');
                          setIsGdriveModalOpen(true);
                        }}
                        className="bg-blue-50 text-blue-600 border border-blue-200 font-bold py-2 px-3 rounded-xl hover:bg-blue-100 transition-colors text-xs sm:text-sm flex items-center gap-1.5"
                        title="Ambil data Excel Siswa dari Google Drive / Google Sheets"
                      >
                         <span>🔗</span> Google Drive
                      </button>
                    </div>
                  </div>
                  <div className="grid grid-cols-1 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-600 mb-1">Nama Lengkap Siswa</label>
                      <input 
                        type="text" 
                        className="w-full p-3 border rounded-xl bg-white text-slate-800 focus:ring-2 focus:ring-[#098f41]" 
                        placeholder="Nama Lengkap Siswa" 
                        value={newStudent.name} 
                        onChange={(e) => setNewStudent({...newStudent, name: e.target.value})} 
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-600 mb-1">NIS (Nomor Induk Siswa)</label>
                      <input 
                        type="text" 
                        className="w-full p-3 border rounded-xl bg-white text-slate-800 focus:ring-2 focus:ring-[#098f41]" 
                        placeholder="NIS" 
                        value={newStudent.nisn} 
                        onChange={(e) => setNewStudent({...newStudent, nisn: e.target.value.replace(/\D/g, '')})} 
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-600 mb-1">Kelas</label>
                      <select 
                        className="w-full p-3 bg-white border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-[#098f41]" 
                        value={newStudent.class} 
                        onChange={(e) => {
                          const selectedC = e.target.value;
                          const existingWali = getWaliKelasForClass(selectedC, classWaliMap, students);
                          const existingWaliNiy = getWaliKelasNiyForClass(selectedC, classWaliNiyMap, students, teachers, existingWali);
                          setNewStudent(prev => ({
                            ...prev,
                            class: selectedC,
                            waliKelas: prev.waliKelas ? prev.waliKelas : (existingWali || ''),
                            waliKelasNiy: prev.waliKelasNiy ? prev.waliKelasNiy : (existingWaliNiy || '')
                          }));
                        }}
                      >
                        <option value="">Pilih Kelas</option>
                        {classList.slice().sort((a,b) => compareClass(a, b)).map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    {profileData?.role !== 'Guru Mapel' && profileData?.role !== 'Wali Kelas' && (
                      <>
                        <div>
                          <label className="block text-xs font-bold text-slate-600 mb-1 flex items-center justify-between">
                            <span>Nama Wali Kelas</span>
                            {profileData?.role === 'Petugas Piket' && (
                              <span className="text-[10px] font-bold text-[#098f41] bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                                Petugas Piket
                              </span>
                            )}
                          </label>
                          <input 
                            type="text" 
                            className="w-full p-3 border rounded-xl bg-white text-slate-800 focus:ring-2 focus:ring-[#098f41]" 
                            placeholder="Nama Wali Kelas (Contoh: Bpk. Ahmad, S.Pd)" 
                            value={newStudent.waliKelas || ''} 
                            onChange={(e) => {
                              const newName = e.target.value;
                              const matchedTeacher = teachers.find(t => t.name.trim().toLowerCase() === newName.trim().toLowerCase());
                              setNewStudent(prev => ({
                                ...prev,
                                waliKelas: newName,
                                waliKelasNiy: (matchedTeacher && matchedTeacher.niy && matchedTeacher.niy !== '-') ? matchedTeacher.niy : prev.waliKelasNiy
                              }));
                            }} 
                            list="teacher-wali-suggestions"
                          />
                          <datalist id="teacher-wali-suggestions">
                            {teachers.map(t => (
                              <option key={t.id} value={t.name}>{t.name} (Guru {t.niy && t.niy !== '-' ? `- NIY: ${t.niy}` : ''})</option>
                            ))}
                          </datalist>
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-600 mb-1 flex items-center justify-between">
                            <span>NIY atau NUPTK Wali Kelas</span>
                            <span className="text-[10px] text-slate-400 font-normal">Opsional</span>
                          </label>
                          <input 
                            type="text" 
                            className="w-full p-3 border rounded-xl bg-white text-slate-800 focus:ring-2 focus:ring-[#098f41]" 
                            placeholder="NIY atau NUPTK Wali Kelas (Contoh: 19850101... / 12345)" 
                            value={newStudent.waliKelasNiy || ''} 
                            onChange={(e) => setNewStudent({...newStudent, waliKelasNiy: e.target.value})} 
                          />
                        </div>
                      </>
                    )}
                  </div>
                  <button className="bg-[#098f41] text-white font-bold py-3 px-6 rounded-xl hover:bg-[#077a37] w-full cursor-pointer" onClick={addOrUpdateStudent}>{editingStudentId ? 'Update' : 'Simpan'}</button>
                </div>

                {/* Guru */}
                {profileData?.role !== 'Guru Mapel' && profileData?.role !== 'Wali Kelas' && (
                  <div id="teacher-input-container" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
                     <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                       <h2 className="text-lg font-bold text-slate-800">Tambah Guru</h2>
                       <div className="flex flex-wrap items-center gap-2">
                         <input 
                           type="file" 
                           accept=".xlsx, .xls" 
                           className="hidden" 
                           ref={excelTeacherInputRef} 
                           onChange={handleImportTeacherExcel} 
                         />
                         <button 
                           onClick={() => excelTeacherInputRef.current?.click()}
                           className="bg-[#098f41]/10 text-[#098f41] border border-[#098f41]/20 font-bold py-2 px-3 rounded-xl hover:bg-[#098f41]/20 transition-colors text-xs sm:text-sm flex items-center gap-1.5"
                         >
                            <Download className="w-4 h-4" /> Import Excel
                         </button>
                         <button 
                           onClick={() => {
                             setGdriveTarget('teacher');
                             setIsGdriveModalOpen(true);
                           }}
                           className="bg-blue-50 text-blue-600 border border-blue-200 font-bold py-2 px-3 rounded-xl hover:bg-blue-100 transition-colors text-xs sm:text-sm flex items-center gap-1.5"
                           title="Ambil data Excel Guru dari Google Drive / Google Sheets"
                         >
                            <span>🔗</span> Google Drive
                         </button>
                       </div>
                     </div>
                     <div className="grid grid-cols-1 gap-4">
                       <input type="text" className="p-3 border rounded-xl" placeholder="Nama Lengkap Guru" value={newTeacher.name} onChange={(e) => setNewTeacher({...newTeacher, name: e.target.value})} />
                       <input type="text" className="p-3 border rounded-xl" placeholder="Nomor Induk Yayasan (NIY) (isi '-' jika tidak ada)" value={newTeacher.niy} onChange={(e) => setNewTeacher({...newTeacher, niy: e.target.value.replace(/[^0-9-]/g, '')})} />
                     </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 uppercase mb-1">Jadwal Hari Piket</label>
                        <select 
                          className="w-full p-3 border rounded-xl bg-white font-medium text-slate-800"
                          value={newTeacher.dutyDay || ''}
                          onChange={(e) => setNewTeacher({...newTeacher, dutyDay: e.target.value})}
                        >
                          <option value="">-- Pilih Hari Piket --</option>
                          <option value="Senin">Senin</option>
                          <option value="Selasa">Selasa</option>
                          <option value="Rabu">Rabu</option>
                          <option value="Kamis">Kamis</option>
                          <option value="Jumat">Jumat</option>
                          <option value="Sabtu">Sabtu</option>
                          <option value="Minggu">Minggu</option>
                        </select>
                      </div>
                     <button className="bg-[#098f41] text-white font-bold py-3 px-6 rounded-xl hover:bg-[#077a37] w-full cursor-pointer" onClick={addOrUpdateTeacher}>{editingTeacherId ? 'Update' : 'Simpan'}</button>
                  </div>
                )}
              </div>

            <div className={`grid grid-cols-1 ${(profileData?.role === 'Guru Mapel' || profileData?.role === 'Wali Kelas') ? '' : 'lg:grid-cols-2'} gap-8`}>
              {/* Daftar Siswa */}
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                <div className="flex flex-col justify-between gap-4 mb-6">
                  <div className="flex justify-between items-center w-full">
                    <h2 className="text-lg font-bold text-slate-800">Daftar Siswa</h2>
                    {students.length > 0 && (
                      <button 
                        onClick={() => setResetModalType('clear_all_students')} 
                        className="bg-rose-50 text-rose-600 border border-rose-200 font-bold py-1.5 px-3 rounded-xl hover:bg-rose-100 transition-colors text-[11px] flex items-center justify-center gap-1.5 shrink-0"
                      >
                        <Trash2 className="w-3.5 h-3.5" /> Kosongkan Semua
                      </button>
                    )}
                  </div>
                </div>
                {!studentsLoaded ? (
                  <p className="text-slate-600 text-center py-6 text-sm">Memuat data...</p>
                ) : students.length === 0 ? (
                  <p className="text-slate-600 text-center py-6 text-sm">Belum ada siswa.</p>
                ) : (
                  <div className="overflow-auto max-h-[600px] border border-slate-100 rounded-xl scrollbar-thin">
                    <table className="w-full text-left border-collapse text-sm">
                      <thead className="sticky top-0 bg-slate-50 z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                        <tr>
                          <th className="p-3 font-bold text-slate-600 bg-slate-50">Nomor</th>
                          <th className="p-3 font-bold text-slate-600 bg-slate-50">Nama Lengkap Siswa</th>
                          <th className="p-3 font-bold text-slate-600 bg-slate-50">Nis</th>
                          <th className="p-3 font-bold text-slate-600 bg-slate-50">Kelas</th>
                          {profileData?.role !== 'Guru Mapel' && profileData?.role !== 'Wali Kelas' && (
                            <>
                              <th className="p-3 font-bold text-slate-600 bg-slate-50">Wali Kelas</th>
                              <th className="p-3 font-bold text-slate-600 bg-slate-50">NIY Wali Kelas</th>
                            </>
                          )}
                          <th className="p-3 font-bold text-slate-600 bg-slate-50">Aksi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sortStudents(students).map((student, i) => {
                          const studentWali = student.waliKelas || getWaliKelasForClass(student.class, classWaliMap, students);
                          const studentWaliNiy = student.waliKelasNiy || getWaliKelasNiyForClass(student.class, classWaliNiyMap, students, teachers, studentWali);
                          return (
                            <tr key={student.id} className={`border-b last:border-b-0 hover:bg-slate-50/80 transition-colors ${deletingStudentId === student.id ? 'opacity-50 bg-rose-50' : ''}`}>
                              <td className="p-3 text-slate-700">{i + 1}</td>
                              <td className="p-3 font-bold text-slate-900">
                                {student.name}
                                {deletingStudentId === student.id && <span className="ml-2 text-[9px] text-rose-500 font-bold animate-pulse">HAPUS...</span>}
                              </td>
                              <td className="p-3 text-slate-700">{student.nisn}</td>
                              <td className="p-3 text-slate-700">{student.class}</td>
                              {profileData?.role !== 'Guru Mapel' && profileData?.role !== 'Wali Kelas' && (
                                <>
                                  <td className="p-3 text-slate-700 text-xs">
                                    {studentWali ? (
                                      <span className="inline-flex items-center px-2 py-0.5 rounded-md text-xs font-semibold bg-emerald-50 text-emerald-800 border border-emerald-200/60 w-fit">
                                        {studentWali}
                                      </span>
                                    ) : (
                                      <span className="text-slate-400 italic">-</span>
                                    )}
                                  </td>
                                  <td className="p-3 text-slate-700 text-xs font-mono">
                                    {studentWaliNiy && studentWaliNiy !== '-' ? (
                                      <span className="inline-flex items-center px-2 py-0.5 rounded-md text-xs font-medium bg-slate-100 text-slate-700 border border-slate-200/60">
                                        {studentWaliNiy}
                                      </span>
                                    ) : (
                                      <span className="text-slate-400 italic">-</span>
                                    )}
                                  </td>
                                </>
                              )}
                              <td className="p-3 flex gap-1.5">
                                <button 
                                  onClick={() => requestEditStudent(student)} 
                                  disabled={!!deletingStudentId}
                                  title="Edit"
                                  className="p-1.5 border border-slate-100 rounded-lg hover:border-[#098f41] text-[#098f41] hover:text-[#077a37] hover:bg-emerald-50/50 transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
                                >
                                  <Pencil className="w-3.5 h-3.5" />
                                </button>
                                <button 
                                  onClick={() => requestDeleteStudent(student)} 
                                  disabled={!!deletingStudentId}
                                  title="Hapus"
                                  className="p-1.5 border border-slate-100 rounded-lg hover:border-rose-200 text-rose-600 hover:text-rose-800 hover:bg-rose-50/50 transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              {/* Daftar Guru */}
              {profileData?.role !== 'Guru Mapel' && profileData?.role !== 'Wali Kelas' && (
                <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                  <div className="flex flex-col justify-between gap-4 mb-6">
                    <div className="flex justify-between items-center w-full">
                      <h2 className="text-lg font-bold text-slate-800">Daftar Guru</h2>
                      {teachers.length > 0 && (
                        <button 
                          onClick={() => setResetModalType('clear_all_teachers')} 
                          className="bg-rose-50 text-rose-600 border border-rose-200 font-bold py-1.5 px-3 rounded-xl hover:bg-rose-100 transition-colors text-[11px] flex items-center justify-center gap-1.5 shrink-0"
                        >
                          <Trash2 className="w-3.5 h-3.5" /> Kosongkan Semua
                        </button>
                      )}
                    </div>
                  </div>
                  {!teachersLoaded ? (
                    <p className="text-slate-600 text-center py-6 text-sm">Memuat data...</p>
                  ) : teachers.length === 0 ? (
                    <p className="text-slate-600 text-center py-6 text-sm">Belum ada guru.</p>
                  ) : (
                    <div className="overflow-auto max-h-[600px] border border-slate-100 rounded-xl scrollbar-thin">
                      <table className="w-full text-left border-collapse text-sm">
                        <thead className="sticky top-0 bg-slate-50 z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                          <tr>
                            <th className="p-3 font-bold text-slate-600 bg-slate-50">No</th>
                            <th className="p-3 font-bold text-slate-600 bg-slate-50">Nama Guru</th>
                            <th className="p-3 font-bold text-slate-600 bg-slate-50">NIY</th>
                            <th className="p-3 font-bold text-slate-600 bg-slate-50">Jadwal Piket</th>
                            <th className="p-3 font-bold text-slate-600 bg-slate-50">Aksi</th>
                          </tr>
                        </thead>
                        <tbody>
                          {teachers.map((teacher, i) => (
                            <tr key={teacher.id} className={`border-b last:border-b-0 hover:bg-slate-50/80 transition-colors ${deletingTeacherId === teacher.id ? 'opacity-50 bg-rose-50' : ''}`}>
                              <td className="p-3 text-slate-700">{i + 1}</td>
                              <td className="p-3 font-bold text-slate-900">
                                {teacher.name}
                                {deletingTeacherId === teacher.id && <span className="ml-2 text-[9px] text-rose-500 font-bold animate-pulse">HAPUS...</span>}
                              </td>
                              <td className="p-3 text-slate-700">{teacher.niy}</td>
                              <td className="p-3 text-slate-700">
                                {teacher.dutyDay ? (
                                  <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">
                                    {teacher.dutyDay}
                                  </span>
                                ) : (
                                  <span className="text-slate-400 italic text-xs">-</span>
                                )}
                              </td>
                              <td className="p-3 flex gap-1.5">
                                <button 
                                  onClick={() => requestEditTeacher(teacher)} 
                                  disabled={!!deletingTeacherId}
                                  className="p-1.5 border border-slate-100 rounded-lg hover:border-blue-500 text-blue-500 hover:text-blue-700 hover:bg-blue-50/50 transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
                                >
                                  <Pencil className="w-3.5 h-3.5" />
                                </button>
                                <button 
                                  onClick={() => requestDeleteTeacher(teacher)} 
                                  disabled={!!deletingTeacherId}
                                  className="p-1.5 border border-slate-100 rounded-lg hover:border-rose-200 text-rose-600 hover:text-rose-800 hover:bg-rose-50/50 transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        );
      case 'attendance':
        return (
          <AttendanceView
            classList={classList}
            students={students}
            teachers={teachers}
            attendanceSessions={attendanceSessions}
            showToast={showToast}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
            profileData={profileData}
            classWaliMap={classWaliMap}
            currentUser={currentUser}
          />
        );
      case 'reports':
        return (
          <ReportsView 
            classList={classList}
            students={students}
            teachers={teachers}
            attendanceSessions={attendanceSessions}
            profileData={profileData}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
            showToast={showToast}
            classWaliMap={classWaliMap}
            classWaliNiyMap={classWaliNiyMap}
            onNavigateToProfile={() => setActiveTab('profile')}
            onSaveClassWali={handleSaveClassWali}
          />
        );
      case 'homeroom_report':
        return (
          <ReportsView 
            classList={classList}
            students={students}
            teachers={teachers}
            attendanceSessions={attendanceSessions}
            profileData={profileData}
            activeDb={activeDb}
            activeAuth={activeAuth}
            trackOp={trackOp}
            showToast={showToast}
            classWaliMap={classWaliMap}
            classWaliNiyMap={classWaliNiyMap}
            initialFrame="wali_kelas"
            onNavigateToProfile={() => setActiveTab('profile')}
            onSaveClassWali={handleSaveClassWali}
          />
        );
      case 'profile':
        return (
          <div className="p-4 sm:p-6 lg:p-8 space-y-6 flex flex-col items-center max-w-7xl mx-auto w-full">
            <h2 className="text-2xl sm:text-3xl font-black text-slate-800 tracking-tight self-start mb-2">Profil Saya</h2>
            
            <div className="bg-white w-full rounded-[2rem] p-6 sm:p-8 border border-slate-100 shadow-sm flex flex-col items-center">
              {/* Avatar Section with Interactive Click, Hover & Action Buttons */}
              <div className="flex flex-col items-center mb-6">
                <div className="relative group">
                  {/* Clickable Avatar Circle */}
                  <div 
                    onClick={() => !isUploadingPhoto && profilePhotoInputRef.current?.click()}
                    className={`w-32 h-32 sm:w-36 sm:h-36 rounded-full overflow-hidden bg-slate-100 border-4 border-white shadow-xl flex items-center justify-center relative cursor-pointer transition-all hover:ring-4 hover:ring-emerald-300/60 ${isUploadingPhoto ? 'opacity-70 pointer-events-none' : ''}`}
                    title="Klik untuk memilih/mengganti foto profil"
                  >
                    {avatarUrl ? (
                      <img src={avatarUrl} alt="Foto Profil" className="w-full h-full object-cover rounded-full" />
                    ) : (
                      <UserIcon className="w-14 h-14 text-slate-400" />
                    )}
                    
                    {/* Hover Overlay */}
                    {!isUploadingPhoto && (
                      <div className="absolute inset-0 bg-black/45 backdrop-blur-2xs opacity-0 group-hover:opacity-100 transition-opacity flex flex-col items-center justify-center text-white p-2 text-center">
                        <Camera className="w-6 h-6 mb-1 text-emerald-300" />
                        <span className="text-[10px] font-black uppercase tracking-wider text-emerald-200">{avatarUrl ? 'Ganti Foto' : 'Unggah Foto'}</span>
                      </div>
                    )}

                    {/* Uploading Spinner */}
                    {isUploadingPhoto && (
                      <div className="absolute inset-0 bg-black/75 backdrop-blur-xs flex flex-col items-center justify-center text-white">
                        <span className="w-6 h-6 border-2 border-emerald-400 border-t-transparent rounded-full animate-spin"></span>
                        <p className="text-[9px] font-black mt-1.5 text-emerald-300 uppercase tracking-widest">Menyimpan...</p>
                      </div>
                    )}
                  </div>

                  {/* Camera Badge Icon Button */}
                  <button 
                    type="button"
                    onClick={() => !isUploadingPhoto && profilePhotoInputRef.current?.click()}
                    disabled={isUploadingPhoto}
                    className={`absolute bottom-0 right-0 bg-[#098f41] text-white p-2.5 rounded-full shadow-lg border-2 border-white hover:bg-[#077a37] transition-all hover:scale-110 active:scale-95 cursor-pointer disabled:opacity-50 ${isUploadingPhoto ? 'opacity-50 pointer-events-none' : ''}`}
                    title="Pilih Foto Profil"
                  >
                    <Camera className="w-4.5 h-4.5" />
                  </button>

                  {/* Hidden Dedicated File Input */}
                  <input 
                    ref={profilePhotoInputRef}
                    type="file" 
                    className="hidden" 
                    accept="image/png, image/jpeg, image/jpg, image/webp" 
                    onChange={handleProfilePhotoUpload}
                  />
                </div>

                {/* Explicit Action Buttons */}
                <div className="flex items-center gap-2 mt-4">
                  <button
                    type="button"
                    onClick={() => !isUploadingPhoto && profilePhotoInputRef.current?.click()}
                    disabled={isUploadingPhoto}
                    className="px-3.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 text-[#098f41] rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 shadow-2xs hover:scale-102 active:scale-98 disabled:opacity-50 cursor-pointer"
                  >
                    <Camera className="w-3.5 h-3.5" />
                    <span>{avatarUrl ? 'Ganti Foto' : 'Unggah Foto'}</span>
                  </button>

                  {avatarUrl && (
                    <button
                      type="button"
                      onClick={handleRemoveProfilePhoto}
                      disabled={isUploadingPhoto}
                      className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-600 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 shadow-2xs hover:scale-102 active:scale-98 disabled:opacity-50 cursor-pointer"
                      title="Hapus Foto Profil"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Hapus</span>
                    </button>
                  )}
                </div>
                <p className="text-[11px] text-slate-400 mt-1.5 font-medium">Format: JPG, PNG, WEBP (Otomatis dipotong persegi)</p>
              </div>

              <div className="text-center w-full mb-8 flex flex-col items-center">
                <h3 className="text-xl font-black text-slate-800">{activeUserCustomData?.fullname || profileData.namaGuruMapel || 'Agan Parta,S.Kom.,Gr'}</h3>
                <p className="text-sm font-medium text-slate-600 my-1">{activeAuth.currentUser?.email || 'N/A'}</p>
                <div className="inline-flex items-center gap-1.5 bg-emerald-50 text-[#098f41] px-3 py-1 rounded-lg text-xs font-bold mt-2">
                  <span className="w-2 h-2 rounded-full bg-[#098f41]"></span> Online
                </div>

                {/* PWA Install Card inside Profile */}
                <div className="mt-6 w-full max-w-md bg-stone-50 border border-slate-100 rounded-2xl p-5 text-center shadow-xs">
                  <div className="flex items-center justify-center gap-2 mb-2 text-slate-800">
                    <Download className="w-4.5 h-4.5 text-[#098f41]" />
                    <span className="font-extrabold text-[10px] tracking-widest text-slate-700 uppercase">Aplikasi PWA (Absensi Seluler)</span>
                  </div>
                  {isInsideIframe ? (
                    <div className="space-y-2">
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Anda sedang membuka aplikasi ini di dalam frame preview. Untuk menginstalnya ke HP / Desktop agar dapat dibuka langsung tanpa browser:
                      </p>
                      <a
                        href="https://ais-pre-56w2g4yxpoxk4k23siyzdq-901834158843.asia-southeast1.run.app"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-2 px-4 py-2 bg-[#098f41] hover:bg-[#077a37] text-white text-xs font-bold rounded-xl transition-all shadow-[0_4px_12px_rgba(9,143,65,0.25)] hover:scale-102"
                      >
                        Buka di Tab Baru & Instal <ArrowRight className="w-3.5 h-3.5" />
                      </a>
                    </div>
                  ) : isAppInstalled ? (
                    <p className="text-xs text-emerald-600 font-bold">
                      🎉 Aplikasi telah terpasang dengan sukses! Anda dapat membukanya langsung dari layar utama perangkat Anda.
                    </p>
                  ) : isInstallable ? (
                    <div className="space-y-2.5">
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Aplikasi absensi ini siap dipasang langsung di HP atau Komputer Anda untuk akses cepat dan hemat kuota.
                      </p>
                      <button
                        onClick={handleInstallClick}
                        className="inline-flex items-center gap-2 px-4 py-2 bg-[#098f41] hover:bg-[#077a37] text-white text-xs font-bold rounded-xl transition-all shadow-[0_4px_12px_rgba(9,143,65,0.25)] hover:scale-102 active:scale-98"
                      >
                        Instal Sekarang
                      </button>
                    </div>
                  ) : (
                    <p className="text-[11px] text-slate-500 leading-relaxed">
                      Buka menu browser Anda (klik titik tiga di Chrome atau tombol Bagikan di Safari iOS) dan pilih <b>"Instal Aplikasi"</b> atau <b>"Tambahkan ke Layar Utama"</b>.
                    </p>
                  )}
                </div>
              </div>



              <div className="w-full h-px bg-slate-100 mb-6"></div>

              <div className="w-full mb-8 text-left">
                <div className="flex justify-between items-center mb-5">
                  <h3 className="text-lg font-bold text-slate-800">Detail Akademik</h3>
                  {!isProfileEditing ? (
                    <button 
                      onClick={() => setIsProfileEditing(true)}
                      className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition-colors"
                    >
                      Edit Data
                    </button>
                  ) : null}
                </div>

                 {/* Form Khusus Petugas Piket VS Pengguna Guru & Admin */}
                 {((activeUserCustomData?.username || '').toLowerCase().trim() === 'petugaspiket' || (currentUser?.email || '').toLowerCase().startsWith('petugaspiket@') || (currentUser?.uid || '').toLowerCase() === 'petugaspiket') ? (
                   /* Khusus Akun Petugas Piket */
                   <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-5">
                     <div className="space-y-1.5">
                       <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">
                         Koordinator Piket
                       </label>
                       <input 
                         type="text" 
                         disabled={!isProfileEditing} 
                         value={profileData.namaGuruMapel} 
                         onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaGuruMapel: e.target.value }))} 
                         placeholder="Nama Koordinator Piket"
                         className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                       />
                     </div>
                     
                     <div className="space-y-1.5">
                       <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">
                         NIY Koordinator Piket
                       </label>
                       <input 
                         type="text" 
                         disabled={!isProfileEditing} 
                         value={profileData.nipGuruMapel} 
                         onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipGuruMapel: e.target.value }))} 
                         placeholder="NIY Koordinator Piket"
                         className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                       />
                     </div>

                     <div className="space-y-1.5">
                       <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Semester</label>
                       <select 
                         disabled={!isProfileEditing} 
                         value={profileData.semester} 
                         onChange={e => setProfileData((p: typeof profileData) => ({ ...p, semester: e.target.value }))} 
                         className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100"
                       >
                         <option value="Ganjil">Ganjil</option>
                         <option value="Genap">Genap</option>
                       </select>
                     </div>

                     <div className="space-y-1.5">
                       <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Tahun Pelajaran</label>
                       <input 
                         type="text" 
                         disabled={!isProfileEditing} 
                         value={profileData.tahunPelajaran} 
                         onChange={e => setProfileData((p: typeof profileData) => ({ ...p, tahunPelajaran: e.target.value }))} 
                         placeholder="Contoh: 2026/2027" 
                         className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                       />
                     </div>
                   </div>
                 ) : (
                   /* Khusus User Admin dan Pengguna/Guru yang Lain (Selain Akun Petugas Piket) */
                   <div className="space-y-6">
                     <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-5">
                       {/* 1. Nama Guru */}
                       <div className="space-y-1.5">
                         <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider flex items-center gap-1.5">
                           <span>Nama Guru</span>
                           <span className="text-red-500 font-bold">*</span>
                         </label>
                         <input 
                           type="text" 
                           disabled={!isProfileEditing} 
                           value={profileData.namaGuruMapel} 
                           onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaGuruMapel: e.target.value }))} 
                           placeholder="Nama lengkap beserta gelar guru"
                           className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                         />
                       </div>
                       
                       {/* 2. NIY Guru */}
                       <div className="space-y-1.5">
                         <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider flex items-center gap-1.5">
                           <span>NIY Guru</span>
                           <span className="text-red-500 font-bold">*</span>
                         </label>
                         <input 
                           type="text" 
                           disabled={!isProfileEditing} 
                           value={profileData.nipGuruMapel} 
                           onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipGuruMapel: e.target.value }))} 
                           placeholder="Nomor Induk Yayasan (NIY)"
                           className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                         />
                       </div>

                       {/* 3. Peran Role (Dropdown: Guru Mata Pelajaran atau Wali Kelas) */}
                       <div className="space-y-1.5">
                         <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider flex items-center gap-1.5">
                           <span>Peran (Role)</span>
                           <span className="text-red-500 font-bold">*</span>
                         </label>
                         <select 
                           disabled={!isProfileEditing} 
                           value={profileData.role === 'Wali Kelas' ? 'Wali Kelas' : 'Guru Mapel'} 
                           onChange={e => {
                             const newRole = e.target.value;
                             setProfileData((p: typeof profileData) => ({ ...p, role: newRole }));
                           }} 
                           className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100"
                         >
                           <option value="Guru Mapel">Guru Mata Pelajaran</option>
                           <option value="Wali Kelas">Wali Kelas</option>
                         </select>
                       </div>

                       {/* 4. Mata Pelajaran (Jika Memilih sebagai Guru Mata Pelajaran) */}
                       {profileData.role !== 'Wali Kelas' && (
                         <div className="space-y-1.5">
                           <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider flex items-center gap-1.5">
                             <span>Mata Pelajaran</span>
                             <span className="text-red-500 font-bold">*</span>
                           </label>
                           <input 
                             type="text" 
                             disabled={!isProfileEditing} 
                             value={profileData.mataPelajaran} 
                             onChange={e => setProfileData((p: typeof profileData) => ({ ...p, mataPelajaran: e.target.value }))} 
                             placeholder="Contoh: Informatika / Bahasa Indonesia"
                             className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                           />
                         </div>
                       )}

                       {/* Semester */}
                       <div className="space-y-1.5">
                         <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Semester</label>
                         <select 
                           disabled={!isProfileEditing} 
                           value={profileData.semester} 
                           onChange={e => setProfileData((p: typeof profileData) => ({ ...p, semester: e.target.value }))} 
                           className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100"
                         >
                           <option value="Ganjil">Ganjil</option>
                           <option value="Genap">Genap</option>
                         </select>
                       </div>

                       {/* Tahun Pelajaran */}
                       <div className="space-y-1.5 sm:col-span-2">
                         <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Tahun Pelajaran</label>
                         <input 
                           type="text" 
                           disabled={!isProfileEditing} 
                           value={profileData.tahunPelajaran} 
                           onChange={e => setProfileData((p: typeof profileData) => ({ ...p, tahunPelajaran: e.target.value }))} 
                           placeholder="Contoh: 2026/2027" 
                           className="w-full px-4 py-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:text-slate-600 disabled:border-slate-100" 
                         />
                       </div>
                     </div>

                     {/* 5, 6, 7. Inputan Khusus Untuk Pengguna Wali Kelas */}
                     {profileData.role === 'Wali Kelas' && (
                       <div className="p-5 bg-gradient-to-br from-emerald-50/60 to-slate-50 border-2 border-emerald-200/80 rounded-2xl space-y-4">
                         <div className="flex items-center justify-between border-b border-emerald-100 pb-3">
                           <div className="flex items-center gap-2">
                             <span className="p-1.5 bg-emerald-500 text-white rounded-lg text-xs">🏫</span>
                             <div>
                               <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider">Pengaturan Data Wali Kelas</h4>
                               <p className="text-[11px] text-slate-500 font-medium">Isi data kelas perwalian dan jumlah siswa untuk Laporan Bulanan Wali Kelas</p>
                             </div>
                           </div>
                           {((Number(profileData.jumlahSiswaLakiLaki) || 0) + (Number(profileData.jumlahSiswaPerempuan) || 0) > 0) && (
                             <span className="px-3 py-1 bg-emerald-600 text-white rounded-full text-xs font-bold shadow-sm">
                               Total: {(Number(profileData.jumlahSiswaLakiLaki) || 0) + (Number(profileData.jumlahSiswaPerempuan) || 0)} Siswa
                             </span>
                           )}
                         </div>

                         <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                           {/* Kelas (Jika Memilih Wali Kelas) */}
                           <div className="space-y-1.5">
                             <label className="text-[10px] font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1">
                               <span>Kelas</span>
                               <span className="text-red-500 font-bold">*</span>
                             </label>
                             <select 
                               disabled={!isProfileEditing} 
                               value={profileData.waliKelasClass || ''} 
                               onChange={e => setProfileData((p: typeof profileData) => ({ ...p, waliKelasClass: e.target.value }))} 
                               className="w-full px-4 py-3 bg-white border-2 border-emerald-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-200 transition-all text-sm font-bold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:border-slate-200"
                             >
                               <option value="">-- Pilih Kelas --</option>
                               {classList.map(c => <option key={c} value={c}>{c}</option>)}
                             </select>
                           </div>

                           {/* Laki-laki (Jumlah Siswa Laki Laki) */}
                           <div className="space-y-1.5">
                             <label className="text-[10px] font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1">
                               <span>Laki-laki (L)</span>
                               <span className="text-slate-400 font-normal">(Jumlah Siswa)</span>
                             </label>
                             <input 
                               type="number" 
                               min="0" 
                               disabled={!isProfileEditing} 
                               value={profileData.jumlahSiswaLakiLaki || ''} 
                               onChange={e => setProfileData((p: typeof profileData) => ({ ...p, jumlahSiswaLakiLaki: e.target.value }))} 
                               placeholder="Contoh: 18" 
                               className="w-full px-4 py-3 bg-white border-2 border-emerald-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-200 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:border-slate-200" 
                             />
                           </div>

                           {/* Perempuan (Jumlah Siswa Perempuan) */}
                           <div className="space-y-1.5">
                             <label className="text-[10px] font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1">
                               <span>Perempuan (P)</span>
                               <span className="text-slate-400 font-normal">(Jumlah Siswa)</span>
                             </label>
                             <input 
                               type="number" 
                               min="0" 
                               disabled={!isProfileEditing} 
                               value={profileData.jumlahSiswaPerempuan || ''} 
                               onChange={e => setProfileData((p: typeof profileData) => ({ ...p, jumlahSiswaPerempuan: e.target.value }))} 
                               placeholder="Contoh: 18" 
                               className="w-full px-4 py-3 bg-white border-2 border-emerald-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-200 transition-all text-sm font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50 disabled:border-slate-200" 
                             />
                           </div>
                         </div>
                       </div>
                     )}
                   </div>
                 )}

                {/* Section Pejabat Penandatangan Laporan (Mengetahui) */}
                <div className="mt-8 pt-6 border-t-2 border-dashed border-slate-200">
                  <div className="mb-4">
                    <div className="flex items-center gap-2 mb-1">
                      <span className="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-bold uppercase tracking-wider">Opsi Tanda Tangan</span>
                      <h4 className="text-sm font-bold text-slate-800">Pejabat Penandatangan Laporan (Mengetahui - Kiri)</h4>
                    </div>
                    <p className="text-xs text-slate-500">
                      Lengkapi nama dan NIY pejabat di bawah ini. Anda dapat memilih salah satu pejabat saat mencetak Laporan Presensi atau Laporan Wali Kelas.
                    </p>
                  </div>

                  <div className="space-y-4">
                    {/* Kepala Sekolah */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                        1. Kepala Sekolah
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Kepala Sekolah</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaKepalaSekolah} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaKepalaSekolah: e.target.value }))} placeholder="Nama lengkap beserta gelar" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIY Kepala Sekolah</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipKepalaSekolah} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipKepalaSekolah: e.target.value }))} placeholder="Nomor Induk Yayasan (NIY)" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Pihak Kurikulum */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                        2. Pihak Kurikulum (Wakasek Kurikulum)
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Pihak Kurikulum</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaKurikulum || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaKurikulum: e.target.value }))} placeholder="Nama pejabat kurikulum" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIY Pihak Kurikulum</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipKurikulum || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipKurikulum: e.target.value }))} placeholder="NIY pejabat kurikulum" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Pihak Kesiswaan */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                        3. Pihak Kesiswaan (Wakasek Kesiswaan)
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Pihak Kesiswaan</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaKesiswaan || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaKesiswaan: e.target.value }))} placeholder="Nama pejabat kesiswaan" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIY Pihak Kesiswaan</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipKesiswaan || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipKesiswaan: e.target.value }))} placeholder="NIY pejabat kesiswaan" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Guru Wali */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-purple-500"></span>
                        4. Guru Wali / Koordinator Guru Wali
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Guru Wali</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaGuruWali || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaGuruWali: e.target.value }))} placeholder="Contoh: Agan Parta,S.Kom.,Gr." className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIY Guru Wali</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipGuruWali || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipGuruWali: e.target.value }))} placeholder="Contoh: 198203152009021003" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Guru BK */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-teal-500"></span>
                        5. Guru BK / Bimbingan Konseling
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Guru BK</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaBK || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaBK: e.target.value }))} placeholder="Nama guru BK" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIY Guru BK</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipBK || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipBK: e.target.value }))} placeholder="NIY guru BK" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>

                    {/* Wakasek Humas */}
                    <div className="p-4 bg-slate-50 border-2 border-slate-200/80 rounded-2xl">
                      <div className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-orange-500"></span>
                        6. Wakil Kepala Sekolah Bidang Humas (Hubungan Masyarakat)
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">Nama Wakasek Humas</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.namaHumas || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, namaHumas: e.target.value }))} placeholder="Nama pejabat humas" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider">NIY Wakasek Humas</label>
                          <input type="text" disabled={!isProfileEditing} value={profileData.nipHumas || ''} onChange={e => setProfileData((p: typeof profileData) => ({ ...p, nipHumas: e.target.value }))} placeholder="NIY pejabat humas" className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-xs font-semibold text-slate-800 disabled:opacity-60 disabled:bg-slate-50" />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="w-full bg-white border border-slate-100 rounded-[1.5rem] p-6 shadow-sm mb-8 text-left space-y-6">
                  {/* Header Statistik */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-5">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-blue-50 rounded-2xl text-blue-600 shadow-xs">
                        <Activity className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="text-base font-black text-slate-800 tracking-tight">Statistik Pemakaian Limit</h4>
                        <p className="text-xs text-slate-500 font-medium">Pantau kuota operasional Firestore Anda secara real-time.</p>
                      </div>
                    </div>

                    {/* Live Clock & WIB Badge */}
                    <div className="flex items-center gap-2 bg-slate-50 border border-slate-200/80 px-3.5 py-2 rounded-2xl shrink-0">
                      <Clock className="w-4 h-4 text-blue-600 animate-pulse" />
                      <div className="text-right">
                        <div className="text-[11px] font-black text-slate-800 font-mono tracking-wide">{cycleInfo.currentTimeStr}</div>
                        <div className="text-[9px] font-bold text-slate-500 uppercase tracking-wider">Waktu Indonesia Barat</div>
                      </div>
                    </div>
                  </div>

                  {/* Banner Hari, Tanggal Sekarang & Jadwal Reset 14.00 WIB */}
                  <div className="bg-gradient-to-r from-blue-50/80 via-emerald-50/60 to-slate-50 rounded-2xl p-4.5 border border-blue-100/70 text-left space-y-3">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                      {/* Hari & Tanggal Sekarang */}
                      <div className="flex items-start gap-3 bg-white/80 backdrop-blur-xs p-3 rounded-xl border border-blue-100/50 shadow-2xs">
                        <div className="p-2 bg-blue-100/70 text-blue-700 rounded-lg shrink-0 mt-0.5">
                          <Calendar className="w-4 h-4" />
                        </div>
                        <div>
                          <div className="text-[10px] font-black uppercase tracking-wider text-slate-500">Hari & Tanggal Sekarang</div>
                          <div className="text-xs sm:text-sm font-black text-slate-800 mt-0.5">{cycleInfo.currentDateFormatted}</div>
                        </div>
                      </div>

                      {/* Jadwal Reset Harian 14.00 WIB */}
                      <div className="flex items-start gap-3 bg-white/80 backdrop-blur-xs p-3 rounded-xl border border-emerald-100/50 shadow-2xs">
                        <div className="p-2 bg-emerald-100/70 text-[#077a37] rounded-lg shrink-0 mt-0.5">
                          <RefreshCw className="w-4 h-4" />
                        </div>
                        <div>
                          <div className="text-[10px] font-black uppercase tracking-wider text-slate-500">Jadwal Reset Limit (00:00 PST / 14:00 WIB)</div>
                          <div className="text-xs sm:text-sm font-black text-slate-800 mt-0.5">
                            Setiap Hari Pukul <span className="text-[#077a37]">14.00 WIB</span>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Countdown & Info Strip */}
                    <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 pt-1 px-1">
                      <div className="flex items-center gap-1.5 text-xs text-slate-600 font-medium">
                        <Timer className="w-4 h-4 text-blue-500 shrink-0" />
                        <span>Sisa waktu menuju reset pukul 14:00 WIB:</span>
                        <span className="font-mono font-black text-blue-700 bg-blue-100/60 px-2 py-0.5 rounded-md text-[11px]">{cycleInfo.countdownFormatted}</span>
                      </div>
                      <div className="inline-flex items-center gap-1 text-[10px] text-emerald-800 font-bold bg-emerald-100/80 px-2.5 py-1 rounded-full">
                        <CheckCircle className="w-3 h-3 text-[#098f41]" />
                        <span>Reset Otomatis ke 0 Aktif</span>
                      </div>
                    </div>
                  </div>
                  
                  {(() => {
                    const writeLimit = 20000;
                    const readLimit = 50000;
                    
                    const writeUsed = sessionUsage.writes || 0;
                    const writeRemaining = Math.max(0, writeLimit - writeUsed);
                    const writePercentage = Math.min(100, (writeUsed / writeLimit) * 100);
                    
                    const readUsed = sessionUsage.reads || 0;
                    const readRemaining = Math.max(0, readLimit - readUsed);
                    const readPercentage = Math.min(100, (readUsed / readLimit) * 100);
                    
                    return (
                      <div className="space-y-6 pt-2">
                        {/* Write Stats */}
                        <div className="space-y-3 bg-slate-50/50 p-4 rounded-2xl border border-slate-100">
                          <div className="flex items-center justify-between border-b border-slate-200/60 pb-2">
                            <h5 className="text-xs font-black text-slate-800">Batas Operasi Penulisan (Writes)</h5>
                            <span className="text-[10px] font-bold text-slate-500 bg-white px-2 py-0.5 rounded-md border border-slate-200">
                              Limit: {writeLimit.toLocaleString('id-ID')} / hari (Reset 14:00 WIB)
                            </span>
                          </div>
                          
                          <div className="flex justify-between items-end">
                            <div>
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Total Digunakan Sesi Ini</p>
                              <p className="text-xl font-black text-slate-800">{writeUsed.toLocaleString('id-ID')}</p>
                            </div>
                            <div className="text-right">
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Sisa Kuota</p>
                              <p className="text-base font-black text-blue-600">{writeRemaining.toLocaleString('id-ID')}</p>
                            </div>
                          </div>

                          <div className="w-full bg-slate-200/70 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className={`h-full rounded-full transition-all duration-500 ${writePercentage > 90 ? 'bg-rose-500' : writePercentage > 75 ? 'bg-amber-400' : 'bg-blue-500'}`}
                              style={{ width: `${writePercentage}%` }}
                            />
                          </div>
                          
                          <p className="text-[10px] text-slate-600 leading-relaxed bg-white p-2.5 rounded-xl border border-slate-200/60 mt-1">
                            Batas maksimal operasi penulisan perangkat Anda adalah <b>{writeLimit.toLocaleString('id-ID')}</b>. Kuota tersisa <b>{writeRemaining.toLocaleString('id-ID')}</b>. Setiap Anda menyimpan data, kuota penulisan berkurang dan akan direset ke 0 setiap pukul <b>14.00 WIB</b>.
                          </p>
                        </div>

                        {/* Read Stats */}
                        <div className="space-y-3 bg-slate-50/50 p-4 rounded-2xl border border-slate-100">
                          <div className="flex items-center justify-between border-b border-slate-200/60 pb-2">
                            <h5 className="text-xs font-black text-slate-800">Batas Operasi Pembacaan (Reads)</h5>
                            <span className="text-[10px] font-bold text-slate-500 bg-white px-2 py-0.5 rounded-md border border-slate-200">
                              Limit: {readLimit.toLocaleString('id-ID')} / hari (Reset 14:00 WIB)
                            </span>
                          </div>

                          <div className="flex justify-between items-end">
                            <div>
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Total Digunakan Sesi Ini</p>
                              <p className="text-xl font-black text-slate-800">{readUsed.toLocaleString('id-ID')}</p>
                            </div>
                            <div className="text-right">
                              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Sisa Kuota</p>
                              <p className="text-base font-black text-[#077a37]">{readRemaining.toLocaleString('id-ID')}</p>
                            </div>
                          </div>

                          <div className="w-full bg-slate-200/70 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className={`h-full rounded-full transition-all duration-500 ${readPercentage > 90 ? 'bg-rose-500' : readPercentage > 75 ? 'bg-amber-400' : 'bg-[#098f41]'}`}
                              style={{ width: `${readPercentage}%` }}
                            />
                          </div>
                          
                          <p className="text-[10px] text-slate-600 leading-relaxed bg-white p-2.5 rounded-xl border border-slate-200/60 mt-1">
                            Batas maksimal operasi pembacaan perangkat Anda adalah <b>{readLimit.toLocaleString('id-ID')}</b>. Kuota tersisa <b>{readRemaining.toLocaleString('id-ID')}</b>. Setiap Anda memuat aplikasi atau data, kuota pembacaan berkurang dan akan direset ke 0 setiap pukul <b>14.00 WIB</b>.
                          </p>
                        </div>
                      </div>
                    );
                  })()}
                </div>

                <div className="w-full mt-8">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4">
                    <div>
                      <h3 className="text-lg font-black text-slate-800">Rekapitulasi Penggunaan Limit Harian Semua Akun</h3>
                      <p className="text-xs text-slate-500 font-medium">Statistik pemakaian limit per user dalam siklus harian berjalan (Reset setiap 14.00 WIB).</p>
                    </div>
                    <span className="self-start sm:self-auto text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200/70 px-2.5 py-1 rounded-lg">
                      Siklus: 14:00 WIB
                    </span>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-100/80">
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider whitespace-nowrap">Nama Pengguna</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-right">Pembacaan (Reads)</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-right">Penulisan (Writes)</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-center whitespace-nowrap">Status Siklus</th>
                            <th className="px-4 py-3 text-[10px] font-bold text-slate-600 uppercase tracking-wider text-right whitespace-nowrap">Tanggal Catatan</th>
                          </tr>
                        </thead>
                        <tbody>
                          {isLoadingUsage ? (
                            <tr>
                              <td colSpan={5} className="px-4 py-8 text-center text-sm font-medium text-slate-500">Loading data statistik pengguna...</td>
                            </tr>
                          ) : allUsersUsage.length > 0 ? (
                            allUsersUsage.map((usage, idx) => (
                              <tr key={usage.username} className={`border-b border-slate-50 last:border-none ${idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'}`}>
                                <td className="px-4 py-3">
                                  <div className="text-sm font-bold text-slate-800">{usage.fullname}</div>
                                  <div className="text-[10px] font-medium text-slate-500">{usage.username}</div>
                                </td>
                                <td className="px-4 py-3 text-right">
                                  <span className={`inline-flex items-center px-2 py-0.5 rounded-md text-xs font-bold ${usage.reads > 10000 ? 'bg-rose-100 text-rose-700' : usage.reads > 5000 ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-[#077a37]'}`}>
                                    {usage.reads.toLocaleString('id-ID')}
                                  </span>
                                </td>
                                <td className="px-4 py-3 text-right">
                                  <span className={`inline-flex items-center px-2 py-0.5 rounded-md text-xs font-bold ${usage.writes > 5000 ? 'bg-rose-100 text-rose-700' : usage.writes > 2000 ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'}`}>
                                    {usage.writes.toLocaleString('id-ID')}
                                  </span>
                                </td>
                                <td className="px-4 py-3 text-center">
                                  {usage.isCurrentCycle ? (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200/60">
                                      <span className="w-1.5 h-1.5 rounded-full bg-[#098f41]"></span>
                                      Siklus Aktif
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200/60">
                                      Reset 14:00 WIB
                                    </span>
                                  )}
                                </td>
                                <td className="px-4 py-3 text-right">
                                  <div className="text-[10px] font-semibold text-slate-700">{usage.date}</div>
                                  {usage.time && <div className="text-[9px] font-medium text-slate-400">{usage.time}</div>}
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={5} className="px-4 py-8 text-center text-sm font-medium text-slate-500">Belum ada data penggunaan tercatat.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                {isProfileEditing && (
                  <div className="flex justify-end gap-3 mt-6">
                    <button 
                      onClick={() => {
                        setIsProfileEditing(false);
                      }}
                      disabled={isProfileSaving}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold justify-center rounded-xl transition-all"
                    >
                      Batal
                    </button>
                    <button 
                      onClick={async () => {
                        const uid = currentUser?.uid || activeAuth.currentUser?.uid;
                        setIsProfileSaving(true);
                        trackOp('write', 1);
                        try {
                          const savePayload = {
                            profileData, // Correct nested format
                            // Flat format for backward compatibility
                            namaGuruMapel: profileData.namaGuruMapel,
                            namaKepalaSekolah: profileData.namaKepalaSekolah,
                            nipGuruMapel: profileData.nipGuruMapel,
                            nipKepalaSekolah: profileData.nipKepalaSekolah,
                            namaKurikulum: profileData.namaKurikulum,
                            nipKurikulum: profileData.nipKurikulum,
                            namaKesiswaan: profileData.namaKesiswaan,
                            nipKesiswaan: profileData.nipKesiswaan,
                            namaGuruWali: profileData.namaGuruWali,
                            nipGuruWali: profileData.nipGuruWali,
                            namaBK: profileData.namaBK,
                            nipBK: profileData.nipBK,
                            namaHumas: profileData.namaHumas,
                            nipHumas: profileData.nipHumas,
                            semester: profileData.semester,
                            tahunPelajaran: profileData.tahunPelajaran,
                            mataPelajaran: profileData.mataPelajaran,
                            role: profileData.role,
                            hariPiket: profileData.hariPiket,
                            waliKelasClass: profileData.waliKelasClass,
                            jumlahSiswaLakiLaki: profileData.jumlahSiswaLakiLaki,
                            jumlahSiswaPerempuan: profileData.jumlahSiswaPerempuan
                          };
                          
                          // Save to activeDb (private database) but gracefully handle custom firestore permission rules
                          let savePromise = Promise.resolve();
                          if (uid) {
                            savePromise = setDoc(doc(activeDb, 'users', uid), savePayload, { merge: true })
                              .catch(err => {
                                const msg = err instanceof Error ? err.message : String(err);
                                if (msg.toLowerCase().includes('permission') || msg.toLowerCase().includes('insufficient')) {
                                  console.warn('Gagal menyimpan profil ke database mandiri (Rules ditolak). Mengandalkan pencadangan pusat FRESH gratis.');
                                  return; // Let the promise resolve to proceed seamlessly with the central backup
                                }
                                throw err;
                              });
                          }
                          
                          // Save to custom_accounts (central database) to guarantee permanent sync and recovery
                          let backupPromise = Promise.resolve();
                          if (activeUserCustomData?.username) {
                            trackOp('write', 1);
                            backupPromise = setDoc(doc(dbDefault, 'custom_accounts', activeUserCustomData.username.toLowerCase().trim()), {
                              profileData: savePayload.profileData,
                              // Also keep flat format on root level for backward compatibility
                              namaGuruMapel: savePayload.namaGuruMapel,
                              namaKepalaSekolah: savePayload.namaKepalaSekolah,
                              nipGuruMapel: savePayload.nipGuruMapel,
                              nipKepalaSekolah: savePayload.nipKepalaSekolah,
                              namaKurikulum: savePayload.namaKurikulum,
                              nipKurikulum: savePayload.nipKurikulum,
                              namaKesiswaan: savePayload.namaKesiswaan,
                              nipKesiswaan: savePayload.nipKesiswaan,
                              namaGuruWali: savePayload.namaGuruWali,
                              nipGuruWali: savePayload.nipGuruWali,
                              namaBK: savePayload.namaBK,
                              nipBK: savePayload.nipBK,
                              namaHumas: savePayload.namaHumas,
                              nipHumas: savePayload.nipHumas,
                              semester: savePayload.semester,
                              tahunPelajaran: savePayload.tahunPelajaran,
                              mataPelajaran: savePayload.mataPelajaran,
                              role: savePayload.role,
                              hariPiket: savePayload.hariPiket,
                              waliKelasClass: savePayload.waliKelasClass,
                              jumlahSiswaLakiLaki: savePayload.jumlahSiswaLakiLaki,
                              jumlahSiswaPerempuan: savePayload.jumlahSiswaPerempuan
                            }, { merge: true }).catch(err => {
                              console.warn('Gagal mencadangkan profil ke database pusat:', err);
                            });
                          }
                          
                          // Combine promises
                          const combinedSavePromise = Promise.all([savePromise, backupPromise]);
                          
                          // Race against a short timeout to guarantee instant performance (Firestore syncs in background anyway)
                          const timeoutPromise = new Promise<void>((_, reject) => 
                            setTimeout(() => reject(new Error('timeout')), 3000)
                          );
                          
                          try {
                            await Promise.race([combinedSavePromise, timeoutPromise]);
                          } catch (raceError) {
                            if (raceError instanceof Error && raceError.message === 'timeout') {
                              console.warn("Profile save server-sync timed out. Proceeding since local cache is updated.");
                            } else {
                              throw raceError;
                            }
                          }
                          
                          if (activeUserCustomData?.username) {
                            localStorage.setItem(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`, JSON.stringify(profileData));
                          }
                          
                          showToast('Profile Berhasil Disimpan', 'success');
                          setIsProfileEditing(false);
                        } catch (err) {
                          console.error('Error saving profile:', err);
                          handleFirestoreError(err, OperationType.WRITE, 'users');
                          showToast('Gagal menyimpan profil: ' + (err instanceof Error ? err.message : 'Server error'), 'error');
                        } finally {
                          setIsProfileSaving(false);
                        }
                      }}
                      disabled={isProfileSaving}
                      className="px-5 py-2.5 bg-[#098f41] hover:bg-[#077a37] text-white text-sm font-bold justify-center rounded-xl shadow-[0_4px_12px_rgba(5,150,105,0.3)] transition-all flex items-center gap-2 disabled:opacity-70 disabled:shadow-none"
                    >
                      {isProfileSaving ? (
                        <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                      ) : null}
                      Simpan Data
                    </button>
                  </div>
                )}
              </div>

              {/* KUMPULAN USER (ADMIN PANEL) - Custom Requested */}
              {isAdmin && (
                <>
                  <div className="w-full h-px bg-slate-100 mb-6"></div>
                  <div className="w-full mb-8 text-left bg-white border border-slate-150 rounded-[1.5rem] p-6 shadow-sm">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                      <div className="flex items-center gap-2.5">
                        <div className="p-2 bg-emerald-50 rounded-lg text-[#098f41]">
                          <Users className="w-5 h-5" />
                        </div>
                        <div>
                          <h3 className="text-base font-extrabold text-slate-800 tracking-tight">Kumpulan User (Sistem Administrator)</h3>
                          <p className="text-[10px] text-slate-500">Kelola dan lihat seluruh akun guru/pengajar terpusat.</p>
                        </div>
                      </div>
                      <button
                        onClick={fetchAllUsers}
                        className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition-all self-start sm:self-center"
                      >
                        Perbarui Daftar
                      </button>
                    </div>

                    <div className="bg-slate-50/50 border border-slate-100 rounded-2xl overflow-hidden shadow-xs">
                      {isLoadingUsers ? (
                        <div className="p-12 flex flex-col items-center justify-center gap-2.5 text-slate-500">
                          <span className="w-6 h-6 border-2 border-[#098f41] border-t-transparent rounded-full animate-spin"></span>
                          <span className="text-[10px] font-bold uppercase tracking-wider text-slate-550">Memuat User Terdaftar...</span>
                        </div>
                      ) : allUsers.length === 0 ? (
                        <div className="p-8 text-center text-xs text-slate-500 font-medium">
                          Tidak ada user terdaftar atau gagal memuat data.
                        </div>
                      ) : (
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-100/60 border-b border-slate-200/50">
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest text-center w-12">No</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest">Nama Pengguna</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest">Username</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest">Password</th>
                                <th className="px-4 py-3 text-[10px] font-extrabold text-slate-600 uppercase tracking-widest text-center w-28">Aksi (Edit / Hapus)</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {allUsers.map((u, idx) => {
                                const isSelf = u.username.toLowerCase().trim() === activeUserCustomData?.username?.toLowerCase().trim();
                                return (
                                  <tr key={u.id} className={`hover:bg-slate-100/40 transition-colors ${isSelf ? 'bg-emerald-50/30 font-semibold text-[#077a37]' : 'text-slate-700'}`}>
                                    <td className="px-4 py-3 text-xs text-slate-500 font-mono text-center">{idx + 1}</td>
                                    <td className="px-4 py-3 text-xs font-semibold">
                                      <div className="flex flex-col gap-1.5">
                                        <div className="flex items-center gap-1.5">
                                          {u.fullname}
                                          {isSelf && (
                                            <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-full font-black text-[8px] uppercase tracking-wider inline-block">
                                              Anda
                                            </span>
                                          )}
                                        </div>
                                        <div className="flex items-center gap-1.5">
                                          <span className={`w-1.5 h-1.5 rounded-full ${isSelf && isOnline ? 'bg-[#098f41] animate-pulse' : 'bg-slate-300'}`}></span>
                                          <span className={`text-[9px] font-bold ${isSelf && isOnline ? 'text-[#098f41]' : 'text-slate-400'}`}>
                                            {isSelf && isOnline ? 'Online' : 'Offline'}
                                          </span>
                                        </div>
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-xs font-mono font-bold text-slate-600">{u.username}</td>
                                    <td className="px-4 py-3 text-xs font-mono text-slate-600">{u.password}</td>
                                    <td className="px-4 py-3 text-xs text-center">
                                      <div className="flex items-center justify-center gap-1.5">
                                        <button
                                          onClick={() => handleEditUserClick(u)}
                                          className="p-2 rounded-xl text-indigo-600 hover:bg-indigo-50 hover:text-indigo-700 active:scale-90 transition-all cursor-pointer"
                                          title="Edit Nama/Sandi Pengguna"
                                        >
                                          <Pencil className="w-4 h-4" />
                                        </button>
                                        <button
                                          onClick={() => {
                                            if (u.username.toLowerCase().trim() === 'petugaspiket') {
                                              showToast('Akun Petugas Piket tidak bisa dihapus.', 'error');
                                              return;
                                            }
                                            handleDeleteUserClick(u);
                                          }}
                                          disabled={isSelf || u.username.toLowerCase().trim() === 'petugaspiket'}
                                          className={`p-2 rounded-xl transition-all ${
                                            (isSelf || u.username.toLowerCase().trim() === 'petugaspiket')
                                              ? 'text-slate-200 cursor-not-allowed opacity-50' 
                                              : 'text-rose-600 hover:bg-rose-50 hover:text-rose-700 active:scale-90 cursor-pointer'
                                          }`}
                                          title={
                                            isSelf 
                                              ? 'Tidak dapat menghapus akun Anda sendiri' 
                                              : u.username.toLowerCase().trim() === 'petugaspiket'
                                              ? 'Akun Petugas Piket tidak bisa dihapus'
                                              : 'Hapus Pengguna'
                                          }
                                        >
                                          <Trash2 className="w-4.5 h-4.5" />
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="w-full h-px bg-slate-100 mb-6"></div>
                  
                  <div className="w-full mb-8 text-left bg-indigo-50 border border-indigo-100 rounded-[1.5rem] p-6 shadow-sm">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                      <div className="flex items-center gap-2.5">
                        <div className="p-2 bg-indigo-100 rounded-lg text-indigo-600">
                          <Key className="w-5 h-5" />
                        </div>
                        <div>
                          <h3 className="text-base font-extrabold text-slate-800 tracking-tight">Generator Token Pendaftaran</h3>
                          <p className="text-[10px] text-slate-600">Buat token khusus yang diperlukan saat registrasi user baru.</p>
                        </div>
                      </div>
                      <button
                        onClick={async () => {
                          const token = Math.random().toString(36).substring(2, 10).toUpperCase();
                          setGeneratedToken(token);
                          try {
                            await setDoc(doc(dbDefault, 'register_tokens', token), {
                              createdAt: new Date().toISOString(),
                              createdBy: activeAuth.currentUser?.uid || 'admin',
                              used: false
                            });
                            showToast('Token berhasil dibuat dan disimpan.', 'success');
                          } catch (err) {
                            console.error('Gagal menyimpan token:', err);
                            showToast('Gagal menyimpan token ke database pusat.', 'error');
                            try { handleFirestoreError(err, OperationType.WRITE, 'register_tokens'); } catch { /* ignore */ }
                          }
                        }}
                        className="px-4 py-2 flex gap-2 items-center justify-center bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold rounded-xl transition-all self-start sm:self-center"
                      >
                        <RefreshCw className="w-3.5 h-3.5" />
                        Generate Token Baru
                      </button>
                    </div>

                    {generatedToken && (
                      <div className="mt-4 p-4 bg-white border border-indigo-100 rounded-xl text-center space-y-2">
                         <p className="text-xs text-slate-600 font-semibold mb-1">Token Terakhir Anda:</p>
                         <div className="flex items-center justify-center gap-3">
                           <div className="text-3xl font-black font-mono text-indigo-700 tracking-[0.2em] bg-indigo-50/50 py-3 px-6 rounded-lg border border-dashed border-indigo-200">
                             {generatedToken}
                           </div>
                           <button
                             onClick={() => {
                               navigator.clipboard.writeText(generatedToken);
                               showToast('Token disalin ke clipboard!', 'success');
                             }}
                             className="p-3 bg-indigo-100 hover:bg-indigo-200 text-indigo-600 rounded-lg transition-colors border border-indigo-200 shadow-sm"
                             title="Salin Token"
                           >
                             <Copy className="w-6 h-6" />
                           </button>
                         </div>
                         <p className="text-[10px] text-indigo-500">Salin token ini dan berikan kepada guru/user yang ingin mendaftar.</p>
                      </div>
                    )}
                  </div>
                </>
              )}

              <div className="w-full h-px bg-slate-100 mb-6"></div>



              {/* Maintenance & Reset Data Section */}
              <div className="w-full mb-8 text-left bg-stone-50 border-2 border-slate-300 rounded-2xl p-5">
                <h3 className="text-sm font-black text-slate-800 uppercase tracking-widest mb-2 flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse"></span>
                  Pusat Pemeliharaan &amp; Reset Data
                </h3>
                <p className="text-xs text-slate-600 mb-5 leading-relaxed">
                  Kelola database absensi Anda secara fleksibel. Lakukan pembersihan riwayat secara berkala untuk menyambut ajaran baru maupun reset total.
                </p>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <button
                    onClick={() => {
                      setResetModalType('new_semester');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                      setNewSemesterChoice(profileData.semester === 'Ganjil' ? 'Genap' : 'Ganjil');
                      setNewTahunPelajaran(profileData.tahunPelajaran || '');
                    }}
                    className="flex flex-col items-start gap-1.5 p-4 bg-white hover:bg-emerald-50/50 border-2 border-slate-300 hover:border-emerald-300 rounded-xl transition-all text-left group cursor-pointer"
                  >
                    <span className="flex items-center gap-2 text-xs font-bold text-slate-700 group-hover:text-emerald-800 transition-colors">
                      <span className="p-1 px-1.5 bg-emerald-100 rounded text-emerald-800 text-[10px] font-black">SEM</span>
                      Reset Semester Baru
                    </span>
                    <span className="text-[10px] text-slate-500 leading-normal group-hover:text-emerald-700 transition-colors">
                      Hapus riwayat sesi absensi semester lalu dengan konfirmasi password. Data siswa &amp; kelas tetap aman.
                    </span>
                  </button>

                  <button
                    onClick={() => {
                      setResetModalType('new_year');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                    }}
                    className="flex flex-col items-start gap-1.5 p-4 bg-white hover:bg-amber-50/50 border-2 border-slate-300 hover:border-amber-200 rounded-xl transition-all text-left group cursor-pointer"
                  >
                    <span className="flex items-center gap-2 text-xs font-bold text-slate-700 group-hover:text-amber-800 transition-colors">
                      <span className="p-1 px-1.5 bg-amber-100 rounded text-amber-700 text-[10px] font-black">TA</span>
                      Reset Tahun Ajaran Baru
                    </span>
                    <span className="text-[10px] text-slate-500 leading-normal group-hover:text-amber-600 transition-colors">
                      Hapus semua riwayat absensi, tapi tetap pertahankan seluruh biodata siswa &amp; daftar kelas.
                    </span>
                  </button>

                  <button
                    onClick={() => {
                      setResetModalType('everything');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                    }}
                    className="flex flex-col items-start gap-1.5 p-4 bg-white hover:bg-rose-50/50 border-2 border-slate-300 hover:border-rose-200 rounded-xl transition-all text-left group cursor-pointer"
                  >
                    <span className="flex items-center gap-2 text-xs font-bold text-slate-700 group-hover:text-rose-800 transition-colors">
                      <span className="p-1 px-1.5 bg-rose-100 rounded text-rose-700 text-[10px] font-black">ALL</span>
                      Reset Semua Data
                    </span>
                    <span className="text-[10px] text-slate-500 leading-normal group-hover:text-rose-500 transition-colors">
                      Hapus seluruh daftar siswa, kelas, &amp; seluruh riwayat absensi secara total dan permanen.
                    </span>
                  </button>
                </div>
              </div>

              <div className="w-full h-px bg-slate-100 mb-6"></div>

              <button 
                  onClick={() => setShowLogoutConfirm(true)}
                  className="w-full flex items-center justify-center gap-2 bg-rose-50 text-rose-600 font-bold py-4 px-6 rounded-2xl hover:bg-rose-100 transition-all border border-rose-100"
              >
                  <LogOut className="w-5 h-5" />
                  Keluar dari Aplikasi
              </button>
            </div>
          </div>
        );
      default:
        return (
          <div className="flex flex-col items-center justify-center h-full p-8">
            <h2 className="text-xl font-bold text-slate-800">Konten untuk {activeTab} akan datang.</h2>
          </div>
        );
    }
  };
  
  // Define UI before return
  const splashNode = (
    <AnimatePresence mode="wait">
      {showSplash && (
        isFirstSessionLoad ? (
          <motion.div
            key="splash-white"
            initial={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.8, ease: [0.25, 1, 0.5, 1] }}
            className="fixed inset-0 z-[9999] bg-white flex flex-col items-center justify-center p-6 select-none"
          >
            {/* Ambient lighting */}
            <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(9,143,65,0.05)_0%,transparent_70%)] pointer-events-none" />
            
            <motion.div
              initial={{ scale: 0.96, opacity: 0, y: 10 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              transition={{ duration: 1.0, ease: [0.16, 1, 0.3, 1] }}
              className="flex flex-col items-center relative z-10"
            >
              {/* School Logo Container */}
              <motion.div
                initial={{ scale: 0.85, opacity: 0, y: -10 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                transition={{ delay: 0.15, duration: 0.8, ease: "easeOut" }}
                className="w-32 h-32 md:w-36 md:h-36 mb-6 flex items-center justify-center"
              >
                <img 
                  src="/school_logo.png" 
                  alt="SMART DF Logo" 
                  className="w-full h-full object-contain filter drop-shadow-[0_12px_24px_rgba(9,143,65,0.15)]"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = "https://drive.google.com/thumbnail?id=1c0ibueBZudROdPwR1oJKDC4y0HnJ1R4n&sz=w1000";
                  }}
                  referrerPolicy="no-referrer"
                />
              </motion.div>

              {/* Micro divider */}
              <motion.div
                initial={{ scaleX: 0, opacity: 0 }}
                animate={{ scaleX: 1, opacity: 1 }}
                transition={{ delay: 0.25, duration: 0.8, ease: "easeOut" }}
                className="h-[1.5px] w-12 bg-[#098f41]/30 rounded-full mb-6 origin-center"
              />

              <motion.h1
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.35, duration: 0.8 }}
                className="text-[#098f41] text-3xl sm:text-4xl font-extrabold tracking-tight text-center"
              >
                SMART DF App
              </motion.h1>

              <motion.p
                initial={{ y: 8, opacity: 0 }}
                animate={{ y: 0, opacity: 1 }}
                transition={{ delay: 0.5, duration: 0.8, ease: "easeOut" }}
                className="text-[#077a37] text-xs sm:text-[13px] font-medium uppercase tracking-[0.22em] mt-3.5 text-center"
              >
                Sistem Informasi Absensi Digital
              </motion.p>
                
              <motion.p
                initial={{ opacity: 0 }}
                animate={{ opacity: 0.8 }}
                transition={{ delay: 0.7, duration: 0.9 }}
                className="text-slate-400 text-[9px] font-bold tracking-widest mt-1.5 uppercase text-center"
              >
                Future • Religious • Educative • Smart • Harmony
              </motion.p>

              {/* Micro progress bar */}
              <div className="w-48 h-[3px] bg-slate-100 rounded-full mt-10 overflow-hidden relative border border-slate-200/50">
                <motion.div
                  initial={{ left: "-100%" }}
                  animate={{ left: "100%" }}
                  transition={{
                    repeat: Infinity,
                    duration: 1.8,
                    ease: "easeInOut"
                  }}
                  className="absolute top-0 bottom-0 w-1/2 bg-gradient-to-r from-transparent via-[#098f41] to-transparent"
                />
              </div>
            </motion.div>
          </motion.div>
        ) : (
          <motion.div
            key="splash-refresh"
            initial={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.4, ease: "easeInOut" }}
            className="fixed inset-0 z-[9999] bg-white flex flex-col items-center justify-center p-6 select-none"
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              transition={{ duration: 0.5, ease: "easeOut" }}
              className="flex flex-col items-center relative z-10"
            >
              {/* School Logo Container */}
              <div className="w-24 h-24 sm:w-28 sm:h-28 mb-4 flex items-center justify-center">
                <img 
                  src="/school_logo.png" 
                  alt="SMART DF Logo" 
                  className="w-full h-full object-contain filter drop-shadow-sm"
                  onError={(e) => {
                    (e.target as HTMLImageElement).src = "https://drive.google.com/thumbnail?id=1c0ibueBZudROdPwR1oJKDC4y0HnJ1R4n&sz=w1000";
                  }}
                  referrerPolicy="no-referrer"
                />
              </div>

              <h1 className="text-[#098f41] text-2xl sm:text-3xl font-extrabold tracking-tight text-center">
                SMART DF App
              </h1>

              {/* Animated Memuat Halaman Badge */}
              <div className="flex items-center gap-2.5 mt-5 bg-emerald-50 border border-[#098f41]/20 px-5 py-2 rounded-full shadow-xs">
                <Loader2 className="w-4 h-4 text-[#098f41] animate-spin shrink-0" />
                <motion.p
                  animate={{ opacity: [0.6, 1, 0.6] }}
                  transition={{ repeat: Infinity, duration: 1.5, ease: "easeInOut" }}
                  className="text-[#098f41] font-extrabold text-xs tracking-wider uppercase"
                >
                  Memuat Data...
                </motion.p>
              </div>

              <div className="w-48 h-1.5 bg-slate-100 rounded-full mt-5 overflow-hidden relative border border-slate-200/60">
                <motion.div
                  initial={{ left: "-100%" }}
                  animate={{ left: "100%" }}
                  transition={{
                    repeat: Infinity,
                    duration: 1.6,
                    ease: "easeInOut"
                  }}
                  className="absolute top-0 bottom-0 w-1/2 bg-[#098f41] rounded-full"
                />
              </div>
            </motion.div>
          </motion.div>
        )
      )}
    </AnimatePresence>
  );

  const GlobalConnectivityBanner = () => (
    <AnimatePresence>
      {(!isOnline || !firebaseConnected || syncStatus === 'error' || quotaExceeded) && (
        <motion.div 
          initial={{ height: 0, opacity: 0 }}
          animate={{ height: 'auto', opacity: 1 }}
          exit={{ height: 0, opacity: 0 }}
          className="overflow-hidden z-[100] sticky top-0 shrink-0 hidden md:block"
        >
          {/* Main Sync Error / Offline Banner */}
          {(!isOnline || !firebaseConnected || syncStatus === 'error') && (
            <div className="bg-rose-600 text-white shadow-lg">
              <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-white/20 rounded-xl animate-pulse">
                    <Cloud className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="text-sm font-black leading-tight flex items-center gap-2">
                       {!isOnline ? 'Koneksi Internet Terputus' : 'Sinkronisasi Cloud Bermasalah'}
                       <span className="px-1.5 py-0.5 bg-rose-500 rounded text-[9px] uppercase tracking-widest">{!isOnline ? 'OFFLINE' : 'ERR_CLOUD'}</span>
                    </h4>
                    <p className="text-[11px] opacity-90 font-medium mt-0.5">
                      {!isOnline 
                        ? 'Periksa koneksi Wi-Fi atau data seluler Anda. Data yang Anda buat saat ini tersimpan sementara di peramban ini dan akan diunggah otomatis saat kembali online.' 
                        : lastSyncError || 'Terjadi gangguan sinkronisasi dengan database Firebase. Harap periksa izin akses rules Firestore atau muat ulang halaman.'}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => window.location.reload()}
                    className="px-3 py-2 bg-white text-rose-600 rounded-xl text-[10px] font-black uppercase tracking-wider hover:bg-rose-50 transition-all shadow-sm active:scale-95 shrink-0"
                  >
                    Muat Ulang
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Quota Banner */}
          {quotaExceeded && isOnline && firebaseConnected && syncStatus !== 'error' && (
            <div className="bg-amber-500 text-white shadow-lg">
               <div className="max-w-7xl mx-auto px-4 py-2.5 flex items-center justify-between gap-4">
                  <div className="flex items-center gap-3">
                    <Database className="w-4 h-4" />
                    <p className="text-xs font-bold leading-tight">
                      Kuota harian database tercapai. Beberapa pembaruan data mungkin tertunda masuk ke Cloud sampai besok.
                    </p>
                  </div>
                  <button 
                    onClick={() => setShowQuotaModal(true)}
                    className="px-3 py-1.5 bg-white/20 rounded-lg text-[10px] font-bold uppercase transition-colors hover:bg-white/30"
                  >
                    Info Lanjut
                  </button>
               </div>
            </div>
          )}
        </motion.div>
      )}
    </AnimatePresence>
  );

  // Kami menghapus screen loading terpisah agar transisi langsung dari screen hijau (splash) ke aplikasi utama tanpa ada jeda/screen kedua.

  if (!isLoggedIn) {
    return (
      <div className="h-[100dvh] bg-stone-50 flex flex-col relative overflow-y-auto font-sans pb-8">
        <GlobalConnectivityBanner />
        {splashNode}
        <div className="w-full mx-auto px-6 py-8 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-12 md:gap-8 items-center relative z-10 my-auto max-w-5xl">
          
          {/* Left/Top Content */}
          <div className="flex flex-col text-center lg:text-left space-y-4 md:pr-8">
              <div className="flex flex-col lg:items-start items-center">
                <h1 className="text-4xl md:text-[2.75rem] font-black text-[#098f41] tracking-tighter drop-shadow-sm scale-y-105 origin-bottom lg:origin-left">SMART DF App</h1>
                <p className="text-[8px] sm:text-[10px] font-bold text-slate-600 tracking-[0.2em] mt-3 mb-2">FUTURE • RELIGIOUS • EDUCATIVE • SMART • HARMONY</p>
              </div>
            
            <h2 className="text-2xl md:text-3xl font-bold text-slate-800 leading-snug mt-6 lg:text-left text-center max-w-sm mx-auto lg:mx-0">
              Sistem Monitoring Absensi Real - time SMA Darul Falah Cihampelas
            </h2>
            
            <p className="text-sm text-slate-600 leading-relaxed max-w-md mx-auto lg:mx-0 mt-4 lg:text-left text-center">
              SMART DF adalah aplikasi absensi digital khusus untuk SMA Darul Falah Cihampelas yang digunakan oleh Petugas Piket, Wali Kelas dan Guru Mata Pelajaran Untuk mencatat kehadiran murid setiap hari secara cepat,akurat dan real-time.
            </p>

            {/* PWA Install Info Box on Login screen */}
            <div className="mt-6 p-5 bg-white/75 backdrop-blur-md rounded-2xl border border-slate-100 shadow-[0_12px_30px_-15px_rgba(9,143,65,0.18)] text-left max-w-md mx-auto lg:mx-0">
              <div className="flex items-center gap-2 mb-2 text-slate-800">
                <div className="p-1.5 bg-[#098f41]/10 text-[#098f41] rounded-lg">
                  <Download className="w-4 h-4" />
                </div>
                <span className="font-extrabold text-[10px] tracking-widest text-[#098f41] uppercase">Aplikasi PWA (Absensi Seluler)</span>
              </div>
              {isInsideIframe ? (
                <div>
                  <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                    Agar aplikasi absensi digital ini bisa diinstal langsung ke layar utama HP / Laptop Anda (tanpa Play Store), Anda perlu membukanya di tab browser mandiri.
                  </p>
                  <a
                    href="https://ais-pre-56w2g4yxpoxk4k23siyzdq-901834158843.asia-southeast1.run.app"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-[#098f41] hover:bg-[#077a37] text-white text-[10px] font-black uppercase tracking-wider rounded-lg transition-all shadow-[0_4px_10px_rgba(9,143,65,0.25)] hover:scale-102"
                  >
                    Buka Tab Mandiri & Instal <ArrowRight className="w-3.5 h-3.5" />
                  </a>
                </div>
              ) : isAppInstalled ? (
                <p className="text-[11px] text-emerald-600 font-bold leading-relaxed">
                  🎉 Aplikasi Absensi PWA ini telah terpasang dengan sukses di perangkat Anda!
                </p>
              ) : isInstallable ? (
                <div>
                  <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                    Aplikasi ini siap dipasang langsung tanpa App Store / Play Store. Cepat, ringan, dan hemat kuota internet.
                  </p>
                  <button
                    onClick={handleInstallClick}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#098f41] text-white hover:bg-[#077a37] text-[10px] font-black uppercase tracking-wider rounded-lg transition-all shadow-md shadow-emerald-200 hover:scale-102 active:scale-98"
                  >
                    Instal Sekarang <ArrowRight className="w-3.5 h-3.5" />
                  </button>
                </div>
              ) : (
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  Buka menu browser Anda (klik tombol titik tiga di Chrome, atau tombol Share/Bagikan di Safari iOS), lalu pilih <b>"Instal Aplikasi"</b> atau <b>"Tambahkan ke Layar Utama"</b>.
                </p>
              )}
            </div>
          </div>

          {/* Right/Bottom Content - Login & Recovery Card */}
          <div className="flex justify-center w-full">
            <div className="bg-white p-7 md:p-10 rounded-2xl shadow-[0_20px_50px_-12px_rgba(0,0,0,0.1)] border border-slate-100/80 w-full max-w-md">
              {isForgotPassword ? (
                <div>
                  <div className="text-center mb-6">
                    <h3 className="text-xl font-bold text-slate-800">Lupa Password / Username</h3>
                    <p className="text-xs text-slate-600 mt-1 lines-relaxed">
                      Sistem Pencarian Akun & Kredensial Sekolah Mandiri
                    </p>
                  </div>

                  <form onSubmit={async (e) => {
                    e.preventDefault();
                    if (!recoverySearchVal.trim()) {
                      showToast('Masukkan nilai pencarian.', 'error');
                      return;
                    }

                    setIsRecoveryLoading(true);
                    setRecoveryResult(null);

                    try {
                      if (recoverySearchType === 'username') {
                        // Look up password from Username/Email
                        const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', recoverySearchVal.toLowerCase().trim()));
                        if (lookupDoc.exists()) {
                          setRecoveryResult([lookupDoc.data()]);
                          showToast('Siswa/Akun Ditemukan', 'success');
                        } else {
                          setRecoveryResult([]);
                          showToast('Akun tidak terregistrasi di sistem pusat.', 'error');
                        }
                      } else {
                        // Look up Username/Email from Password!
                        const q = query(collection(dbDefault, 'custom_accounts'), where('password', '==', recoverySearchVal.trim()));
                        const qSnap = await getDocs(q);
                        if (!qSnap.empty) {
                          const results = qSnap.docs.map(d => d.data());
                          setRecoveryResult(results);
                          showToast('Siswa/Akun Ditemukan', 'success');
                        } else {
                          setRecoveryResult([]);
                          showToast('Gagal menemukan data dengan kata sandi tersebut.', 'error');
                        }
                      }
                    } catch (err) {
                      const error = err as Error;
                      showToast('Pencarian Gagal: ' + error.message, 'error');
                    } finally {
                      setIsRecoveryLoading(false);
                    }
                  }} className="space-y-4">
                    
                    {/* Toggle Search Mode */}
                    <div className="bg-slate-100 p-1.5 rounded-xl flex gap-1 mb-2">
                      <button
                        type="button"
                        onClick={() => {
                          setRecoverySearchType('username');
                          setRecoveryResult(null);
                        }}
                        className={`flex-1 text-center py-2 text-xs font-bold rounded-lg transition-all ${
                          recoverySearchType === 'username' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-600 hover:text-slate-700'
                        }`}
                      >
                        Cari Dengan Username
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setRecoverySearchType('password');
                          setRecoveryResult(null);
                        }}
                        className={`flex-1 text-center py-2 text-xs font-bold rounded-lg transition-all ${
                          recoverySearchType === 'password' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-600 hover:text-slate-700'
                        }`}
                      >
                        Cari Dengan Password
                      </button>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-600 mb-2 tracking-wider">
                        {recoverySearchType === 'username' ? 'MASUKKAN USERNAME / EMAIL' : 'MASUKKAN KATA SANDI'}
                      </label>
                      <div className="relative">
                        {recoverySearchType === 'username' ? (
                          <UserIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        ) : (
                          <Key className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        )}
                        <input 
                          type="text" 
                          className="w-full pl-11 pr-4 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-lg focus:ring-2 focus:ring-[#098f41] focus:border-[#098f41] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm" 
                          placeholder={recoverySearchType === 'username' ? 'Masukkan username / email' : 'Ketik kata sandi'} 
                          value={recoverySearchVal}
                          onChange={e => setRecoverySearchVal(e.target.value)}
                          required 
                        />
                      </div>
                    </div>

                    <button 
                      type="submit" 
                      disabled={isRecoveryLoading}
                      className="w-full bg-[#098f41] text-white font-bold py-3 rounded-lg hover:bg-[#077a37] transition-colors text-sm shadow-sm flex items-center justify-center gap-2"
                    >
                      {isRecoveryLoading ? (
                        <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                      ) : (
                        'Temukan Kredensial Saya'
                      )}
                    </button>

                    {/* RECOVERY RESULT DISPLAY DISPLAY */}
                    {recoveryResult !== null && (
                      <div className="mt-4 p-4 rounded-xl border border-dashed text-left space-y-3 bg-stone-50">
                        {recoveryResult.length > 0 ? (
                          <div>
                            <span className="block text-[10px] font-bold text-[#098f41] mb-2 uppercase">✓ Kredensial Ditemukan</span>
                            {recoveryResult.map((acc, index) => (
                              <div key={index} className="space-y-2 border-t border-slate-200 pt-2 first:border-0 first:pt-0">
                                <div>
                                  <span className="text-[9px] font-bold text-slate-500 block">NAMA LENGKAP</span>
                                  <span className="text-xs font-bold text-slate-800">{acc.fullname}</span>
                                </div>
                                <div className="grid grid-cols-2 gap-2">
                                  <div>
                                    <span className="text-[9px] font-bold text-slate-500 block">USERNAME</span>
                                    <span className="text-xs font-mono font-bold text-slate-800 break-all bg-white px-1.5 py-0.5 rounded border border-slate-100 block">{acc.username}</span>
                                  </div>
                                  <div>
                                    <span className="text-[9px] font-bold text-slate-500 block">PASSWORD</span>
                                    <span className="text-xs font-mono font-bold text-lime-700 bg-white px-1.5 py-0.5 rounded border border-slate-100 block">{acc.password}</span>
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        ) : (
                          <div className="text-center py-2">
                            <span className="text-xs font-bold text-rose-500">❌ Tidak ada hasil yang cocok!</span>
                            <p className="text-[10px] text-slate-500 mt-1">Pastikan email / password yang dimasukkan tepat.</p>
                          </div>
                        )}
                      </div>
                    )}

                    <div className="relative flex items-center py-2">
                      <div className="flex-grow border-t border-slate-100"></div>
                      <span className="flex-shrink-0 mx-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Atau</span>
                      <div className="flex-grow border-t border-slate-100"></div>
                    </div>

                    <button 
                      type="button" 
                      onClick={() => {
                        setIsForgotPassword(false);
                        setRecoveryResult(null);
                        setRecoverySearchVal('');
                      }}
                      className="w-full bg-transparent border-2 border-slate-400 text-slate-600 font-bold py-3.5 rounded-lg hover:bg-slate-50 transition-colors text-sm"
                    >
                      Kembali ke Halaman Masuk
                    </button>
                  </form>
                </div>
              ) : (
                <div>
                  <div className="text-center mb-6">
                    <h3 className="text-xl font-bold text-slate-800">Selamat Datang</h3>
                    <p className="text-xs text-slate-600 mt-1">Silakan masuk ke akun atau sekolah Anda</p>
                  </div>

                  <form onSubmit={async (e) => { 
                    e.preventDefault(); 
                    setIsAuthLoading(true);

                    try {
                      const rawInput = (authEmail || '').trim();
                      const resolvedUser = rawInput.toLowerCase();
                      const usernamePrefix = resolvedUser.includes('@') ? resolvedUser.split('@')[0] : resolvedUser;
                      
                      const isDefaultAdmin = (resolvedUser === 'admin' || usernamePrefix === 'admin') && authPassword === '123456@#';
                      const isDefaultPiket = (resolvedUser === 'petugaspiket' || usernamePrefix === 'petugaspiket') && authPassword === 'petugaspiket123#';

                      let lookupDoc: DocumentSnapshot | null = null;
                      
                      // Auto-seed default admin if needed
                      if (isDefaultAdmin) {
                        const adminSnap = await getDoc(doc(dbDefault, 'custom_accounts', 'admin'));
                        if (!adminSnap.exists()) {
                          await setDoc(doc(dbDefault, 'custom_accounts', 'admin'), {
                            fullname: 'Endang Sukmaya,S.Pd.I., Gr.',
                            username: 'admin',
                            password: '123456@#',
                            role: 'Admin',
                            configText: '',
                            createdAt: new Date().toISOString(),
                            updatedAt: new Date().toISOString()
                          });
                        }
                        lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', 'admin'));
                      } else if (isDefaultPiket) {
                        const piketSnap = await getDoc(doc(dbDefault, 'custom_accounts', 'petugaspiket'));
                        if (!piketSnap.exists()) {
                          await setDoc(doc(dbDefault, 'custom_accounts', 'petugaspiket'), {
                            fullname: 'Petugas Piket',
                            username: 'petugaspiket',
                            password: 'petugaspiket123#',
                            role: 'Petugas Piket',
                            configText: '',
                            createdAt: new Date().toISOString(),
                            updatedAt: new Date().toISOString()
                          });
                        }
                        lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', 'petugaspiket'));
                      } else {
                        // 1. Try exact document match by username or email
                        const directSnap = await getDoc(doc(dbDefault, 'custom_accounts', resolvedUser));
                        if (directSnap.exists()) {
                          lookupDoc = directSnap;
                        } else if (usernamePrefix !== resolvedUser) {
                          // 2. Try prefix match
                          const prefixSnap = await getDoc(doc(dbDefault, 'custom_accounts', usernamePrefix));
                          if (prefixSnap.exists()) {
                            lookupDoc = prefixSnap;
                          }
                        }

                        // 3. Fallback query if stored under username or email field
                        if (!lookupDoc || !lookupDoc.exists()) {
                          try {
                            const qUser = query(collection(dbDefault, 'custom_accounts'), where('username', '==', resolvedUser));
                            const qSnap = await getDocs(qUser);
                            if (!qSnap.empty) {
                              lookupDoc = qSnap.docs[0];
                            } else if (usernamePrefix !== resolvedUser) {
                              const qPrefix = query(collection(dbDefault, 'custom_accounts'), where('username', '==', usernamePrefix));
                              const qPrefixSnap = await getDocs(qPrefix);
                              if (!qPrefixSnap.empty) {
                                lookupDoc = qPrefixSnap.docs[0];
                              }
                            }
                          } catch (qErr) {
                            console.warn("Query fallback error:", qErr);
                          }
                        }
                      }

                      // If doc found in custom_accounts:
                      if (lookupDoc && lookupDoc.exists()) {
                        const accData = lookupDoc.data();

                        // Verify password if stored
                        if (accData.password && accData.password !== authPassword && !isDefaultAdmin && !isDefaultPiket) {
                          setLoginError({
                            title: 'Kata Sandi Salah!',
                            message: `Akun "${rawInput}" terdaftar di sistem FRESH, namun kata sandi yang Anda masukkan salah.`,
                            recommendations: [
                              'Periksa kembali penulisan kata sandi Anda (pastikan huruf besar/kecil prasyarat Caps Lock sudah tepat).',
                              'Gunakan fitur Lupa Akun / Kata Sandi di bawah jika Anda lupa.'
                            ],
                            username: rawInput
                          });
                          setIsAuthLoading(false);
                          return;
                        }

                        const targetUsername = accData.username || usernamePrefix || resolvedUser;
                        const customUserData = {
                          fullname: accData.fullname || targetUsername,
                          username: targetUsername,
                          configText: accData.configText || '',
                          role: accData.role || ''
                        };

                        setActiveUserCustomData(customUserData);
                        localStorage.setItem('kaguci_active_custom_user', JSON.stringify(customUserData));
                        localStorage.setItem('kaguci_has_logged_in', 'true');

                        if (accData.photoURL) {
                          setAvatarUrl(accData.photoURL);
                          safeSetLocalStorage(`kaguci_avatar_${targetUsername.toLowerCase()}`, accData.photoURL);
                          safeSetLocalStorage('kaguci_avatar_current', accData.photoURL);
                        }

                        if (accData.profileData) {
                          setProfileData((prev: typeof profileData) => ({ ...prev, ...accData.profileData }));
                          localStorage.setItem(`kaguci_profile_${targetUsername.toLowerCase()}`, JSON.stringify(accData.profileData));
                        }

                        // Authenticate or safely establish Firebase user
                        const authUser = await safeFirebaseAuth(auth, authEmail, authPassword, targetUsername);
                        if (authUser) {
                          setCurrentUser(authUser);
                        } else {
                          setCurrentUser({
                            uid: targetUsername,
                            email: toAuthEmail(targetUsername),
                            displayName: accData.fullname || targetUsername,
                            photoURL: accData.photoURL || null
                          } as unknown as User);
                        }

                        showToast('Selamat Datang Kembali! Login berhasil.', 'success');
                        try {
                          window.history.replaceState(null, '', window.location.pathname);
                        } catch {
                          window.location.hash = '';
                        }

                        const greetingName = accData.fullname || targetUsername || "Pengguna";
                        const phrase = `Selamat datang ${greetingName} di aplikasi SMART DF App.`;
                        sessionStorage.setItem(`kaguci_welcomed_${targetUsername.toLowerCase().trim()}`, 'true');
                        speakText(phrase);

                        setIsLoggedIn(true);
                        setIsAuthLoading(false);
                        return;
                      }

                      // If not found in custom_accounts, try Firebase Auth directly
                      try {
                        const authUser = await safeFirebaseAuth(auth, authEmail, authPassword, usernamePrefix);
                        if (authUser && (authUser.uid || auth.currentUser)) {
                          // Successfully authenticated! Auto-provision the custom_accounts entry
                          const newAccountData = {
                            fullname: authUser.displayName || rawInput,
                            username: usernamePrefix,
                            password: authPassword,
                            role: 'Guru Mapel',
                            configText: '',
                            createdAt: new Date().toISOString(),
                            updatedAt: new Date().toISOString()
                          };
                          await setDoc(doc(dbDefault, 'custom_accounts', usernamePrefix), newAccountData, { merge: true });

                          const customUserData = {
                            fullname: newAccountData.fullname,
                            username: usernamePrefix,
                            configText: '',
                            role: newAccountData.role
                          };

                          setActiveUserCustomData(customUserData);
                          localStorage.setItem('kaguci_active_custom_user', JSON.stringify(customUserData));
                          localStorage.setItem('kaguci_has_logged_in', 'true');

                          setCurrentUser(authUser);
                          showToast('Selamat Datang! Login berhasil.', 'success');

                          const greetingName = newAccountData.fullname || usernamePrefix;
                          const phrase = `Selamat datang ${greetingName} di aplikasi SMART DF App.`;
                          sessionStorage.setItem(`kaguci_welcomed_${usernamePrefix.toLowerCase().trim()}`, 'true');
                          speakText(phrase);

                          setIsLoggedIn(true);
                          setIsAuthLoading(false);
                          return;
                        }
                      } catch (fbErr) {
                        const err = fbErr as Error;
                        const errMsg = (err.message || '').toLowerCase();
                        if (errMsg.includes('wrong-password') || errMsg.includes('invalid-credential') || errMsg.includes('invalid-email')) {
                          setLoginError({
                            title: 'Kata Sandi Salah!',
                            message: `Username atau Email "${authEmail}" tidak cocok dengan kata sandi yang dimasukkan.`,
                            recommendations: [
                              'Periksa kembali penulisan kata sandi Anda.',
                              'Pastikan Caps Lock tidak aktif saat mengetik kata sandi.'
                            ],
                            username: authEmail
                          });
                          setIsAuthLoading(false);
                          return;
                        }
                      }

                      // If neither found nor valid
                      setLoginError({
                        title: 'Akun Belum Terdaftar!',
                        message: `Username atau Email "${authEmail}" tidak terdaftar di sistem.`,
                        recommendations: [
                          'Periksa kembali penulisan nama pengguna / email Anda (pastikan tidak ada salah ketik).',
                          'Klik tombol "Tambah Akun Baru" di bagian bawah untuk mendaftarkan akun baru.'
                        ],
                        username: authEmail
                      });
                      setIsAuthLoading(false);
                    } catch (error) {
                      const err = error as Error;
                      console.warn("Login error:", err);
                      setLoginError({
                        title: 'Gagal Masuk',
                        message: `Terjadi kendala saat memproses login: ${err.message}`,
                        recommendations: [
                          'Periksa koneksi internet perangkat Anda.',
                          'Muat ulang halaman ini dan coba kembali.'
                        ],
                        username: authEmail
                      });
                    } finally {
                      setIsAuthLoading(false);
                    }
                  }} 
                  name="login"
                  autoComplete="on"
                  className="space-y-5">
                    <div>
                      <label htmlFor="auth-username" className="block text-[10px] font-bold text-slate-600 mb-2 tracking-wider">USERNAME / EMAIL</label>
                      <div className="relative">
                        <UserIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        <input 
                          id="auth-username"
                          name="username"
                          type="text" 
                          autoComplete="username"
                          className="w-full pl-11 pr-4 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-lg focus:ring-2 focus:ring-[#098f41] focus:border-[#098f41] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm" 
                          placeholder="Contoh: admin atau budi" 
                          value={authEmail}
                          onChange={e => setAuthEmail(e.target.value)}
                          required 
                        />
                      </div>
                    </div>

                    <div>
                      <label htmlFor="auth-password" className="block text-[10px] font-bold text-slate-600 mb-2 tracking-wider">KATA SANDI</label>
                      <div className="relative flex items-center">
                        <Key className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                        <input 
                          id="auth-password"
                          name="password"
                          type={showPassword ? 'text' : 'password'} 
                          autoComplete="current-password"
                          className="w-full pl-11 pr-12 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-lg focus:ring-2 focus:ring-[#098f41] focus:border-[#098f41] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-lg tracking-[0.2em]" 
                          placeholder="•••••" 
                          value={authPassword}
                          onChange={e => setAuthPassword(e.target.value)}
                          required 
                        />
                        {showPassword ? (
                          <EyeOff 
                            onClick={() => setShowPassword(false)} 
                            className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500 cursor-pointer hover:text-slate-600 transition-colors" 
                          />
                        ) : (
                          <Eye 
                            onClick={() => setShowPassword(true)} 
                            className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500 cursor-pointer hover:text-slate-600 transition-colors" 
                          />
                        )}
                      </div>
                    </div>

                    <div className="pt-2">
                      <button 
                        type="submit" 
                        className="w-full bg-[#098f41] text-white font-bold py-3.5 rounded-lg hover:bg-[#077a37] transition-colors text-sm shadow-sm active:scale-[0.98]"
                      >
                        Masuk Ke Sistem
                      </button>
                      
                      <div className="text-center mt-4">
                        <button 
                          type="button"
                          onClick={() => {
                            setIsForgotPassword(true);
                            setRecoveryResult(null);
                            setRecoverySearchVal('');
                          }}
                          className="inline-block text-xs font-bold text-[#098f41] hover:underline cursor-pointer focus:outline-none"
                        >
                          Lupa Password / Username?
                        </button>
                      </div>
                    </div>

                    <button 
                      type="button" 
                      onClick={() => setIsRegisterModalOpen(true)}
                      className="w-full bg-transparent border border-[#098f41] text-[#098f41] font-bold py-3.5 rounded-lg hover:bg-emerald-50 transition-colors text-sm mt-3 flex items-center justify-center gap-1"
                    >
                      Tambah Akun Baru
                    </button>
                  </form>
                </div>
              )}
            </div>
          </div>
        </div>

          {/* Kotak Dialog Tambah Akun Baru (Modal Dialog Box) */}
          <AnimatePresence>
            {isRegisterModalOpen && (
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 className="fixed inset-0 bg-[#020617]/70 backdrop-blur-sm z-[100] flex items-center justify-center overflow-y-auto p-4"
              >
                <motion.div 
                   initial={{ opacity: 0, scale: 0.95, y: 15 }}
                   animate={{ opacity: 1, scale: 1, y: 0 }}
                   exit={{ opacity: 0, scale: 0.95, y: 15 }}
                   className="bg-white rounded-3xl p-6 sm:p-8 pb-20 sm:pb-24 shadow-2xl max-w-lg w-[95%] my-8 space-y-6 max-h-[90vh] overflow-y-auto text-left m-4"
                >
                  <div>
                     <h3 className="text-xl font-bold text-slate-800">Tambah Akun Baru</h3>
                     <p className="text-xs text-slate-600 mt-1">Isi formulir pendaftaran akun menggunakan token valid dari Admin.</p>
                  </div>

                  <div className="space-y-4 text-left">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Nama Lengkap Pemilik</label>
                      <input 
                        type="text"
                        className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-xl focus:ring-2 focus:ring-[#098f41] focus:border-[#098f41] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                        placeholder="Contoh: Budi Santoso, S.Pd."
                        value={regFullName}
                        onChange={e => setRegFullName(e.target.value)}
                      />
                    </div>

                    <div>
                       <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Username / Email Akun</label>
                       <input 
                         type="text"
                         className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-xl focus:ring-2 focus:ring-[#098f41] focus:border-[#077a37] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                         placeholder="Contoh: admin atau budi"
                         value={regUsername}
                         onChange={e => setRegUsername(e.target.value)}
                       />
                    </div>

                    <div>
                       <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Kata Sandi (Min. 6 digit)</label>
                       <input 
                         type="text"
                         className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-xl focus:ring-2 focus:ring-[#098f41] focus:border-[#077a37] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                         placeholder="Ketik password untuk pendaftaran"
                         value={regPassword}
                         onChange={e => setRegPassword(e.target.value)}
                       />
                    </div>

                    <div>
                       <label className="block text-[10px] font-bold text-slate-600 mb-1.5 uppercase tracking-wider text-left">Token Pendaftaran (Dari Admin - Opsional)</label>
                       <input 
                         type="text"
                         className="w-full px-4 py-3 bg-[#fefce8] border-2 border-[#098f41] rounded-xl focus:ring-2 focus:ring-[#098f41] focus:border-[#077a37] transition-all outline-none text-slate-700 font-medium placeholder:text-slate-500 text-sm"
                         placeholder="Masukkan token 6-8 digit (Opsional)"
                         value={regToken}
                         onChange={e => setRegToken(e.target.value)}
                       />
                    </div>

                    <div>
                      <div className="flex justify-between items-center mb-1.5">
                        <label className="block text-[10px] font-bold text-slate-600 uppercase tracking-wider text-left">Status Koneksi Database</label>
                        <span className="px-2 py-0.5 bg-emerald-100 text-[#077a37] rounded-full font-bold text-[9px] uppercase tracking-wide">Otomatis & Terpusat</span>
                      </div>
                      <div className="p-4 bg-emerald-50/50 border border-lime-100/50 rounded-2xl space-y-2 text-left">
                        <p className="text-[11px] text-slate-600 leading-relaxed font-semibold">
                          Sistem akan mendaftarkan akun Anda secara langsung ke Database Pusat FRESH. Semua data presensi, absensi suara, dan data siswa akan aman terenkripsi dan disinkronkan secara otomatis.
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="flex flex-col sm:flex-row gap-3 pt-2">
                    <button 
                      type="button" 
                      onClick={() => {
                        setIsRegisterModalOpen(false);
                        setRegFullName('');
                        setRegUsername('');
                        setRegPassword('');
                        setRegToken('');
                      }}
                      className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-600 text-sm transition-colors"
                    >
                      Batal
                    </button>
                    <button 
                      type="button" 
                      disabled={!regFullName || !regUsername || !regPassword}
                      onClick={async () => {
                        setIsAuthLoading(true);
                        try {
                          const trimmedFullName = regFullName.trim();
                          const rawUsername = regUsername.toLowerCase().trim();
                          const resolvedUsername = rawUsername.includes('@') ? rawUsername.split('@')[0] : rawUsername.replace(/\s+/g, '');
                          const trimmedPassword = regPassword;

                          if (trimmedFullName.length < 3) {
                            showToast('Nama Lengkap minimal berisi 3 karakter.', 'error');
                            setIsAuthLoading(false);
                            return;
                          }
                          if (resolvedUsername.length < 3) {
                            showToast('Username minimal berisi 3 karakter (tanpa spasi).', 'error');
                            setIsAuthLoading(false);
                            return;
                          }
                          if (trimmedPassword.length < 6) {
                            showToast('Kata Sandi minimal berisi 6 karakter.', 'error');
                            setIsAuthLoading(false);
                            return;
                          }

                          const resolvedToken = regToken.trim();

                          if (resolvedToken) {
                            try {
                              const tokenDoc = await getDoc(doc(dbDefault, 'register_tokens', resolvedToken));
                              if (tokenDoc.exists()) {
                                const tokenData = tokenDoc.data();
                                if (tokenData?.used) {
                                  console.log('Token sudah terpakai, melanjutkan tanpa kendala.');
                                }
                              }
                            } catch (err) {
                              console.error('Gagal mengecek token:', err);
                            }
                          }

                          let usernameExists = false;
                          let usernameOwnerName = '';

                          // 1. Periksa apakah Username sudah terdaftar di database sistem pusat
                          try {
                            const lookupUsernameDoc = await getDoc(doc(dbDefault, 'custom_accounts', resolvedUsername));
                            if (lookupUsernameDoc.exists()) {
                              const data = lookupUsernameDoc.data();
                              usernameExists = true;
                              usernameOwnerName = data.fullname || 'N/A';
                            }
                          } catch (err) {
                            console.error('Kendala saat melacak username:', err);
                          }

                          if (usernameExists) {
                            setIsAuthLoading(false);
                            setRegistrationResult({
                              success: false,
                              title: 'Username Sudah Terdaftar!',
                              message: `Maaf, Nama Akun (Username) "${resolvedUsername}" sudah terdaftar di sistem pusat atas nama "${usernameOwnerName}".\n\nSilakan pilih Username lain yang berbeda.`
                            });
                            return;
                          }

                          // 2. Buat akun native di database pusat
                          await safeFirebaseAuth(auth, resolvedUsername, trimmedPassword, resolvedUsername);

                          // 3. Simpan data registrasi ke Firestore terpusat custom_accounts
                          const teacherProfile = {
                            namaGuruMapel: trimmedFullName,
                            nipGuruMapel: '',
                            mataPelajaran: '',
                            role: 'Guru Mapel',
                            semester: 'Ganjil',
                            tahunPelajaran: ''
                          };

                          await setDoc(doc(dbDefault, 'custom_accounts', resolvedUsername), {
                            fullname: trimmedFullName,
                            username: resolvedUsername,
                            password: trimmedPassword,
                            role: 'Guru Mapel',
                            profileData: teacherProfile,
                            configText: '',
                            createdAt: new Date().toISOString(),
                            updatedAt: new Date().toISOString()
                          }, { merge: true });

                          // 4. Otomatis daftarkan ke koleksi 'teachers' agar guru langsung muncul di sistem
                          try {
                            const teacherId = `teacher_${resolvedUsername}`;
                            await setDoc(doc(dbDefault, 'teachers', teacherId), {
                              id: teacherId,
                              name: trimmedFullName,
                              niy: '-',
                              dutyDay: '',
                              userId: resolvedUsername
                            }, { merge: true });
                          } catch (tErr) {
                            console.warn('Gagal menambahkan ke daftar guru:', tErr);
                          }

                          // 5. Tandai token jika digunakan
                          if (resolvedToken && resolvedToken !== 'SUPERADMINTOKEN') {
                            try {
                              await setDoc(doc(dbDefault, 'register_tokens', resolvedToken), {
                                 used: true,
                                 usedAt: new Date().toISOString(),
                                 usedBy: resolvedUsername
                              }, { merge: true });
                            } catch (err) {
                              console.error('Gagal menandai token sebagai terpakai:', err);
                            }
                          }

                          // Auto-login session for user
                          const customUserData = {
                            fullname: trimmedFullName,
                            username: resolvedUsername,
                            configText: '',
                            role: 'Guru Mapel'
                          };

                          setActiveUserCustomData(customUserData);
                          localStorage.setItem('kaguci_active_custom_user', JSON.stringify(customUserData));
                          localStorage.setItem('kaguci_has_logged_in', 'true');
                          setProfileData((prev: typeof profileData) => ({ ...prev, ...teacherProfile }));
                          setIsLoggedIn(true);

                          // On success: trigger results popup
                          setRegistrationResult({
                            success: true,
                            title: 'Pendaftaran Akun Berhasil!',
                            message: 'Akun Pengajar Anda siap digunakan dengan Cloud Database terpusat FRESH secara aman dan instan.',
                            fullname: trimmedFullName,
                            username: resolvedUsername,
                            projectId: 'Database Terpusat FRESH'
                          });

                          // Autofill login credentials for easy access
                          setAuthEmail(resolvedUsername);
                          setAuthPassword(trimmedPassword);
                          
                          try {
                            window.history.replaceState(null, '', window.location.pathname);
                          } catch {
                            window.location.hash = '';
                          }
                          
                          // Clean fields and close form
                          setIsRegisterModalOpen(false);
                          setRegFullName('');
                          setRegUsername('');
                          setRegPassword('');
                          setRegToken('');
                        } catch (error) {
                          const err = error as Error;
                          let IndonesianError = err.message;
                          if (err.message.includes('email-already-in-use')) {
                            const lookupUsername = regUsername.toLowerCase().trim();
                            let centralUserFound: { fullname?: string; username?: string; configText?: string } | null = null;
                            try {
                              const lookupDoc = await getDoc(doc(dbDefault, 'custom_accounts', lookupUsername));
                              if (lookupDoc.exists()) {
                                const data = lookupDoc.data();
                                centralUserFound = {
                                  fullname: data.fullname,
                                  username: data.username,
                                  configText: data.configText
                                };
                              }
                            } catch (lookupErr) {
                              console.error('Error lookup di catch:', lookupErr);
                            }

                            if (centralUserFound) {
                              IndonesianError = `Maaf, Username "${lookupUsername}" sudah terdaftar di database sistem pusat.\n\n• Nama Pengguna: ${centralUserFound.fullname || 'N/A'}\n• Akun (Username): ${centralUserFound.username || 'N/A'}\n\nSilakan gunakan menu "Masuk" (Login) dan gunakan akun tersebut beserta kata sandinya untuk login.`;
                            } else {
                              IndonesianError = `Username "${lookupUsername}" sudah pernah didaftarkan pada database pusat, namun kata sandi yang Anda ketik salah.\n\nLangkah Solusi:\n1. Masukkan kata sandi yang tepat jika Anda adalah pemilik akun tersebut.\n2. ATAU, silakan daftar dengan memakai Username yang berbeda.`;
                            }
                          } else if (err.message.includes('weak-password')) {
                            IndonesianError = 'Kata sandi minimal berisi 6 karakter.';
                          } else if (err.message.includes('invalid-api-key') || err.message.includes('API key')) {
                            IndonesianError = 'API Key yang terdapat pada konfigurasi Web Firebase Anda salah atau tidak valid.';
                          } else if (err.message.includes('network-request-failed')) {
                            IndonesianError = 'Koneksi jaringan gagal. Periksa koneksi internet Anda atau rules Firebase.';
                          } else if (err.message.includes('operation-not-allowed')) {
                            IndonesianError = 'Provider "Email/Password" belum aktif di Firebase Console Anda!';
                          } else if (err.message.includes('configuration-not-found') || err.message.includes('auth/configuration-not-found')) {
                            IndonesianError = 'Layanan Autentikasi belum diinisialisasi di Proyek Firebase Anda!';
                          }

                          setRegistrationResult({
                            success: false,
                            title: 'Pendaftaran Gagal!',
                            message: `${IndonesianError}\n\n(Detail Teknis: ${err.message})`
                          });
                        } finally {
                          setIsAuthLoading(false);
                        }
                      }}
                      className="flex-1 px-5 py-3 rounded-xl font-bold text-white bg-[#098f41] hover:bg-[#077a37] disabled:bg-slate-300 disabled:cursor-not-allowed text-sm transition-all text-center"
                    >
                      Konfigurasi Database & Daftar Akun
                    </button>
                  </div>
                  {/* Spacer bottom to handle mobile browsers and overlays */}
                  <div className="h-4 sm:h-6 shrink-0" />
                </motion.div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Kotak Dialog Hasil Pendaftaran Akun / Firebase */}
          <AnimatePresence>
            {registrationResult && (
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 className="fixed inset-0 bg-[#020617]/80 backdrop-blur-md z-[200] flex items-center justify-center p-4 overflow-y-auto"
              >
                <motion.div 
                   initial={{ opacity: 0, scale: 0.95, y: 15 }}
                   animate={{ opacity: 1, scale: 1, y: 0 }}
                   exit={{ opacity: 0, scale: 0.95, y: 15 }}
                   className="bg-white rounded-3xl p-6 sm:p-8 shadow-2xl max-w-md w-full space-y-5 text-center relative max-h-[85vh] sm:max-h-[90vh] overflow-y-auto"
                >
                  {/* Decorative status header background with matching rounded top corners */}
                  <div className={`absolute top-0 left-0 right-0 h-2 rounded-t-3xl ${
                    registrationResult.success ? 'bg-[#098f41]' : 'bg-rose-500'
                  }`} />

                  {/* Icon indicator */}
                  <div className="flex justify-center pt-3">
                    {registrationResult.success ? (
                      <div className="w-16 h-16 bg-[#f7fee7] border-2 border-[#bbf7d0] rounded-full flex items-center justify-center text-[#098f41]">
                        <svg className="w-8 h-8 font-black" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7" />
                        </svg>
                      </div>
                    ) : (
                      <div className="w-16 h-16 bg-rose-50 border-2 border-rose-200 rounded-full flex items-center justify-center text-rose-500">
                        <svg className="w-8 h-8 font-black" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M6 18L18 6M6 6l12 12" />
                        </svg>
                      </div>
                    )}
                  </div>

                  <div className="space-y-2">
                     <h3 className={`text-lg font-black tracking-tight ${
                       registrationResult.success ? 'text-slate-800' : 'text-rose-600'
                     }`}>
                       {registrationResult.title}
                     </h3>
                     <p className="text-xs text-slate-600 leading-relaxed px-2">
                       {registrationResult.message}
                     </p>
                  </div>

                  {registrationResult.success && (
                    <div className="bg-slate-50 border border-slate-100 rounded-2xl p-4 text-left space-y-2 text-xs">
                      <div>
                        <span className="text-[9px] font-bold text-slate-500 block uppercase">NAMA LENGKAP UTAMA</span>
                        <span className="font-bold text-slate-700">{registrationResult.fullname}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-200/60">
                        <div>
                          <span className="text-[9px] font-bold text-slate-500 block uppercase">USERNAME LOGIN</span>
                          <span className="font-mono font-bold text-[#098f41] break-all">{registrationResult.username}</span>
                        </div>
                        <div>
                          <span className="text-[9px] font-bold text-slate-500 block uppercase">PROJECT ID</span>
                          <span className="font-mono font-bold text-slate-700 break-all">{registrationResult.projectId}</span>
                        </div>
                      </div>
                    </div>
                  )}

                  {!registrationResult.success && (
                    <div className="bg-rose-50/50 border border-rose-100 rounded-2xl p-3 text-left text-xs text-rose-700 leading-relaxed">
                      <strong className="block mb-1">Rekomendasi Solusi:</strong>
                      <ul className="list-disc pl-4 mt-1 space-y-1 text-[11px]">
                        <li>Verifikasi kebenaran dan format Config Firebase yang dipaste.</li>
                        <li>Pastikan database auth di Firebase Console sudah diaktifkan (Email/Password).</li>
                        <li>Pastikan aturan Firestore rules Anda tidak memblokir operasi pembuatan user.</li>
                      </ul>
                    </div>
                  )}

                  {!registrationResult.success && registrationResult.showBypassButton && registrationResult.configToSave && (
                    <button 
                       type="button" 
                       onClick={() => handleForceRegister(registrationResult.configToSave!)}
                       className="w-full py-3 rounded-xl font-bold text-white bg-[#098f41] hover:bg-[#077a37] transition-all text-sm shadow-md active:scale-[0.98] mt-2 block"
                    >
                      Hubungkan & Tetap Daftar Akun Baru
                    </button>
                  )}

                  <button 
                     type="button" 
                     onClick={() => setRegistrationResult(null)}
                     className={`w-full py-3 rounded-xl font-bold text-white transition-all text-sm shadow-md active:scale-[0.98] ${
                       registrationResult.success 
                         ? 'bg-[#098f41] hover:bg-[#077a37]' 
                         : 'bg-slate-800 hover:bg-slate-900 border border-slate-700'
                     }`}
                  >
                    {registrationResult.success ? 'Selesai & Masuk Sekarang' : registrationResult.showBypassButton ? 'Batal' : 'Tutup Dialog'}
                  </button>
                </motion.div>
              </motion.div>
            )}

            {/* Kotak Dialog Alert Modern - Login Gagal / Salah Password / Belum Terdaftar */}
            {loginError && (
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 className="fixed inset-0 bg-[#020617]/80 backdrop-blur-md z-[210] flex items-center justify-center p-4 overflow-y-auto"
              >
                <motion.div 
                   initial={{ opacity: 0, scale: 0.95, y: 15 }}
                   animate={{ opacity: 1, scale: 1, y: 0 }}
                   exit={{ opacity: 0, scale: 0.95, y: 15 }}
                   className="bg-white rounded-3xl p-6 sm:p-8 shadow-2xl max-w-md w-full space-y-5 text-center relative max-h-[85vh] sm:max-h-[90vh] overflow-y-auto"
                >
                  {/* Decorative error status line at top */}
                  <div className="absolute top-0 left-0 right-0 h-2 bg-rose-500 rounded-t-3xl" />

                  {/* Icon indicator with warnings */}
                  <div className="flex justify-center pt-3">
                    <div className="w-16 h-16 bg-rose-50 border-2 border-rose-200 rounded-full flex items-center justify-center text-rose-500 shadow-sm animate-pulse">
                      <svg className="w-8 h-8 font-black" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                      </svg>
                    </div>
                  </div>

                  <div className="space-y-2">
                     <h3 className="text-lg font-black tracking-tight text-rose-600">
                       {loginError.title}
                     </h3>
                     <p className="text-xs text-slate-600 leading-relaxed px-2">
                       {loginError.message}
                     </p>
                  </div>

                  {/* List of custom helpful step-by-step procedures */}
                  <div className="bg-rose-50/50 border border-rose-100/60 rounded-2xl p-4 text-left text-xs text-rose-700 leading-relaxed">
                    <strong className="block mb-2 font-black text-rose-800 text-[11px] uppercase tracking-wider">Rekomendasi Solusi & Navigasi:</strong>
                    <ul className="list-disc pl-4 space-y-2 text-[11px] text-slate-600">
                      {loginError.recommendations.map((rec, idx) => (
                        <li key={idx} className="leading-snug">
                          {rec}
                        </li>
                      ))}
                    </ul>
                  </div>

                  <button 
                     type="button" 
                     onClick={() => setLoginError(null)}
                     className="w-full py-3 rounded-xl font-bold text-white bg-slate-800 hover:bg-slate-900 border border-slate-700 transition-all text-sm shadow-md active:scale-[0.98]"
                  >
                    Tutup Pesan
                  </button>
                </motion.div>
              </motion.div>
            )}


          </AnimatePresence>

        {/* Global Footer */}
        <div className="mt-auto text-center pb-8 pt-8 z-10 w-full px-4 flex flex-col items-center">
           <div className="flex items-center justify-center gap-2 mb-2">
              <span className="w-2 h-2 rounded-full bg-[#098f41]"></span>
              <p className="text-sm font-black text-[#098f41]">App Development by Agan Parta,S.Kom.,Gr</p>
           </div>
           <p className="text-[8px] md:text-[9px] font-bold text-slate-500 tracking-[0.15em] sm:tracking-[0.25em] uppercase mb-2">
             KREATIVITAS TANPA BATAS • INOVASI TIADA HENTI
           </p>
           <p className="text-[9px] md:text-[10px] italic text-slate-500 text-center max-w-xs md:max-w-none">
             "Transformasi Digital Pendidikan Untuk Generasi Emas yang Cerdas dan Berakhlak"
           </p>
           
           <div className="flex justify-center items-center gap-6 md:gap-8 mt-10 opacity-70">
             <span className="text-[9px] md:text-[10px] font-bold text-slate-500">V2.1.0</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-200">|</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Enterprise</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-200">|</span>
             <span className="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Stable</span>
           </div>
        </div>
      </div>
    );
  }

  return (
    <div className="h-[100dvh] bg-slate-50 flex flex-col overflow-hidden print:h-auto print:overflow-visible selection:bg-[#098f41]/20 selection:text-[#077a37]">
      <GlobalConnectivityBanner />
      {splashNode}

      {/* Persistence / Connectivity Banners - OLD REMOVED */}

      {/* Header */}
      <header className="bg-white px-6 py-3.5 flex justify-between items-center border-b border-slate-100 z-40 shadow-[0_2px_15px_rgba(148,163,184,0.03)] shrink-0 print:hidden">
        <div className="flex items-center gap-3">
          <img 
            src="/school_logo.png" 
            alt="Logo" 
            className="w-10 h-10 object-contain drop-shadow-sm" 
            onError={(e) => {
              (e.target as HTMLImageElement).src = "https://drive.google.com/thumbnail?id=1c0ibueBZudROdPwR1oJKDC4y0HnJ1R4n&sz=w1000";
            }}
            referrerPolicy="no-referrer" 
          />
          <div>
            <h1 className="text-lg font-black text-[#098f41] tracking-tight leading-none scale-y-105 origin-left whitespace-nowrap">SMART DF App</h1>
            <p className="text-[10px] text-slate-500 font-bold tracking-wider uppercase mt-1 hidden sm:block whitespace-nowrap">SMA DARUL FALAH CIHAMPELAS</p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          {/* Status Sinkronisasi Real-time Database */}
          <div className={`hidden md:flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold border transition-all ${
            !isOnline || !firebaseConnected || syncStatus === 'error'
              ? 'bg-rose-50 border-rose-100 text-rose-700 animate-pulse'
              : quotaExceeded
                ? 'bg-amber-50 border-amber-100 text-amber-700'
                : 'bg-emerald-50 border-emerald-100 text-[#077a37]'
          }`} title={!isOnline ? 'Koneksi offline, data saat ini disimpan lokal di peramban' : (!firebaseConnected || syncStatus === 'error') ? 'Gagal sinkronisasi dengan database cloud' : quotaExceeded ? 'Kuota harian Firestore habis. Sinkronisasi cloud dijeda.' : `Koneksi Cloud stabil, sinkronisasi aktif otomatis${cloudLastSync ? ' (Terakhir: ' + cloudLastSync + ')' : ''}`}>
            <span className={`w-2 h-2 rounded-full ${!isOnline || !firebaseConnected || syncStatus === 'error' ? 'bg-rose-500' : quotaExceeded ? 'bg-amber-500' : 'bg-[#098f41] animate-pulse'}`}></span>
            <div className="flex flex-col items-start leading-none gap-0.5">
               <span className="hidden xs:inline text-[10px] sm:text-xs">{!isOnline ? 'Internet Putus' : (!firebaseConnected || syncStatus === 'error') ? 'Cloud Gagal' : quotaExceeded ? 'Limit Tercapai' : 'Cloud Terhubung'}</span>
               <span className="xs:hidden text-[10px]">{!isOnline ? 'Offline' : (!firebaseConnected || syncStatus === 'error') ? 'Gagal' : quotaExceeded ? 'Limit' : 'Online'}</span>
               {cloudLastSync && isOnline && firebaseConnected && syncStatus !== 'error' && !quotaExceeded && (
                 <span className="text-[8px] opacity-70 hidden sm:inline">Sync {cloudLastSync}</span>
               )}
            </div>
          </div>

          <div className="flex flex-col items-center gap-0.5 ml-1">
            <button 
              onClick={() => setActiveTab('profile')} 
              className="relative w-10 h-10 rounded-full bg-slate-50 hover:bg-slate-100 text-slate-700 transition-all shadow-sm cursor-pointer flex items-center justify-center border border-slate-200 hover:scale-105 active:scale-95 overflow-hidden p-0"
              title="Menu Profil Pengguna"
            >
              {avatarUrl ? (
                <img src={avatarUrl} alt="Avatar" className="w-full h-full object-cover rounded-full" />
              ) : (
                <UserIcon className="w-5 h-5 text-slate-600" />
              )}
              <span className={`absolute -top-0.5 -right-0.5 w-2.5 h-2.5 border-2 border-white rounded-full ${isOnline ? 'bg-[#098f41] animate-pulse' : 'bg-rose-500'}`}></span>
            </button>
            <span className="text-[9px] font-black text-slate-600 uppercase tracking-wider leading-none">Profil</span>
          </div>

          <div className="flex flex-col items-center gap-0.5 ml-1">
            <button 
              onClick={() => setShowLogoutConfirm(true)} 
              className="w-10 h-10 rounded-full bg-rose-50 hover:bg-rose-100 text-rose-600 hover:text-rose-700 transition-all shadow-sm cursor-pointer flex items-center justify-center border border-rose-100/50 hover:scale-105 active:scale-95"
              title="Keluar dari Aplikasi"
            >
              <Power className="w-5 h-5 stroke-[2.5]" />
            </button>
            <span className="text-[9px] font-black text-rose-600 uppercase tracking-wider leading-none">Keluar</span>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 bg-slate-50 overflow-y-auto print:overflow-visible print:h-auto">
            <AnimatePresence mode="wait">
              <motion.div
                key={activeTab}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                transition={{ duration: 0.3, ease: "easeInOut" }}
                className="h-full relative print:h-auto print:overflow-visible"
              >
                {renderContent()}
              </motion.div>
            </AnimatePresence>

        {/* Toast Notification */}
        <AnimatePresence>
          {toast && (
            <motion.div
              initial={{ opacity: 0, y: 20, scale: 0.95 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 20, scale: 0.95 }}
              className="fixed inset-0 flex items-center justify-center p-4 z-50 pointer-events-none"
            >
                <div className={`px-6 py-4 rounded-2xl shadow-2xl flex items-center gap-3 border pointer-events-auto ${
                toast.type === 'success' ? 'bg-[#098f41] text-white border-[#098f41]' :
                toast.type === 'info' ? 'bg-indigo-500 text-white border-indigo-400' :
                'bg-rose-500 text-white border-rose-400'
              }`}>
                  {toast.type === 'success' && <Check className="w-5 h-5" />}
                  {toast.type === 'info' && <Info className="w-5 h-5" />}
                  {toast.type === 'error' && <AlertCircle className="w-5 h-5" />}
                  <span className="font-bold">{toast.message}</span>
                  <button onClick={() => { if(toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current); setToast(null); }} className="ml-2 hover:opacity-80">
                    <X className="w-4 h-4" />
                  </button>
                </div>
            </motion.div>
          )}
        </AnimatePresence>
        
        {/* Confirmation Modal */}
        <AnimatePresence>
          {confirmationAction && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-slate-900/60 backdrop-blur-xl z-[100] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.9, y: 30 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 30 }}
                transition={{ type: "spring", damping: 25, stiffness: 300 }}
                className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_40px_100px_-20px_rgba(0,0,0,0.3)] text-center relative overflow-hidden group border border-slate-100"
              >
                  {/* Decorative background elements */}
                  <div className="absolute top-0 left-0 w-full h-2 bg-rose-500"></div>
                  <div className="absolute -top-12 -right-12 w-32 h-32 bg-rose-50 rounded-full opacity-50 group-hover:scale-110 transition-transform duration-700"></div>

                  <div className="relative z-10">
                    <div className="mx-auto w-24 h-24 bg-rose-100 rounded-[2rem] flex items-center justify-center mb-8 rotate-3 group-hover:rotate-6 transition-transform">
                       <AlertTriangle className="w-12 h-12 text-rose-600" />
                    </div>
                    
                    <h3 className="text-2xl font-black text-slate-900 mb-3 tracking-tight">Konfirmasi Hapus</h3>
                    <p className="text-slate-600 text-sm mb-10 leading-relaxed font-medium">
                      Apakah Anda yakin ingin menghapus siswa <strong className="text-rose-600 font-bold">"{confirmationAction.student.name}"</strong>? Data absensi terkait mungkin juga akan terpengaruh.
                    </p>

                    <div className="flex gap-3">
                      <button 
                        onClick={() => setConfirmationAction(null)}
                        className="flex-1 bg-slate-100 text-slate-600 py-4 rounded-2xl font-black text-sm hover:bg-slate-200 transition-all active:scale-95"
                      >
                        Batal
                      </button>
                      <button 
                        onClick={confirmAction}
                        className="flex-1 bg-rose-600 text-white py-4 rounded-2xl font-black text-sm hover:bg-rose-700 hover:shadow-xl hover:shadow-rose-100 transition-all active:scale-95"
                      >
                        Ya, Hapus
                      </button>
                    </div>
                  </div>
                </motion.div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Teacher Delete Confirmation Modal */}
          <AnimatePresence>
            {teacherConfirmationAction && teacherConfirmationAction.type === 'delete' && (
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4"
              >
                <motion.div 
                  initial={{ opacity: 0, scale: 0.9, y: 30 }}
                  animate={{ opacity: 1, scale: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.9, y: 30 }}
                  transition={{ type: "spring", damping: 25, stiffness: 300 }}
                  className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_40px_100px_-20px_rgba(0,0,0,0.3)] text-center relative overflow-hidden group border border-slate-100"
                >
                  <div className="absolute top-0 left-0 w-full h-2 bg-rose-500"></div>
                  <div className="absolute -top-12 -right-12 w-32 h-32 bg-rose-50 rounded-full opacity-50 group-hover:scale-110 transition-transform duration-700"></div>

                  <div className="relative z-10">
                    <div className="mx-auto w-24 h-24 bg-rose-100 rounded-[2rem] flex items-center justify-center mb-8 rotate-3 group-hover:rotate-6 transition-transform">
                       <AlertTriangle className="w-12 h-12 text-rose-600" />
                    </div>
                    
                    <h3 className="text-2xl font-black text-slate-900 mb-3 tracking-tight">Konfirmasi Hapus</h3>
                    <p className="text-slate-600 text-sm mb-10 leading-relaxed font-medium">
                      Apakah Anda yakin ingin menghapus guru <strong className="text-rose-600 font-bold">"{teacherConfirmationAction.teacher.name}"</strong>?
                    </p>

                    <div className="flex gap-3">
                      <button 
                        onClick={() => setTeacherConfirmationAction(null)}
                        className="flex-1 bg-slate-100 text-slate-600 py-4 rounded-2xl font-black text-sm hover:bg-slate-200 transition-all active:scale-95"
                      >
                        Batal
                      </button>
                      <button 
                        onClick={confirmTeacherAction}
                        className="flex-1 bg-rose-600 text-white py-4 rounded-2xl font-black text-sm hover:bg-rose-700 hover:shadow-xl hover:shadow-rose-100 transition-all active:scale-95"
                      >
                        Ya, Hapus
                      </button>
                    </div>
                  </div>
                </motion.div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Reset Database Confirmation Modal */}
        <AnimatePresence>
          {resetModalType !== 'none' && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-md w-full space-y-6 border border-slate-100 text-left"
              >
                <div className="flex items-start gap-4">
                  <div className={`p-3 rounded-2xl ${resetModalType === 'new_semester' ? 'bg-emerald-50 text-[#098f41]' : resetModalType === 'new_year' ? 'bg-amber-50 text-amber-600' : 'bg-rose-50 text-rose-600'}`}>
                    {resetModalType === 'new_semester' ? <RefreshCw className="w-6 h-6" /> : <AlertCircle className="w-6 h-6" />}
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-slate-800">
                      {resetModalType === 'new_semester' ? 'Reset Semester Baru' : resetModalType === 'new_year' ? 'Reset Tahun Ajaran Baru' : resetModalType === 'clear_all_students' ? 'Kosongkan Semua Siswa' : resetModalType === 'clear_all_teachers' ? 'Kosongkan Semua Guru' : 'Reset Semua Data (Total)'}
                    </h3>
                    <p className={`text-xs font-bold mt-0.5 tracking-wider uppercase ${resetModalType === 'new_semester' ? 'text-emerald-700' : 'text-slate-500'}`}>
                      {resetModalType === 'new_semester' ? '🔒 Konfirmasi Sandi Akun' : '⚠️ TINDAKAN TIDAK DAPAT DIBATALKAN'}
                    </p>
                  </div>
                </div>

                <div className="space-y-3 text-sm text-slate-600 leading-relaxed bg-slate-50 border border-slate-100 p-4 rounded-2xl">
                  {resetModalType === 'new_semester' ? (
                    <>
                      <p>
                        Anda akan mereset data sesi absensi untuk menyambut <strong>Semester Baru</strong>.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 font-sans">
                        <li>Semua <strong>{attendanceSessions.length} sesi absensi</strong> semester sebelumnya akan direset bersih.</li>
                        <li>Master biodata siswa <strong>({students.length} orang)</strong> &amp; daftar kelas tetap aman terlindungi.</li>
                        <li>Status semester pada profil Anda akan diperbarui ke semester baru yang dipilih.</li>
                      </ul>
                    </>
                  ) : resetModalType === 'new_year' ? (
                    <>
                      <p>
                        Anda akan mereset database untuk menyambut <strong>Tahun Ajaran Baru</strong>.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 font-sans">
                        <li>Semua <strong>{attendanceSessions.length} sesi absensi</strong> akan dihapus permanen.</li>
                        <li>Data biodata siswa <strong>({students.length} orang)</strong> &amp; daftar kelas tetap aman.</li>
                        <li>Statistik kehadiran siswa akan kembali kosong (0%).</li>
                      </ul>
                    </>
                  ) : resetModalType === 'clear_all_students' ? (
                    <>
                      <p>
                        Anda akan <strong>menghapus seluruh data siswa</strong> yang terdaftar.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 mb-2 font-sans">
                        <li>Semua biodata siswa <strong>({students.length} siswa)</strong> akan terhapus.</li>
                        <li>Daftar kelas dan riwayat absensi <strong>tetap dipertahankan</strong>.</li>
                      </ul>
                    </>
                  ) : resetModalType === 'clear_all_teachers' ? (
                    <>
                      <p>
                        Anda akan <strong>menghapus seluruh data guru</strong> yang terdaftar.
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 mb-2 font-sans">
                        <li>Semua biodata guru <strong>({teachers.length} guru)</strong> akan terhapus.</li>
                        <li>Data siswa, kelas, dan riwayat absensi <strong>tetap dipertahankan</strong>.</li>
                      </ul>
                    </>
                  ) : (
                    <>
                      <p>
                        Anda akan melakukan <strong>Reset Total (Hapus Semua)</strong>. Tindakan ini akan menghapus:
                      </p>
                      <ul className="list-disc list-inside space-y-1 text-slate-600 text-xs mt-2 pl-1 mb-2 font-sans">
                        <li>Semua biodata siswa <strong>({students.length} siswa)</strong>.</li>
                        <li>Semua biodata guru <strong>({teachers.length} guru)</strong>.</li>
                        <li>Semua riwayat &amp; <strong>{attendanceSessions.length} sesi absensi</strong> di cloud.</li>
                        <li>Seluruh daftar kelas yang terdaftar.</li>
                        <li>Biodata profil sekolah (Nama Guru, Mata Pelajaran, dsb).</li>
                      </ul>
                      <p className="text-xs text-rose-500 font-bold font-sans italic border-l-2 border-rose-300 pl-2">
                        Data Anda di cloud maupun cache lokal akan hilang seutuhnya dan tidak dapat dikembalikan!
                      </p>
                    </>
                  )}
                </div>

                {resetModalType === 'new_semester' && (
                  <div className="space-y-4 pt-1">
                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider block">Pilih Semester Baru</label>
                        <select 
                          value={newSemesterChoice} 
                          onChange={e => setNewSemesterChoice(e.target.value as 'Ganjil' | 'Genap')}
                          className="w-full px-3 py-2.5 bg-stone-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 text-xs font-bold text-slate-800"
                        >
                          <option value="Ganjil">Semester Ganjil</option>
                          <option value="Genap">Semester Genap</option>
                        </select>
                      </div>

                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-600 uppercase tracking-wider block">Tahun Pelajaran</label>
                        <input 
                          type="text" 
                          value={newTahunPelajaran} 
                          onChange={e => setNewTahunPelajaran(e.target.value)}
                          placeholder="Contoh: 2026/2027"
                          className="w-full px-3 py-2.5 bg-stone-50 border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 text-xs font-bold text-slate-800"
                        />
                      </div>
                    </div>
                  </div>
                )}

                <div className="space-y-1.5 pt-2">
                  <label className="text-[10px] font-black text-slate-700 uppercase tracking-wider block">
                    Masukkan Password Admin:
                  </label>
                  <div className="relative">
                    <input 
                      type={showResetPassword ? "text" : "password"} 
                      disabled={isResettingData || isVerifyingPassword}
                      value={resetPasswordInput}
                      onChange={e => {
                        setResetPasswordInput(e.target.value);
                        if (resetPasswordError) setResetPasswordError('');
                      }}
                      placeholder="Masukkan password admin default"
                      className="w-full px-4 py-3 pr-11 bg-white border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-bold text-slate-800 disabled:opacity-50"
                    />
                    <button
                      type="button"
                      onClick={() => setShowResetPassword(!showResetPassword)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
                    >
                      {showResetPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                    </button>
                  </div>
                  {resetPasswordError && (
                    <p className="text-xs font-bold text-rose-600 flex items-center gap-1.5 mt-1.5 bg-rose-50 p-2 rounded-lg border border-rose-200">
                      <AlertCircle className="w-4 h-4 shrink-0" />
                      {resetPasswordError}
                    </p>
                  )}
                </div>

                {isResettingData && (
                  <div className="space-y-2 py-2 bg-slate-50 p-4 rounded-2xl border-2 border-slate-300/50">
                    <div className="flex justify-between items-center text-[10px] font-black tracking-wider text-slate-600 uppercase font-sans">
                      <span>Proses Reset Berlangsung...</span>
                      <span className="font-mono text-slate-800 text-xs font-bold bg-white px-2 py-0.5 rounded-full border border-slate-100">{resetProgress}%</span>
                    </div>
                    <div className="w-full bg-slate-200 rounded-full h-3 overflow-hidden p-0.5 border-2 border-slate-400/30">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${resetProgress}%` }}
                        transition={{ duration: 0.1 }}
                        className={`h-full rounded-full ${
                          resetModalType === 'new_semester'
                            ? 'bg-[#098f41] shadow-[0_0_8px_rgba(9,143,65,0.5)]'
                            : resetModalType === 'new_year' 
                            ? 'bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.5)]' 
                            : 'bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.5)]'
                        }`}
                      />
                    </div>
                  </div>
                )}

                <div className="flex gap-4 pt-2">
                  <button 
                    disabled={isResettingData || isVerifyingPassword}
                    onClick={() => {
                      setResetModalType('none');
                      setResetConfirmInput('');
                      setResetPasswordInput('');
                      setResetPasswordError('');
                    }} 
                    className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all text-xs text-center disabled:opacity-50 cursor-pointer"
                  >
                    Batal
                  </button>
                  <button 
                    disabled={isResettingData || isVerifyingPassword || !resetPasswordInput.trim()}
                    onClick={async () => {
                      if (!resetPasswordInput.trim()) {
                        setResetPasswordError('Silakan masukkan sandi konfirmasi.');
                        return;
                      }
                      setIsVerifyingPassword(true);
                      setResetPasswordError('');

                      const trimmedPass = resetPasswordInput.trim();
                      let isMatch = false;

                      // Get Admin password dynamically
                      let adminPass = '123456@#';
                      try {
                        const adminSnap = await getDoc(doc(dbDefault, 'custom_accounts', 'admin'));
                        if (adminSnap.exists() && adminSnap.data()?.password) {
                          adminPass = adminSnap.data().password.trim();
                        }
                      } catch (e) {
                        console.warn('Failed to load admin password, using default.', e);
                      }

                      const loggedInUser = (activeUserCustomData?.username || '').toLowerCase().trim();

                      if (loggedInUser === 'petugaspiket') {
                        // Petugas piket must use admin's password
                        isMatch = (trimmedPass === '123456@#' || trimmedPass === adminPass);
                      } else if (loggedInUser === 'admin') {
                        // Admin must use admin password
                        isMatch = (trimmedPass === '123456@#' || trimmedPass === adminPass);
                      } else {
                        // Any other registered account must use their own password
                        if (activeUserCustomData?.password && trimmedPass === activeUserCustomData.password.trim()) {
                          isMatch = true;
                        } else {
                          // Look up centrally
                          try {
                            const snap = await getDoc(doc(dbDefault, 'custom_accounts', loggedInUser));
                            if (snap.exists() && snap.data()?.password === trimmedPass) {
                              isMatch = true;
                            }
                          } catch (err) {
                            console.warn('Central account lookup error:', err);
                          }
                        }
                      }

                      setIsVerifyingPassword(false);

                      if (!isMatch) {
                        setResetPasswordError('Sandi konfirmasi salah. Harap masukkan sandi yang tepat.');
                        return;
                      }

                      await handleResetData(resetModalType);
                    }} 
                    className={`flex-1 px-5 py-3 rounded-xl font-bold text-white transition-all text-xs flex items-center justify-center gap-2 ${
                      resetModalType === 'new_semester'
                        ? 'bg-[#098f41] hover:bg-[#077a37] shadow-[0_4px_12px_rgba(9,143,65,0.3)]'
                        : resetModalType === 'new_year' 
                        ? 'bg-amber-500 hover:bg-amber-600 shadow-[0_4px_12px_rgba(245,158,11,0.3)]' 
                        : 'bg-rose-500 hover:bg-rose-600 shadow-[0_4px_12px_rgba(244,63,94,0.3)]'
                    } disabled:opacity-50 disabled:shadow-none cursor-pointer`}
                  >
                    {isResettingData || isVerifyingPassword ? (
                      <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                    ) : null}
                    {resetModalType === 'new_semester' ? 'Konfirmasi & Reset' : 'Ya, Reset Sekarang'}
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Reset Success Modal */}
        <AnimatePresence>
          {resetSuccessModal !== 'none' && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 15 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 15 }}
                className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-sm w-full space-y-6 border border-slate-100 text-center"
              >
                <div className="w-16 h-16 bg-emerald-50 text-[#098f41] rounded-full mx-auto flex items-center justify-center mb-2 animate-bounce">
                  <CheckCircle className="w-8 h-8" />
                </div>

                <div className="space-y-2">
                  <h3 className="text-xl font-black text-slate-800">
                    {resetSuccessModal === 'new_semester' ? 'Semester Baru Siap!' : resetSuccessModal === 'new_year' ? 'Tahun Ajaran Baru Siap!' : resetSuccessModal === 'clear_all_students' ? 'Kosongkan Semua Siswa Berhasil' : resetSuccessModal === 'clear_all_teachers' ? 'Kosongkan Semua Guru Berhasil' : 'Reset Berhasil Dilakukan'}
                  </h3>
                  <p className="text-xs text-slate-600 leading-relaxed font-sans">
                    {resetSuccessModal === 'new_semester'
                      ? 'Seluruh riwayat dan sesi absensi semester lalu telah direset bersih. Biodata siswa & daftar kelas Anda tetap aman, serta semester baru telah aktif!'
                      : resetSuccessModal === 'new_year' 
                      ? 'Seluruh riwayat, kehadiran, dan sesi absensi telah dibersihkan sepenuhnya dari cloud. Daftar siswa dan kelas Anda tetap terawat dengan aman dan siap kembali digunakan!' 
                      : resetSuccessModal === 'clear_all_students'
                      ? 'Seluruh data siswa berhasil dihapus secara permanen. Daftar kelas dan riwayat absensi tidak ikut terhapus.'
                      : resetSuccessModal === 'clear_all_teachers'
                      ? 'Seluruh data guru berhasil dihapus secara permanen. Data siswa, kelas, dan riwayat absensi tidak ikut terhapus.'
                      : 'Seluruh database Anda (daftar siswa, daftar guru, riwayat absensi, sesi absensi, daftar kelas, dan profil sekolah) telah berhasil dihapus seutuhnya secara permanen dari server cloud.'}
                  </p>
                </div>

                <button 
                  onClick={() => setResetSuccessModal('none')} 
                  className="w-full px-5 py-3 rounded-xl font-bold bg-[#098f41] hover:bg-[#077a37] text-white shadow-[0_4px_12px_rgba(5,150,105,0.3)] transition-all text-sm cursor-pointer"
                >
                  Selesai &amp; Kembali
                </button>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Student Success Modal */}
        <AnimatePresence>
          {studentSuccessModal !== 'none' && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-slate-900/60 backdrop-blur-xl z-[150] flex items-center justify-center p-4"
            >
              <motion.div 
                initial={{ opacity: 0, scale: 0.9, y: 30 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 30 }}
                transition={{ type: "spring", damping: 25, stiffness: 300 }}
                className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_45px_110px_-20px_rgba(0,0,0,0.4)] text-center relative overflow-hidden group border border-slate-100/50"
              >
                 {/* Decorative background elements */}
                 <div className={`absolute top-0 left-0 w-full h-2.5 bg-gradient-to-r ${
                    studentSuccessModal === 'added' ? 'from-emerald-400 to-teal-400' :
                    studentSuccessModal === 'edited' ? 'from-sky-400 to-indigo-400' :
                    'from-rose-400 to-orange-400'
                 }`}></div>
                 
                 <div className={`absolute -top-12 -right-12 w-36 h-36 rounded-full opacity-40 blur-2xl transition-transform duration-1000 group-hover:scale-125 ${
                    studentSuccessModal === 'added' ? 'bg-emerald-100' :
                    studentSuccessModal === 'edited' ? 'bg-sky-100' :
                    'bg-rose-100'
                 }`}></div>

                 <div className="relative z-10">
                    <div className={`mx-auto w-24 h-24 rounded-[2.5rem] flex items-center justify-center mb-9 rotate-3 group-hover:rotate-6 transition-all duration-500 shadow-lg ${
                        studentSuccessModal === 'added' ? 'bg-emerald-50 text-[#098f41] shadow-emerald-100' :
                        studentSuccessModal === 'edited' ? 'bg-sky-50 text-sky-600 shadow-sky-100' :
                        'bg-rose-50 text-rose-600 shadow-rose-100'
                    }`}>
                       {studentSuccessModal === 'added' && <Users className="w-11 h-11" />}
                       {studentSuccessModal === 'edited' && <Pencil className="w-11 h-11" />}
                       {studentSuccessModal === 'deleted' && <Trash2 className="w-11 h-11" />}
                    </div>
                    
                    <h3 className="text-2xl font-black text-slate-900 mb-4 tracking-tight leading-tight">
                      {studentSuccessModal === 'added' ? 'Siswa Berhasil Ditambahkan' : 
                       studentSuccessModal === 'edited' ? 'Siswa Berhasil Diedit' : 
                       'Siswa Berhasil Dihapus'}
                    </h3>
                    <p className="text-slate-600 text-[15px] mb-11 leading-relaxed font-medium">
                      {studentSuccessModal === 'added' ? 'Data siswa baru telah aman tersimpan di cloud database sekolah Anda.' : 
                       studentSuccessModal === 'edited' ? 'Profil siswa tersebut telah berhasil diperbarui dan diselaraskan ke sistem.' : 
                       'Siswa dan data terkait telah berhasil dihilangkan dari sistem sekolah secara permanen.'}
                    </p>

                    <button 
                      onClick={() => setStudentSuccessModal('none')}
                      className={`w-full py-5 rounded-[1.5rem] font-black text-base transition-all duration-300 shadow-xl active:scale-[0.97] ${
                        studentSuccessModal === 'added' ? 'bg-[#098f41] hover:bg-[#077a37] text-white shadow-emerald-200' :
                        studentSuccessModal === 'edited' ? 'bg-sky-600 hover:bg-sky-700 text-white shadow-sky-200' :
                        'bg-rose-600 hover:bg-rose-700 text-white shadow-rose-200'
                      }`}
                    >
                      Mengerti, Selesai
                    </button>
                 </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Modal Peringatan Hapus User - Custom Requested */}
        <AnimatePresence>
          {userToDelete && (
            <motion.div 
               initial={{ opacity: 0 }}
               animate={{ opacity: 1 }}
               exit={{ opacity: 0 }}
               className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[250] flex items-center justify-center p-4 animate-in fade-in duration-200"
            >
              <motion.div 
                 initial={{ opacity: 0, scale: 0.95, y: 15 }}
                 animate={{ opacity: 1, scale: 1, y: 0 }}
                 exit={{ opacity: 0, scale: 0.95, y: 15 }}
                 className="bg-white rounded-[2rem] p-6 sm:p-8 shadow-2xl max-w-md w-full space-y-6 text-center relative border border-slate-100"
              >
                <div className="absolute top-0 left-0 right-0 h-2 bg-rose-500 rounded-t-[2rem]" />

                <div className="flex justify-center pt-2">
                  <div className="w-16 h-16 bg-rose-50 border-2 border-rose-100 rounded-full flex items-center justify-center text-rose-600 shadow-sm">
                    <AlertTriangle className="w-8 h-8" />
                  </div>
                </div>

                <div className="space-y-2">
                   <h3 className="text-xl font-black tracking-tight text-slate-800">
                     Hapus Akun Pengajar?
                   </h3>
                   <p className="text-xs text-slate-600 leading-relaxed px-2">
                     Apakah Anda yakin ingin menghapus akun pengajar <b>"{userToDelete.fullname}"</b> (username: <span className="font-mono text-xs text-rose-600 font-bold">{userToDelete.username}</span>) secara permanen?
                   </p>
                </div>

                <div className="bg-rose-50 border border-rose-100 rounded-2xl p-4 text-left text-xs text-rose-700 leading-relaxed space-y-1.5">
                  <strong className="block font-black text-rose-800 uppercase tracking-wide text-[10px]">Dampak Utama:</strong>
                  <ul className="list-disc pl-4 space-y-1 text-slate-600 text-[11px]">
                    <li>Akun ini tidak akan bisa login lagi ke sistem absensi.</li>
                    <li>Data pengajar ini di sistem registrasi pusat FRESH akan dihapus selamanya.</li>
                  </ul>
                </div>

                <div className="flex flex-col sm:flex-row gap-3 pt-2">
                  <button 
                     type="button" 
                     onClick={() => setUserToDelete(null)}
                     disabled={isDeletingUser}
                     className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs rounded-xl font-bold transition-all active:scale-[0.98] cursor-pointer"
                   >
                     Batal
                   </button>
                   <button 
                      type="button" 
                      onClick={handleConfirmDeleteUser}
                      disabled={isDeletingUser}
                      className="flex-1 py-3 rounded-xl font-bold text-white bg-rose-600 hover:bg-rose-700 transition-all text-xs shadow-md shadow-rose-200 active:scale-[0.98] flex items-center justify-center gap-2 cursor-pointer"
                   >
                     {isDeletingUser ? (
                       <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                     ) : null}
                     Ya, Hapus Akun
                   </button>
                 </div>
               </motion.div>
             </motion.div>
           )}
        </AnimatePresence>

        {/* Modal Edit User - Custom Requested */}
        <AnimatePresence>
          {userToEdit && (
            <motion.div 
               initial={{ opacity: 0 }}
               animate={{ opacity: 1 }}
               exit={{ opacity: 0 }}
               className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[250] flex items-center justify-center p-4 animate-in fade-in duration-200"
            >
              <motion.div 
                 initial={{ opacity: 0, scale: 0.95, y: 15 }}
                 animate={{ opacity: 1, scale: 1, y: 0 }}
                 exit={{ opacity: 0, scale: 0.95, y: 15 }}
                 className="bg-white rounded-[2rem] p-6 sm:p-8 shadow-2xl max-w-sm w-full space-y-6 text-left relative border border-slate-100"
              >
                <div className="absolute top-0 left-0 right-0 h-2 bg-indigo-600 rounded-t-[2rem]" />

                <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                  <div className="w-11 h-11 bg-indigo-50 border-2 border-indigo-100 rounded-full flex items-center justify-center text-indigo-600 shadow-xs">
                    <Pencil className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-black tracking-tight text-slate-800">
                      Edit Akun Pengajar
                    </h3>
                    <p className="text-[10px] text-slate-500 font-medium">
                      Perbarui nama & sandi akun pengajar
                    </p>
                  </div>
                </div>

                <div className="space-y-4">
                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                      Username (ID Akun)
                    </label>
                    <input 
                       type="text" 
                       value={editUsername} 
                       onChange={(e) => setEditUsername(e.target.value)}
                       disabled={userToEdit.username.toLowerCase().trim() === 'admin' || userToEdit.username.toLowerCase().trim() === 'petugaspiket'} 
                       className={`w-full border-2 text-xs rounded-xl px-4 py-2.5 font-mono text-slate-800 transition-all ${
                         (userToEdit.username.toLowerCase().trim() === 'admin' || userToEdit.username.toLowerCase().trim() === 'petugaspiket')
                           ? 'bg-slate-100 border-slate-300 text-slate-500 cursor-not-allowed opacity-80'
                           : 'bg-slate-50 border-slate-300 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 outline-hidden font-bold'
                       }`}
                       placeholder="Masukkan username baru"
                       title={
                         (userToEdit.username.toLowerCase().trim() === 'admin' || userToEdit.username.toLowerCase().trim() === 'petugaspiket')
                           ? 'Username admin atau petugas piket tidak dapat diubah'
                           : 'Ubah username pengguna ini'
                       }
                    />
                    {(userToEdit.username.toLowerCase().trim() === 'admin' || userToEdit.username.toLowerCase().trim() === 'petugaspiket') ? (
                      <span className="text-[9px] text-slate-500 italic mt-1 block">Username akun bawaan sistem bersifat permanen.</span>
                    ) : (
                      <span className="text-[9px] text-slate-500 italic mt-1 block">Admin dapat mengubah nama pengguna (ID Akun) ini.</span>
                    )}
                  </div>

                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                      Nama Lengkap
                    </label>
                    <input 
                       type="text" 
                       value={editFullname} 
                       onChange={(e) => setEditFullname(e.target.value)}
                       placeholder="Contoh: Budi Santoso, S.Pd."
                       className="w-full bg-slate-50 border-2 border-slate-300 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 outline-hidden font-bold text-slate-800 text-xs rounded-xl px-4 py-2.5 transition-all text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                      Sandi Baru / Password
                    </label>
                    <input 
                       type="text" 
                       value={editPassword} 
                       onChange={(e) => setEditPassword(e.target.value)}
                       placeholder="Sandi minimal 4 karakter"
                       className="w-full bg-slate-50 border-2 border-slate-300 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 outline-hidden font-bold text-slate-800 text-xs rounded-xl px-4 py-2.5 transition-all font-mono text-slate-800"
                    />
                  </div>
                </div>

                <div className="flex flex-col sm:flex-row gap-2.5 pt-2">
                  <button 
                     type="button" 
                     onClick={() => setUserToEdit(null)}
                     disabled={isSavingUser}
                     className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs rounded-xl font-bold transition-all active:scale-[0.98] cursor-pointer text-center"
                  >
                    Batal
                  </button>
                  <button 
                     type="button" 
                     onClick={handleSaveEditedUser}
                     disabled={isSavingUser}
                     className="flex-1 py-2.5 rounded-xl font-bold text-white bg-indigo-600 hover:bg-indigo-700 transition-all text-xs shadow-md shadow-indigo-100 active:scale-[0.98] flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    {isSavingUser ? (
                      <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                    ) : null}
                    Simpan Perubahan
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Google Drive Import Modal */}
        <AnimatePresence>
          {isGdriveModalOpen && (
            <motion.div 
               initial={{ opacity: 0 }}
               animate={{ opacity: 1 }}
               exit={{ opacity: 0 }}
               className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
            >
               <motion.div 
                  initial={{ scale: 0.95, opacity: 0, y: 10 }}
                  animate={{ scale: 1, opacity: 1, y: 0 }}
                  exit={{ scale: 0.95, opacity: 0, y: 10 }}
                  className="bg-white rounded-3xl shadow-xl w-full max-w-lg overflow-hidden border border-slate-100 p-6 space-y-6"
               >
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-12 h-12 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center text-xl font-bold">
                        🔗
                      </div>
                      <div>
                        <h3 className="text-lg font-black text-slate-800 tracking-tight">Impor Data {gdriveTarget === 'teacher' ? 'Guru' : 'Siswa'} dari Google Drive</h3>
                        <p className="text-xs text-slate-500">Ambil file Excel atau Google Sheets secara langsung</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => setIsGdriveModalOpen(false)}
                      className="text-slate-400 hover:text-slate-600 font-bold p-2"
                    >
                      ✕
                    </button>
                  </div>

                  <div className="space-y-4">
                    <div className="bg-emerald-50/70 border border-emerald-100 p-4 rounded-2xl text-xs text-slate-700 leading-relaxed space-y-1.5">
                      <p className="font-bold text-[#098f41]">💡 Petunjuk Penggunaan:</p>
                      <ul className="list-disc pl-4 space-y-1 text-slate-600">
                        <li>Buka file Google Sheets atau file `.xlsx` di Google Drive Anda.</li>
                        <li>Pastikan pengaturan akses file disetel ke <strong>"Siapa saja yang memiliki link dapat melihat"</strong> (Public / Anyone with the link).</li>
                        <li>Salin (Copy) tautan/link tersebut dan tempel (Paste) di bawah ini.</li>
                      </ul>
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block">
                        Tautan / Link Google Drive / Google Sheets:
                      </label>
                      <input 
                        type="url" 
                        placeholder="https://docs.google.com/spreadsheets/d/.../edit" 
                        className="w-full p-3.5 border-2 border-slate-200 rounded-2xl text-sm focus:border-[#098f41] focus:outline-none"
                        value={gdriveUrl}
                        onChange={(e) => setGdriveUrl(e.target.value)}
                      />
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-3 pt-2">
                    <button 
                      onClick={() => setIsGdriveModalOpen(false)}
                      className="px-5 py-3 rounded-2xl text-sm font-bold text-slate-600 hover:bg-slate-100 transition-colors"
                      disabled={isGdriveImporting}
                    >
                      Batal
                    </button>
                    <button 
                      onClick={handleImportFromGdrive}
                      disabled={isGdriveImporting}
                      className="bg-[#098f41] hover:bg-[#077a37] text-white font-bold py-3 px-6 rounded-2xl text-sm shadow-sm transition-all flex items-center gap-2 disabled:opacity-50"
                    >
                      {isGdriveImporting ? (
                        <>
                          <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                          <span>Mengunduh...</span>
                        </>
                      ) : (
                        <>
                          <span>📥 Ambil & Proses Data</span>
                        </>
                      )}
                    </button>
                  </div>
               </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Import Result Modal */}
        <AnimatePresence>
          {importResult && importResult.isOpen && (
             <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
             >
                <motion.div 
                   initial={{ scale: 0.95, opacity: 0, y: 10 }}
                   animate={{ scale: 1, opacity: 1, y: 0 }}
                   exit={{ scale: 0.95, opacity: 0, y: 10 }}
                   className="bg-white rounded-3xl shadow-xl w-full max-w-md overflow-hidden border border-slate-100"
                >
                   {importResult.error ? (
                     // State error/gagal
                     <div className="p-6 text-center space-y-4">
                        <div className="w-16 h-16 bg-rose-50 text-rose-500 rounded-full mx-auto flex items-center justify-center mb-2">
                            <AlertCircle className="w-8 h-8" />
                        </div>
                        <h3 className="text-xl font-black text-rose-600 tracking-tight">Gagal Impor Excel</h3>
                        <p className="text-slate-600 text-xs font-medium leading-relaxed">
                          Tidak dapat mengimpor data siswa. Sistem menemukan masalah format atau database sebagai berikut:
                        </p>
                        
                        <div className="bg-rose-50 text-rose-700 rounded-2xl p-4 text-left border border-rose-100/60 text-xs font-bold leading-relaxed space-y-1 my-3">
                           <div className="text-[10px] font-black uppercase text-rose-800 tracking-wider">Pesan Detail Sistem:</div>
                           <div className="font-semibold text-slate-700 bg-white/70 rounded-lg p-2 border border-rose-200">
                             {importResult.errorMessage || 'Terjadi kesalahan tidak dikenal saat parsing file.'}
                           </div>
                        </div>

                        {importResult.details && importResult.details.length > 0 && (
                          <div className="text-left space-y-1.5 my-3">
                            <p className="text-[10px] font-black text-slate-600 uppercase tracking-wider">Riwayat Log / Catatan Pemeriksaan:</p>
                            <div className="max-h-24 overflow-y-auto rounded-xl bg-slate-50 border-2 border-slate-300/50 p-2 text-[11px] text-slate-600 font-medium space-y-1">
                              {importResult.details.map((detail, idx) => (
                                <div key={idx} className="flex gap-1 items-start">
                                  <span className="shrink-0 text-slate-500">•</span>
                                  <span>{detail}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        <button 
                           onClick={() => setImportResult(null)}
                           className="w-full bg-slate-800 hover:bg-slate-900 text-white font-bold py-3.5 rounded-2xl shadow-sm transition-all text-sm"
                        >
                           Tutup & Perbaiki File
                        </button>
                     </div>
                   ) : (
                     // State sukses/informasi lengkap
                     <div className="p-6 text-center space-y-4">
                        <div className="w-16 h-16 bg-[#098f41]/10 text-[#098f41] rounded-full mx-auto flex items-center justify-center mb-2">
                            <CheckCircle className="w-8 h-8" />
                        </div>
                        <h3 className="text-xl font-black text-slate-800 tracking-tight">Hasil Impor Selesai</h3>
                        <p className="text-slate-600 text-xs font-medium">Proses penarikan data selesai. Berikut adalah ringkasan pemeriksaan lembar file Anda:</p>
                        
                        <div className="bg-slate-50 rounded-2xl p-4 text-left border border-slate-100 flex flex-col gap-2.5 my-4">
                           <div className="flex justify-between items-center text-sm font-bold text-slate-700">
                             <span>✅ Berhasil Diimpor (Firebase):</span>
                             <span className="text-[#098f41] text-base font-black">{importResult.successCount} Siswa</span>
                           </div>
                           <div className="flex justify-between items-center text-sm font-bold text-rose-500">
                             <span>❌ Gagal/Nama Kosong:</span>
                             <span className="text-rose-500 text-base font-black">{importResult.failCount} Baris</span>
                           </div>
                           <div className="flex justify-between items-center text-sm font-semibold text-slate-500">
                             <span>⏩ Baris Kosong Dilewati:</span>
                             <span>{importResult.emptyCount} Baris</span>
                           </div>
                           <div className="border-t border-slate-200/60 pt-2 flex justify-between items-center text-xs text-slate-600 font-bold">
                             <span>Total Baris Diperiksa:</span>
                             <span>{importResult.totalParsed} baris</span>
                           </div>
                        </div>

                        {/* Tampilkan statistik lembar kerja (worksheet list) jika ada */}
                        {importResult.sheetsProcessed && importResult.sheetsProcessed.length > 0 && (
                          <div className="text-left space-y-1.5 my-3">
                            <p className="text-[10px] font-black text-slate-600 uppercase tracking-wider">Daftar Lembar Kerja (Sheets) Terbaca:</p>
                            <div className="max-h-24 overflow-y-auto rounded-xl border border-slate-100/80 bg-slate-50/50 p-2 divide-y divide-slate-100 text-xs">
                              {importResult.sheetsProcessed.map((sh, idx) => (
                                <div key={idx} className="flex justify-between py-1.5 px-1 font-bold text-slate-600 first:pt-0 last:pb-0">
                                  <span className="truncate pr-2">📄 {sh.name}</span>
                                  <span className="text-[#098f41] shrink-0 font-black">{sh.count} siswa</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        {importResult.details && importResult.details.length > 0 && (
                          <div className="text-left space-y-1.5 my-3">
                            <p className="text-[10px] font-black text-slate-600 uppercase tracking-wider font-sans">Log Baris Dilewati / Bermasalah:</p>
                            <div className="max-h-24 overflow-y-auto rounded-xl bg-rose-50/50 border border-rose-100/60 p-2 text-[11px] text-rose-700 font-medium space-y-1 font-sans">
                              {importResult.details.map((detail, idx) => (
                                <div key={idx} className="flex gap-1 items-start">
                                  <span className="shrink-0">•</span>
                                  <span>{detail}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        <button 
                           onClick={() => setImportResult(null)}
                           className="w-full bg-[#098f41] hover:bg-[#077a37] text-white font-bold py-3.5 rounded-2xl shadow-sm transition-all focus:outline-none focus:ring-2 focus:ring-[#098f41]/50 text-sm"
                        >
                           Tutup & Selesai
                        </button>
                     </div>
                   )}
                </motion.div>
             </motion.div>
          )}
        </AnimatePresence>

        {/* Logout Confirm Modal */}
        <AnimatePresence>
          {showLogoutConfirm && (
             <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
               <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 onClick={() => setShowLogoutConfirm(false)}
                 className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
               />
               <motion.div 
                 initial={{ scale: 0.95, opacity: 0, y: 10 }}
                 animate={{ scale: 1, opacity: 1, y: 0 }}
                 exit={{ scale: 0.95, opacity: 0, y: 10 }}
                 className="relative bg-white rounded-[2rem] p-8 max-w-sm w-full shadow-2xl text-center border-2 border-rose-100"
               >
                 <div className="w-16 h-16 bg-rose-100 text-rose-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-inner">
                   <LogOut className="w-8 h-8 ml-1" />
                 </div>
                 
                 <h3 className="text-xl font-black text-slate-800 mb-3 tracking-tight">Keluar Aplikasi?</h3>
                 <div className="text-slate-600 text-sm leading-relaxed mb-8 font-medium">
                   Apakah Anda yakin ingin keluar dari sesi ini? Menghentikan sesi sangat disarankan jika aplikasi sedang tidak digunakan untuk menghemat limit database.
                 </div>
                 
                 <div className="flex gap-3">
                   <button 
                     onClick={() => setShowLogoutConfirm(false)}
                     className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-sm rounded-xl font-bold transition-all active:scale-[0.98] cursor-pointer"
                   >
                     Batal
                   </button>
                   <button 
                     onClick={async () => {
                        setShowLogoutConfirm(false);
                        speakText("Anda Telah Keluar Dari Aplikasi");

                        if (activeUserCustomData?.username) {
                          localStorage.removeItem(`kaguci_profile_${activeUserCustomData.username.toLowerCase()}`);
                          localStorage.removeItem(`kaguci_avatar_${activeUserCustomData.username.toLowerCase()}`);
                          sessionStorage.removeItem(`kaguci_welcomed_${activeUserCustomData.username}`);
                        }
                        localStorage.removeItem('kaguci_active_custom_user');
                        localStorage.removeItem('kaguci_saved_credentials');
                        localStorage.removeItem('kaguci_has_logged_in');
                        safeSetLocalStorage('kaguci_isLoggedIn', 'false');
                        
                        try {
                          window.history.replaceState(null, '', window.location.pathname + window.location.search);
                        } catch {
                          window.location.hash = '';
                        }

                        setActiveUserCustomData(null);
                        setCurrentUser(null);
                        setAuthEmail('');
                        setAuthPassword('');
                        setIsLoggedIn(false);

                        try {
                          await signOut(activeAuth);
                        } catch (e) {
                          console.error("Firebase signOut error: ", e);
                        }
                        showToast('Berhasil keluar.', 'success');
                     }}
                     className="flex-1 py-3 rounded-xl font-bold text-white bg-rose-600 hover:bg-rose-700 transition-all text-sm shadow-md shadow-rose-200 active:scale-[0.98] flex items-center justify-center gap-2 cursor-pointer"
                   >
                     Keluar
                   </button>
                 </div>
               </motion.div>
             </div>
          )}
        </AnimatePresence>

        {/* Homeroom Role Warning Modal */}
        <AnimatePresence>
          {showHomeroomRoleAlert && (
             <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
               <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 onClick={() => setShowHomeroomRoleAlert(false)}
                 className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
               />
               <motion.div 
                 initial={{ scale: 0.92, opacity: 0, y: 15 }}
                 animate={{ scale: 1, opacity: 1, y: 0 }}
                 exit={{ scale: 0.92, opacity: 0, y: 15 }}
                 className="relative bg-white rounded-[2.5rem] p-6 sm:p-8 max-w-md w-full shadow-2xl text-center border-2 border-amber-100/90 overflow-hidden"
               >
                 {/* Top Close Button */}
                 <button 
                   onClick={() => setShowHomeroomRoleAlert(false)}
                   className="absolute top-4 right-4 p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors cursor-pointer"
                   title="Tutup"
                 >
                   <X className="w-5 h-5" />
                 </button>

                 {/* Decorative background glow */}
                 <div className="absolute -top-16 -right-16 w-36 h-36 bg-amber-100/60 rounded-full blur-2xl pointer-events-none" />
                 <div className="absolute -bottom-16 -left-16 w-36 h-36 bg-emerald-100/60 rounded-full blur-2xl pointer-events-none" />

                 {/* Icon Header */}
                 <div className="relative mx-auto mb-4 w-20 h-20 rounded-3xl bg-gradient-to-tr from-amber-500 to-amber-400 text-white flex items-center justify-center shadow-lg shadow-amber-500/25">
                   <GraduationCap className="w-10 h-10" />
                   <div className="absolute -top-1 -right-1 w-6 h-6 bg-white rounded-full flex items-center justify-center text-amber-600 shadow-sm border border-amber-200">
                     <AlertCircle className="w-4 h-4" />
                   </div>
                 </div>

                 {/* Badge & Title */}
                 <div className="inline-flex items-center gap-1.5 px-3.5 py-1 bg-amber-50 border border-amber-200/70 text-amber-800 rounded-full text-xs font-black uppercase tracking-wider mb-2.5">
                   <Sparkles className="w-3.5 h-3.5 text-amber-500" />
                   Akses Khusus Wali Kelas
                 </div>
                 
                 <h3 className="text-xl sm:text-2xl font-black text-slate-800 tracking-tight mb-2.5">
                   Fitur Laporan Wali Kelas
                 </h3>
                 
                 {/* Notice message */}
                 <p className="text-slate-600 text-sm font-medium leading-relaxed mb-5">
                   Fitur ini hanya untuk wali kelas. Rubah peran jadi wali kelas di menu <strong>Profil Pengguna</strong> agar dapat mengakses seluruh data dan laporan wali kelas.
                 </p>

                 {/* Role Status Box */}
                 <div className="bg-slate-50/90 border border-slate-200/80 rounded-2xl p-3.5 mb-6 text-left space-y-2">
                   <div className="flex items-center justify-between text-xs font-bold text-slate-500">
                     <span>Peran Anda Saat Ini:</span>
                     <span className="px-2.5 py-0.5 bg-slate-200 text-slate-700 rounded-lg font-black">
                       {profileData.role || 'Guru Mapel'}
                     </span>
                   </div>
                   <div className="flex items-center justify-between text-xs font-bold text-amber-800 bg-amber-100/70 rounded-xl p-2.5">
                     <span className="flex items-center gap-1.5">
                       <CheckCircle className="w-4 h-4 text-amber-600 shrink-0" />
                       Peran Yang Dibutuhkan:
                     </span>
                     <span className="font-black text-amber-900 uppercase">
                       Wali Kelas
                     </span>
                   </div>
                 </div>

                 {/* Action Buttons */}
                 <div className="flex flex-col gap-2.5">
                   <button 
                     onClick={() => {
                       setShowHomeroomRoleAlert(false);
                       setActiveTab('profile');
                       setIsProfileEditing(true);
                     }}
                     className="w-full py-3.5 rounded-2xl font-black text-white bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 transition-all text-sm shadow-md shadow-amber-500/25 active:scale-[0.98] flex items-center justify-center gap-2 cursor-pointer"
                   >
                     <Settings2 className="w-4 h-4" />
                     <span>Buka Pengaturan Profil</span>
                   </button>

                   <button 
                     onClick={() => setShowHomeroomRoleAlert(false)}
                     className="w-full py-2.5 text-slate-500 hover:text-slate-700 font-bold text-xs transition-colors cursor-pointer bg-slate-100 hover:bg-slate-200 rounded-xl"
                   >
                     Tutup
                   </button>
                 </div>
               </motion.div>
             </div>
          )}
        </AnimatePresence>

        {/* Quota Exceeded Professional Popup */}
        <AnimatePresence>
          {showQuotaModal && (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 sm:p-6">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowQuotaModal(false)}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-md"
              />
              
              <motion.div
                initial={{ scale: 0.9, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.9, opacity: 0, y: 20 }}
                className="relative w-full max-w-lg bg-white rounded-[2.5rem] p-8 shadow-2xl overflow-hidden border border-white"
              >
                {/* Abstract decorative elements */}
                <div className="absolute -top-24 -right-24 w-48 h-48 bg-amber-100 rounded-full blur-3xl opacity-50" />
                <div className="absolute -bottom-24 -left-24 w-48 h-48 bg-blue-100 rounded-full blur-3xl opacity-50" />
                
                <div className="relative z-10 flex flex-col items-center text-center">
                  <div className="w-20 h-20 bg-amber-50 rounded-3xl flex items-center justify-center mb-6 shadow-sm border border-amber-100">
                    <Database className="w-10 h-10 text-amber-500" />
                  </div>
                  
                  <h3 className="text-2xl font-black text-slate-800 mb-4 leading-tight">
                    Kapasitas Layanan <br/><span className="text-rose-600 uppercase tracking-tighter">Mencapai Batas</span>
                  </h3>
                  
                  <p className="text-slate-600 text-sm leading-relaxed mb-8 font-medium">
                    Layanan sinkronisasi database (Firestore) telah mencapai batas penggunaan <span className="text-rose-600 font-bold">Free Quota</span> harian. 
                    <br/><br/>
                    <div className="bg-rose-50 border border-rose-100 p-4 rounded-2xl text-left mb-4">
                      <p className="text-xs text-rose-800 font-bold mb-1 underline">Dampak Saat Ini (Mode Offline Caching):</p>
                      <ul className="text-xs text-rose-700 space-y-1 list-disc pl-4 mb-3">
                        <li>Fungsi <strong>Reset Database</strong> tidak akan tersimpan di cloud.</li>
                        <li>Update data (Absensi, Siswa, Profil) tidak akan sinkron ke perangkat lain.</li>
                        <li>Aplikasi berjalan sepenuhnya dari penyimpanan sementara (Cache / IndexedDB).</li>
                      </ul>
                      <div className="p-2 border-l-4 border-amber-500 bg-amber-50 rounded text-amber-800 text-[11px] font-bold">
                        ⚠️ PERINGATAN PENTING:<br/>
                        Selama kuota habis, MOHON JANGAN membersihkan histori browser (Clear Data / Clear Cache). Jika cache dihapus, maka data yang belum terkirim ke server akan HILANG permanen.
                      </div>
                    </div>
                    Jangan khawatir! Data baru Anda tetap <span className="font-bold text-slate-800">tersimpan sementara di memori perangkat ini</span>, dan akan dilanjutkan besok pagi saat kuota diatur ulang otomatis oleh Google.
                  </p>
                  
                  <div className="w-full flex flex-col gap-3">
                    <button 
                      onClick={() => setShowQuotaModal(false)}
                      className="w-full bg-slate-800 hover:bg-slate-900 text-white font-bold py-4 rounded-2xl transition-all shadow-lg hover:shadow-xl active:scale-95"
                    >
                      Saya Mengerti
                    </button>
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">
                      Infrastruktur didukung oleh Google Firebase Free Tier
                    </p>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Global Account Deleted Alert Modal */}
        <AnimatePresence>
          {accountDeletedAlert && (
             <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ scale: 0.9, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.9, opacity: 0, y: 20 }}
                className="relative bg-white rounded-[2rem] p-8 max-w-sm w-full shadow-2xl text-center border-2 border-rose-100"
              >
                <div className="w-16 h-16 bg-rose-100 text-rose-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-inner">
                  <UserX className="w-8 h-8" />
                </div>
                <h3 className="text-xl font-black text-slate-800 mb-3 tracking-tight">Akun Tidak Terdaftar!</h3>
                <div className="text-slate-600 text-sm leading-relaxed mb-8 font-medium">
                  Sistem mendeteksi bahwa akun Anda telah dihapus oleh Administrator. Sesi Anda dihentikan secara otomatis untuk alasan keamanan.
                </div>
                <button 
                  onClick={() => setAccountDeletedAlert(false)}
                  className="w-full bg-rose-500 hover:bg-rose-600 text-white font-bold py-4 rounded-xl transition-all shadow-lg hover:shadow-xl active:scale-95"
                >
                  Tutup
                </button>
              </motion.div>
            </div>
          )}
        </AnimatePresence>
      </main>

      {/* Action Password Verification Modal (Edit / Delete Student & Teacher) */}
      <AnimatePresence>
        {actionPasswordModal && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[300] flex items-center justify-center p-4"
          >
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 15 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 15 }}
              className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-md w-full space-y-6 border border-slate-100 text-left"
            >
              <div className="flex items-start gap-4">
                <div className="p-3 rounded-2xl bg-amber-50 text-amber-600">
                  <Lock className="w-6 h-6" />
                </div>
                <div>
                  <h3 className="text-xl font-black text-slate-800">{actionPasswordModal.title}</h3>
                  <p className="text-xs font-bold mt-0.5 tracking-wider uppercase text-amber-700">🔒 Autentikasi Diperlukan</p>
                </div>
              </div>

              <p className="text-sm text-slate-600 leading-relaxed bg-slate-50 border border-slate-100 p-4 rounded-2xl font-sans">
                {actionPasswordModal.description}
              </p>

              <div className="space-y-1.5">
                <label className="text-[10px] font-black text-slate-700 uppercase tracking-wider block">
                  {actionPasswordModal.expectedPassword === 'admin' || (activeUserCustomData?.username || '').toLowerCase().trim() === 'petugaspiket' || (activeUserCustomData?.username || '').toLowerCase().trim() === 'admin'
                    ? "Masukkan Password Admin:" 
                    : actionPasswordModal.expectedPassword 
                    ? "Masukkan Password Pengguna Baru:" 
                    : "Masukkan Password Akun Anda:"}
                </label>
                <div className="relative">
                  <input 
                    type={showActionPassword ? "text" : "password"} 
                    disabled={isVerifyingActionPassword}
                    value={actionPasswordInput}
                    onChange={e => {
                      setActionPasswordInput(e.target.value);
                      if (actionPasswordError) setActionPasswordError('');
                    }}
                    placeholder={
                      actionPasswordModal.expectedPassword === 'admin' || (activeUserCustomData?.username || '').toLowerCase().trim() === 'petugaspiket' || (activeUserCustomData?.username || '').toLowerCase().trim() === 'admin'
                        ? "Masukkan password admin" 
                        : "Masukkan password akun"
                    }
                    className="w-full px-4 py-3 pr-11 bg-white border-2 border-slate-300 rounded-xl focus:outline-none focus:border-[#098f41] focus:ring-2 focus:ring-emerald-100 transition-all text-sm font-bold text-slate-800 disabled:opacity-50"
                  />
                  <button
                    type="button"
                    onClick={() => setShowActionPassword(!showActionPassword)}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
                  >
                    {showActionPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
                {actionPasswordError && (
                  <p className="text-xs font-bold text-rose-600 flex items-center gap-1.5 mt-1.5 bg-rose-50 p-2 rounded-lg border border-rose-200">
                    <AlertCircle className="w-4 h-4 shrink-0" />
                    {actionPasswordError}
                  </p>
                )}
              </div>

              <div className="flex gap-4 pt-2">
                <button 
                  disabled={isVerifyingActionPassword}
                  onClick={() => {
                    setActionPasswordModal(null);
                    setActionPasswordInput('');
                    setActionPasswordError('');
                  }} 
                  className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all text-xs text-center disabled:opacity-50 cursor-pointer"
                >
                  Batal
                </button>
                <button 
                  disabled={isVerifyingActionPassword || !actionPasswordInput.trim()}
                  onClick={async () => {
                    const trimmedPass = actionPasswordInput.trim();
                    setIsVerifyingActionPassword(true);
                    setActionPasswordError('');

                    let isMatch = false;

                    // Get Admin password dynamically
                    let adminPass = '123456@#';
                    try {
                      const adminSnap = await getDoc(doc(dbDefault, 'custom_accounts', 'admin'));
                      if (adminSnap.exists() && adminSnap.data()?.password) {
                        adminPass = adminSnap.data().password.trim();
                      }
                    } catch (e) {
                      console.warn('Failed to load admin password, using default.', e);
                    }

                    if (actionPasswordModal.expectedPassword) {
                      if (actionPasswordModal.expectedPassword === 'admin') {
                        isMatch = (trimmedPass === '123456@#' || trimmedPass === adminPass);
                      } else {
                        isMatch = (trimmedPass === actionPasswordModal.expectedPassword.trim());
                      }
                    } else {
                      // No expectedPassword specified: standard action confirmation based on logged-in user
                      const loggedInUser = (activeUserCustomData?.username || '').toLowerCase().trim();

                      if (loggedInUser === 'petugaspiket') {
                        // Petugas piket must use admin's password
                        isMatch = (trimmedPass === '123456@#' || trimmedPass === adminPass);
                      } else if (loggedInUser === 'admin') {
                        // Admin must use admin password
                        isMatch = (trimmedPass === '123456@#' || trimmedPass === adminPass);
                      } else {
                        // Any other registered account must use their own password
                        if (activeUserCustomData?.password && trimmedPass === activeUserCustomData.password.trim()) {
                          isMatch = true;
                        } else {
                          // Look up centrally
                          try {
                            const snap = await getDoc(doc(dbDefault, 'custom_accounts', loggedInUser));
                            if (snap.exists() && snap.data()?.password === trimmedPass) {
                              isMatch = true;
                            }
                          } catch (err) {
                            console.warn('Central account lookup error:', err);
                          }
                        }
                      }
                    }

                    setIsVerifyingActionPassword(false);

                    if (!isMatch) {
                      setActionPasswordError('Sandi konfirmasi salah. Harap masukkan sandi yang tepat.');
                      return;
                    }

                    const cb = actionPasswordModal.onSuccess;
                    setActionPasswordModal(null);
                    setActionPasswordInput('');
                    cb();
                  }}
                  className="flex-1 px-5 py-3 rounded-xl font-bold text-white bg-[#098f41] hover:bg-[#077a37] shadow-[0_4px_12px_rgba(9,143,65,0.3)] transition-all text-xs flex items-center justify-center gap-2 disabled:opacity-50 cursor-pointer"
                >
                  {isVerifyingActionPassword && (
                    <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                  )}
                  Konfirmasi & Lanjutkan
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {classToEdit && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
          >
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 15 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 15 }}
              className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-md w-full space-y-6 border border-slate-100 text-left animate-in fade-in zoom-in duration-200"
            >
              <div className="flex items-start gap-4">
                <div className="p-3 rounded-2xl bg-blue-50 text-blue-600">
                  <Pencil className="w-6 h-6" />
                </div>
                <div className="flex-1 space-y-1">
                  <h3 className="text-xl font-bold text-slate-800">Edit Nama Kelas</h3>
                  <p className="text-slate-500 text-sm leading-relaxed">
                    Ubah nama kelas <strong className="text-slate-700">"{classToEdit}"</strong>. Perubahan ini juga akan memperbarui data kelas pada seluruh siswa dan riwayat absensi terkait secara otomatis.
                  </p>
                </div>
              </div>

              <div className="space-y-2">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider">Nama Kelas Baru</label>
                <input 
                  type="text" 
                  value={newClassNameInput}
                  onChange={(e) => setNewClassNameInput(e.target.value)}
                  className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-[#098f41]/20 focus:border-[#098f41] transition-all outline-none font-semibold text-slate-800"
                  placeholder="Contoh: X-A"
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault();
                      const saveBtn = document.getElementById('btn-save-class-edit');
                      if (saveBtn) saveBtn.click();
                    }
                  }}
                />
              </div>

              <div className="flex gap-4 pt-2">
                <button 
                  onClick={() => {
                    setClassToEdit(null);
                    setNewClassNameInput('');
                  }} 
                  className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all text-xs text-center cursor-pointer"
                >
                  Batal
                </button>
                <button 
                  id="btn-save-class-edit"
                  onClick={() => {
                    const trimmed = newClassNameInput.trim();
                    if (!trimmed) {
                      showToast('Nama kelas tidak boleh kosong', 'error');
                      return;
                    }
                    if (trimmed === classToEdit) {
                      setClassToEdit(null);
                      setNewClassNameInput('');
                      return;
                    }
                    if (classList.some(existing => existing.toLowerCase() === trimmed.toLowerCase())) {
                      showToast('Kelas "' + trimmed + '" sudah terdaftar', 'error');
                      return;
                    }
                    const updatedList = classList.map(existing => existing === classToEdit ? trimmed : existing);
                    updatedList.sort((a,b) => a.localeCompare(b, 'id-ID', { numeric: true }));
                    setClassList(updatedList);
                    
                    const updatedStudents = students.map(s => s.class === classToEdit ? { ...s, class: trimmed } : s);
                    setStudents(updatedStudents);
                    
                    const updatedSessions = attendanceSessions.map(sess => sess.className === classToEdit ? { ...sess, className: trimmed } : sess);
                    setAttendanceSessions(updatedSessions);

                    if (currentUser) {
                      const batch = writeBatch(activeDb);
                      batch.set(doc(activeDb, 'users', currentUser.uid), { classList: updatedList }, { merge: true });
                      
                      updatedStudents.forEach(s => {
                        if (s.class === trimmed) {
                          batch.set(doc(activeDb, 'students', s.id), { class: trimmed }, { merge: true });
                        }
                      });
                      
                      updatedSessions.forEach(sess => {
                        if (sess.className === trimmed) {
                          batch.set(doc(activeDb, 'attendanceSessions', sess.id), { className: trimmed }, { merge: true });
                        }
                      });

                      batch.commit().then(() => {
                        showToast(`Kelas ${classToEdit} berhasil diubah menjadi ${trimmed}`, 'success');
                      }).catch(e => console.error(e));
                    }
                    setClassToEdit(null);
                    setNewClassNameInput('');
                  }}
                  className="flex-1 px-5 py-3 rounded-xl font-bold text-white bg-[#098f41] hover:bg-[#077a37] shadow-[0_4px_12px_rgba(9,143,65,0.3)] transition-all text-xs text-center cursor-pointer"
                >
                  Simpan Perubahan
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {classToDelete && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4"
          >
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 15 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 15 }}
              className="bg-white rounded-3xl p-6 md:p-8 shadow-2xl max-w-md w-full space-y-6 border border-slate-100 text-left animate-in fade-in zoom-in duration-200"
            >
              <div className="flex items-start gap-4">
                <div className="p-3 rounded-2xl bg-rose-50 text-rose-600">
                  <Trash2 className="w-6 h-6" />
                </div>
                <div className="flex-1 space-y-1">
                  <h3 className="text-xl font-bold text-slate-800">Konfirmasi Hapus Kelas</h3>
                  <p className="text-slate-500 text-sm leading-relaxed">
                    Apakah Anda yakin ingin menghapus kelas <strong className="text-slate-700">"{classToDelete.name}"</strong>?
                  </p>
                  <p className="text-slate-500 text-xs mt-2 bg-slate-50 p-3 rounded-xl border border-slate-100 leading-relaxed font-medium">
                    {classToDelete.studentCount > 0 
                      ? `PERINGATAN: Terdapat ${classToDelete.studentCount} siswa di kelas ini. Jika dihapus, data siswa tersebut TIDAK akan hilang, namun kolom Kelas mereka akan dikosongkan.` 
                      : 'Kelas ini kosong dan aman untuk dihapus.'}
                  </p>
                </div>
              </div>

              <div className="flex gap-4 pt-2">
                <button 
                  onClick={() => setClassToDelete(null)} 
                  className="flex-1 px-5 py-3 rounded-xl font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all text-xs text-center cursor-pointer"
                >
                  Batal
                </button>
                <button 
                  onClick={() => {
                    const cName = classToDelete.name;
                    const updatedList = classList.filter(existing => existing !== cName);
                    setClassList(updatedList);
                    
                    const updatedStudents = students.map(s => s.class === cName ? { ...s, class: '' } : s);
                    setStudents(updatedStudents);

                    if (currentUser) {
                      const batch = writeBatch(activeDb);
                      batch.set(doc(activeDb, 'users', currentUser.uid), { classList: updatedList }, { merge: true });
                      
                      students.forEach(s => {
                        if (s.class === cName) {
                          batch.set(doc(activeDb, 'students', s.id), { class: '' }, { merge: true });
                        }
                      });

                      batch.commit()
                        .then(() => showToast(`Kelas ${cName} berhasil dihapus`, 'success'))
                        .catch(e => console.error(e));
                    }
                    setClassToDelete(null);
                  }}
                  className="flex-1 px-5 py-3 rounded-xl font-bold text-white bg-rose-600 hover:bg-rose-700 shadow-[0_4px_12px_rgba(225,29,72,0.3)] transition-all text-xs text-center cursor-pointer"
                >
                  Hapus Kelas
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Bottom Nav */}
      <footer className="bg-white border-t border-slate-100/80 px-2 py-3 grid grid-cols-4 gap-1.5 z-40 shadow-[0_-4px_22px_rgba(148,163,184,0.06)] shrink-0 print:hidden">
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = activeTab === item.id;
          return (
            <button
              key={item.id}
              onClick={() => {
                  setActiveTab(item.id as 'dashboard' | 'students' | 'attendance' | 'reports' | 'homeroom_report' | 'profile');
              }}
              className={`flex flex-col items-center justify-center gap-1.5 py-1.5 px-0.5 rounded-2xl text-xs font-bold transition-all duration-300 transform active:scale-95 group relative ${
                isActive 
                  ? `${item.activeBg} font-black` 
                  : `${item.inactiveColor} hover:text-slate-800 ${item.hoverBg}`
              }`}
            >
              <div className={`p-1.5 sm:p-2.5 rounded-xl transition-all duration-300 ${
                isActive 
                  ? `${item.iconBg} scale-110 shadow-lg shadow-current/15` 
                  : `${item.inactiveBg} group-hover:bg-slate-100 group-hover:text-slate-600`
              }`}>
                <Icon className="w-5 h-5 stroke-[2.3]" />
              </div>
              <span className={`text-[9px] sm:text-[10px] tracking-tight font-black mt-0.5 ${!isActive ? item.inactiveColor : ''}`}>{item.label}</span>
            </button>
          );
        })}
      </footer>
    </div>
  );
}
