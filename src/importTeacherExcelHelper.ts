import * as XLSX from 'xlsx';
import { Firestore, writeBatch, doc } from 'firebase/firestore';
import { Auth } from 'firebase/auth';

export interface Teacher {
  id: string;
  name: string;
  niy: string;
  dutyDay?: string;
  userId?: string;
}

export interface ImportResultState {
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
}

/**
 * Normalizes day names from various formats (numbers, Indonesian, English, abbreviations, mixed casing)
 * to standard Indonesian day name: "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"
 */
export function normalizeDutyDay(val: unknown): string {
  if (val === undefined || val === null) return '';
  const str = String(val).trim();
  if (!str || str === '-' || str === '–' || (str === '0' && typeof val === 'string' && val.trim() === '-')) return '';

  // Numeric check (1=Senin, 2=Selasa, 3=Rabu, 4=Kamis, 5=Jumat, 6=Sabtu, 7 or 0=Minggu)
  if (typeof val === 'number' || /^[0-7]$/.test(str)) {
    const num = Number(str);
    const dayMap: Record<number, string> = {
      1: 'Senin',
      2: 'Selasa',
      3: 'Rabu',
      4: 'Kamis',
      5: 'Jumat',
      6: 'Sabtu',
      7: 'Minggu',
      0: 'Minggu'
    };
    if (dayMap[num]) return dayMap[num];
  }

  // Single day direct matching
  const clean = str.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (clean.startsWith('sen') || clean.startsWith('mon')) return 'Senin';
  if (clean.startsWith('sel') || clean.startsWith('tue')) return 'Selasa';
  if (clean.startsWith('rab') || clean.startsWith('wed')) return 'Rabu';
  if (clean.startsWith('kam') || clean.startsWith('thu')) return 'Kamis';
  if (clean.startsWith('jum') || clean.startsWith('fri')) return 'Jumat';
  if (clean.startsWith('sab') || clean.startsWith('sat')) return 'Sabtu';
  if (clean.startsWith('min') || clean.startsWith('ahad') || clean.startsWith('sun')) return 'Minggu';

  // Multi-day pattern matching (e.g. "Senin, Kamis" or "Senin - Rabu")
  const dayPatterns = [
    { regex: /(senin|monday|\bsen\b)/i, name: 'Senin' },
    { regex: /(selasa|tuesday|\bsel\b)/i, name: 'Selasa' },
    { regex: /(rabu|wednesday|\brab\b)/i, name: 'Rabu' },
    { regex: /(kamis|thursday|\bkam\b)/i, name: 'Kamis' },
    { regex: /(jumat|jum'at|jum`at|jum at|friday|\bjum\b)/i, name: 'Jumat' },
    { regex: /(sabtu|saturday|\bsab\b)/i, name: 'Sabtu' },
    { regex: /(minggu|sunday|ahad|\bmin\b)/i, name: 'Minggu' }
  ];

  const matchedDays: string[] = [];
  for (const item of dayPatterns) {
    if (item.regex.test(str) && !matchedDays.includes(item.name)) {
      matchedDays.push(item.name);
    }
  }

  if (matchedDays.length > 0) {
    return matchedDays.join(', ');
  }

  return str;
}

/**
 * Returns the JavaScript day index (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
 * from standard strings or numeric inputs.
 */
export function getDayIndex(dayStrOrNum: string | number | undefined | null): number {
  if (dayStrOrNum === undefined || dayStrOrNum === null || dayStrOrNum === '') return -1;
  const str = String(dayStrOrNum).trim().toLowerCase();
  if (/^[0-7]$/.test(str)) {
    const n = parseInt(str, 10);
    return n === 7 ? 0 : n;
  }
  if (str.startsWith('sen') || str.startsWith('mon')) return 1;
  if (str.startsWith('sel') || str.startsWith('tue')) return 2;
  if (str.startsWith('rab') || str.startsWith('wed')) return 3;
  if (str.startsWith('kam') || str.startsWith('thu')) return 4;
  if (str.startsWith('jum') || str.startsWith('fri')) return 5;
  if (str.startsWith('sab') || str.startsWith('sat')) return 6;
  if (str.startsWith('min') || str.startsWith('ahad') || str.startsWith('sun')) return 0;
  return -1;
}

export const importTeacherExcelHelper = async (
  file: File,
  activeAuth: Auth,
  activeDb: Firestore,
  setImportResult: React.Dispatch<React.SetStateAction<ImportResultState | null>>,
  showToast: (msg: string, type: 'success' | 'info' | 'error') => void,
  excelInputRef: React.RefObject<HTMLInputElement | null>,
  existingTeachers: Teacher[] = [],
  currentUserId?: string
): Promise<void> => {
  return new Promise((resolve) => {
    const uid = currentUserId || activeAuth.currentUser?.uid || 'admin';
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(new Uint8Array(bstr), { type: 'array' });
        
        let successCount = 0;
        let skipCount = 0;
        let failCount = 0;
        let emptyCount = 0;
        let totalParsed = 0;
        const failedRowDetails: string[] = [];
        const sheetsProcessed: { name: string; count: number }[] = [];

        let batch = writeBatch(activeDb);
        let batchCounter = 0;
        const promises: Promise<void>[] = [];

        showToast('Sedang membaca file Excel Guru...', 'info');

        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as (string | number | boolean | null | undefined)[][];

          if (!data || data.length === 0) continue;

          let headerRowIdx = -1;
          let nameIdx = -1;
          let niyIdx = -1;
          let dutyDayIdx = -1;

          for (let r = 0; r < Math.min(data.length, 30); r++) {
            const row = data[r];
            if (!row || !Array.isArray(row)) continue;

            const currentNameIdx = row.findIndex(h => {
              if (h === undefined || h === null) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              return str.includes('nama') || str.includes('guru') || str.includes('pendidik') || str.includes('pengajar') || str === 'name';
            });

            const currentNiyIdx = row.findIndex((h, idx) => {
              if (h === undefined || h === null || idx === currentNameIdx) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              return str === 'niy' || str === 'nip' || str.includes('nomor') || str.includes('id') || str.includes('yayasan') || str.includes('induk') || str.includes('pegawai');
            });

            const currentDutyDayIdx = row.findIndex((h, idx) => {
              if (h === undefined || h === null || idx === currentNameIdx || idx === currentNiyIdx) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              return (
                str.includes('piket') ||
                str.includes('jadwal') ||
                str.includes('duty') ||
                str.includes('hari') ||
                str.includes('day') ||
                str.includes('tugas')
              );
            });

            if (currentNameIdx !== -1) {
              headerRowIdx = r;
              nameIdx = currentNameIdx;
              if (currentNiyIdx !== -1) {
                niyIdx = currentNiyIdx;
              }
              if (currentDutyDayIdx !== -1) {
                dutyDayIdx = currentDutyDayIdx;
              }
              break;
            }
          }

          if (headerRowIdx === -1 || nameIdx === -1) {
             failedRowDetails.push(`Sheet "${sheetName}": Tidak menemukan kolom 'Nama' guru. Diabaikan.`);
             continue;
          }

          let sheetAdded = 0;
          for (let r = headerRowIdx + 1; r < data.length; r++) {
            const row = data[r];
            if (!row || !Array.isArray(row)) continue;

            totalParsed++;

            const rawName = row[nameIdx];
            const rawNiy = niyIdx !== -1 ? row[niyIdx] : undefined;
            const rawDutyDay = dutyDayIdx !== -1 ? row[dutyDayIdx] : undefined;

            if (rawName === undefined || rawName === null || String(rawName).trim() === '') {
              emptyCount++;
              continue;
            }

            const name = String(rawName).trim();
            let niy = rawNiy !== undefined && rawNiy !== null ? String(rawNiy).trim() : '';

            niy = niy.replace(/[^a-zA-Z0-9.\- ]/g, '').trim();
            if (!niy || niy === '' || niy === '-') {
              niy = '-';
            }

            const dutyDay = normalizeDutyDay(rawDutyDay);

            // Check if teacher already exists
            const existingTeacher = existingTeachers.find(t => 
              (niy !== '-' && t.niy && t.niy.trim() !== '-' && t.niy.trim() === niy) ||
              (t.name.trim().toLowerCase() === name.toLowerCase())
            );

            if (existingTeacher) {
              // If dutyDay is provided and can update or complete the teacher record
              if (dutyDay && (!existingTeacher.dutyDay || existingTeacher.dutyDay === '-' || existingTeacher.dutyDay !== dutyDay)) {
                const teacherRef = doc(activeDb, 'teachers', existingTeacher.id);
                batch.update(teacherRef, { dutyDay });
                batchCounter++;
                successCount++;
                sheetAdded++;
              } else {
                skipCount++;
              }
              continue;
            }

            try {
              const newId = `teacher_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
              const teacherRef = doc(activeDb, 'teachers', newId);
              batch.set(teacherRef, {
                id: newId,
                name,
                niy,
                dutyDay: dutyDay || '',
                userId: uid
              });

              batchCounter++;
              successCount++;
              sheetAdded++;

              if (batchCounter >= 490) {
                const currentBatch = batch;
                batch = writeBatch(activeDb);
                batchCounter = 0;
                promises.push(currentBatch.commit());
              }
            } catch (err: unknown) {
              failCount++;
              const errorMsg = err instanceof Error ? err.message : String(err);
              if (failedRowDetails.length < 50) {
                failedRowDetails.push(`Guru: ${name} (Error: ${errorMsg})`);
              }
            }
          }
          if (sheetAdded > 0) {
            sheetsProcessed.push({ name: sheetName, count: sheetAdded });
          }
        }

        if (batchCounter > 0) {
          promises.push(batch.commit());
        }

        if (promises.length > 0) {
          showToast(`Mengunggah ${successCount} data guru ke database...`, 'info');
          await Promise.all(promises);
        }

        setImportResult({
          isOpen: true,
          successCount,
          skipCount,
          failCount,
          emptyCount,
          totalParsed,
          error: false,
          details: failedRowDetails,
          sheetsProcessed
        });

      } catch (err: unknown) {
        console.error("Error parsing Teacher Excel:", err);
        const errorMsg = err instanceof Error ? err.message : String(err);
        setImportResult({
          isOpen: true,
          successCount: 0,
          skipCount: 0,
          failCount: 0,
          emptyCount: 0,
          totalParsed: 0,
          error: true,
          errorMessage: 'Gagal membaca format file Excel. Pastikan file tidak rusak.',
          details: [errorMsg]
        });
      } finally {
        if (excelInputRef.current) {
          excelInputRef.current.value = '';
        }
        resolve();
      }
    };

    reader.onerror = () => {
      showToast('Terjadi kesalahan saat membaca file.', 'error');
      if (excelInputRef.current) excelInputRef.current.value = '';
      resolve();
    };

    reader.readAsArrayBuffer(file);
  });
};
