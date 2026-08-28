import * as XLSX from 'xlsx';
import { Firestore, writeBatch, doc, setDoc } from 'firebase/firestore';
import { Auth } from 'firebase/auth';

export interface Student {
  id: string;
  name: string;
  nisn: string;
  class: string;
  waliKelas?: string;
  waliKelasNiy?: string;
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

export const importExcelHelper = async (
  file: File,
  activeAuth: Auth,
  activeDb: Firestore,
  classList: string[],
  setClassList: React.Dispatch<React.SetStateAction<string[]>>,
  setImportResult: React.Dispatch<React.SetStateAction<ImportResultState | null>>,
  showToast: (msg: string, type: 'success' | 'info' | 'error') => void,
  excelInputRef: React.RefObject<HTMLInputElement | null>,
  existingStudents: Student[] = [],
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
        const discoveredClasses: string[] = [];
        const discoveredClassWali: Record<string, string> = {};
        const discoveredClassWaliNiy: Record<string, string> = {};

        let batch = writeBatch(activeDb);
        let batchCounter = 0;
        const promises: Promise<void>[] = [];

        showToast('Sedang membaca file Excel...', 'info');

        // Process all sheets in the workbook
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as (string | number | boolean | null | undefined)[][];

          if (!data || data.length === 0) {
            continue; // Skip empty sheets
          }

          // Search for the header row with student name
          let headerRowIdx = -1;
          let classIdx = -1;
          let nameIdx = -1;
          let nisnIdx = -1;
          let waliKelasIdx = -1;
          let waliKelasNiyIdx = -1;

          // Try to search first 30 rows of the sheet
          for (let r = 0; r < Math.min(data.length, 30); r++) {
            const row = data[r];
            if (!row || !Array.isArray(row)) continue;

            const currentClassIdx = row.findIndex(h => {
              if (h === undefined || h === null) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              return str === 'kelas' || str === 'rombel' || str === 'rombongan' || str.includes('kelas') || str.includes('rombel') || str.includes('rombongan') || str.includes('tingkat') || str.includes('group') || str === 'class';
            });

            const currentNameIdx = row.findIndex(h => {
              if (h === undefined || h === null) return false;
              const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
              if (str.includes('ayah') || str.includes('ibu') || str.includes('wali') || str.includes('ortu') || str.includes('orangtua') || str.includes('panggilan')) return false;
              return str === 'nama' || str === 'siswa' || str.includes('nama') || str.includes('siswa') || str.includes('pesertadidik') || str.includes('namalengkap') || str.includes('fullname') || str === 'name';
            });

            if (currentClassIdx !== -1 && currentNameIdx !== -1) {
              headerRowIdx = r;
              classIdx = currentClassIdx;
              nameIdx = currentNameIdx;
              nisnIdx = row.findIndex(h => {
                if (h === undefined || h === null) return false;
                const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                return str.includes('nis') || str.includes('nisn') || str.includes('nomorinduk') || str.includes('induk');
              });
              waliKelasIdx = row.findIndex(h => {
                if (h === undefined || h === null) return false;
                const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                if (str.includes('niy') || str.includes('nuptk') || str.includes('nip')) return false;
                return str.includes('walikelas') || str.includes('homeroom') || str.includes('pembimbing') || (str.includes('wali') && !str.includes('orangtua') && !str.includes('ortu'));
              });
              waliKelasNiyIdx = row.findIndex(h => {
                if (h === undefined || h === null) return false;
                const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                return str.includes('niywali') || str.includes('nuptkwali') || str.includes('nipwali') || str.includes('niywalikelas') || str.includes('nuptkwalikelas') || str.includes('nipwalikelas') || (str.includes('niy') && (str.includes('wali') || str.length < 5)) || (str.includes('nuptk') && (str.includes('wali') || str.length < 8)) || str === 'niy' || str === 'nip' || str === 'nuptk';
              });
              break;
            }
          }

          // Fallback: If combined header class + name is not found, search for at least "nama"
          if (headerRowIdx === -1) {
            for (let r = 0; r < Math.min(data.length, 15); r++) {
              const row = data[r];
              if (!row || !Array.isArray(row)) continue;
              const currentNameIdx = row.findIndex(h => {
                if (h === undefined || h === null) return false;
                const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                if (str.includes('ayah') || str.includes('ibu') || str.includes('wali') || str.includes('ortu') || str.includes('orangtua') || str.includes('panggilan')) return false;
                return str.includes('nama') || str.includes('siswa') || str === 'name';
              });
              if (currentNameIdx !== -1) {
                headerRowIdx = r;
                nameIdx = currentNameIdx;
                classIdx = row.findIndex(h => {
                  if (h === undefined || h === null) return false;
                  const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                  return str.includes('kelas') || str.includes('rombel') || str === 'class';
                });
                nisnIdx = row.findIndex(h => {
                  if (h === undefined || h === null) return false;
                  const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                  return str.includes('nis') || str.includes('nisn');
                });
                waliKelasIdx = row.findIndex(h => {
                  if (h === undefined || h === null) return false;
                  const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                  if (str.includes('niy') || str.includes('nuptk') || str.includes('nip')) return false;
                  return str.includes('walikelas') || str.includes('homeroom') || str.includes('pembimbing') || (str.includes('wali') && !str.includes('orangtua') && !str.includes('ortu'));
                });
                waliKelasNiyIdx = row.findIndex(h => {
                  if (h === undefined || h === null) return false;
                  const str = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
                  return str.includes('niywali') || str.includes('nuptkwali') || str.includes('nipwali') || str.includes('niywalikelas') || str.includes('nuptkwalikelas') || str.includes('nipwalikelas') || (str.includes('niy') && (str.includes('wali') || str.length < 5)) || (str.includes('nuptk') && (str.includes('wali') || str.length < 8)) || str === 'niy' || str === 'nip' || str === 'nuptk';
                });
                break;
              }
            }
          }

          // Final fallback: use row 0 and look up "nama"
          if (headerRowIdx === -1) {
            const firstRow = data[0];
            if (firstRow && Array.isArray(firstRow)) {
              nameIdx = firstRow.findIndex(h => {
                if (!h) return false;
                const str = h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                if (str.includes('ayah') || str.includes('ibu') || str.includes('wali') || str.includes('ortu') || str.includes('orangtua') || str.includes('panggilan')) return false;
                return str.includes('nama') || str === 'name';
              });
              classIdx = firstRow.findIndex(h => h?.toString().toLowerCase().includes('kelas') || h?.toString().toLowerCase() === 'class');
              nisnIdx = firstRow.findIndex(h => h?.toString().toLowerCase().includes('nis'));
              if (nameIdx !== -1) {
                headerRowIdx = 0;
              }
            }
          }

          // Brutal fallback: If totally missing headers, let's just assume Column B (index 1) is Name if Column A is numbers, 
          // OR Column A is name. We'll find the first column that has mostly string data.
          if (headerRowIdx === -1 && data.length > 0) {
             for (let r = 0; r < Math.min(data.length, 5); r++) {
                 const row = data[r];
                 if (!row || !Array.isArray(row)) continue;
                 
                 // Look for a cell that is a string and represents a name
                 const possibleNameIdx = row.findIndex((cell) => {
                     if (typeof cell !== 'string') return false;
                     const s = cell.trim();
                     return s.length > 2 && s.length < 50 && !/^\d+$/.test(s);
                 });

                 if (possibleNameIdx !== -1) {
                     nameIdx = possibleNameIdx;
                     headerRowIdx = Math.max(0, r - 1); // Assume previous row or same row is header
                     break;
                 }
             }
          }

          // If no student name column could be identified on this sheet, skip this sheet gracefully
          if (nameIdx === -1) {
            failedRowDetails.push(`Sheet [${sheetName}]: Dilewati karena kolom Nama/Siswa tidak ditemukan.`);
            continue;
          }

          const rows = data.slice(headerRowIdx + 1) as (string | number | boolean | null | undefined)[][];
          let sheetSuccessCount = 0;
          let lastClass = '';

          for (const [i, row] of rows.entries()) {
            totalParsed++;
            if (!row || row.length === 0) {
              emptyCount++;
              continue;
            }

            const isRowBlank = row.every(cell => cell === undefined || cell === null || String(cell).trim() === '');
            if (isRowBlank) {
              emptyCount++;
              continue;
            }

            // More resilient column data extraction:
            const rawName = row[nameIdx] ?? '';
            const rawNisn = (nisnIdx !== -1) ? (row[nisnIdx] ?? '') : '';
            const rawClass = (classIdx !== -1) ? (row[classIdx] ?? null) : null;
            const rawWali = (waliKelasIdx !== -1) ? (row[waliKelasIdx] ?? '') : '';
            const rawWaliNiy = (waliKelasNiyIdx !== -1) ? (row[waliKelasNiyIdx] ?? '') : '';

            const currentName = rawName ? String(rawName).trim() : '';
            const currentWali = rawWali ? String(rawWali).trim() : '';
            const currentWaliNiy = rawWaliNiy ? String(rawWaliNiy).trim() : '';

            if (rawClass && String(rawClass).trim()) {
              lastClass = String(rawClass).trim();
            }

            // Smart class fallback: if no class cell was found, try the sheetName (especially if sheets are named like 'XII IPA 1')
            let currentClass = lastClass || String(rawClass || '').trim();
            if (!currentClass) {
              // Check if the sheet name looks like a class (is shorter than 15 characters, doesn't say Sheet1 etc.)
              const isSheetGeneric = sheetName.toLowerCase().startsWith('sheet') || sheetName.toLowerCase().includes('halaman');
              currentClass = isSheetGeneric ? 'Umum' : sheetName.trim();
            }

            if (currentWali && currentClass) {
              discoveredClassWali[currentClass] = currentWali;
            }
            if (currentWaliNiy && currentClass) {
              discoveredClassWaliNiy[currentClass] = currentWaliNiy;
            }

            if (currentName) {
              const cleanNisn = rawNisn ? String(rawNisn).trim().replace(/\D/g, '') : '';
              
              // Check for duplicates in existing data or current batch to prevent double imports
              const isDuplicate = existingStudents.some(s => 
                (cleanNisn && s.nisn === cleanNisn) || 
                (s.name.toLowerCase() === currentName.toLowerCase() && s.class.toLowerCase() === currentClass.toLowerCase())
              );

              if (isDuplicate) {
                skipCount++;
                continue;
              }

              const newId = Date.now().toString() + '-' + Math.floor(Math.random() * 1000) + '-' + totalParsed;
              
              const student: Student = {
                id: newId,
                name: currentName,
                nisn: cleanNisn,
                class: currentClass,
                waliKelas: currentWali || discoveredClassWali[currentClass] || '',
                waliKelasNiy: currentWaliNiy || discoveredClassWaliNiy[currentClass] || '',
                userId: uid
              };

              batch.set(doc(activeDb, 'students', newId), student);
              successCount++;
              sheetSuccessCount++;
              batchCounter++;

              if (!discoveredClasses.includes(currentClass)) {
                discoveredClasses.push(currentClass);
              }

              if (batchCounter === 450) {
                promises.push(batch.commit());
                batch = writeBatch(activeDb);
                batchCounter = 0;
              }
            } else {
              failCount++;
              const rowNumber = headerRowIdx + 1 + i + 1;
              failedRowDetails.push(`Sheet [${sheetName}] Baris ${rowNumber}: Nama kosong atau baris tidak valid`);
            }
          }

          if (sheetSuccessCount > 0 || rows.length > 0) {
            sheetsProcessed.push({
              name: sheetName,
              count: sheetSuccessCount
            });
          }
        }

        if (batchCounter > 0) {
          promises.push(batch.commit());
        }

        // Wait for all commits to complete reliably
        if (promises.length > 0) {
          await Promise.all(promises);
        }

        // Update class list in Firestore with newly discovered classes
        if (discoveredClasses.length > 0 || Object.keys(discoveredClassWali).length > 0 || Object.keys(discoveredClassWaliNiy).length > 0) {
          const updatedClasses = Array.from(new Set([...classList, ...discoveredClasses]));
          updatedClasses.sort((a, b) => a.localeCompare(b, 'id-ID', { numeric: true }));

          // Save directly to Firestore users document for guaranteed sync
          try {
            const updatePayload: Record<string, unknown> = {};
            if (discoveredClasses.length > 0) updatePayload.classList = updatedClasses;
            if (Object.keys(discoveredClassWali).length > 0) updatePayload.classWaliMap = discoveredClassWali;
            if (Object.keys(discoveredClassWaliNiy).length > 0) updatePayload.classWaliNiyMap = discoveredClassWaliNiy;
            
            if (activeAuth.currentUser?.uid) {
              await setDoc(doc(activeDb, 'users', activeAuth.currentUser.uid), updatePayload, { merge: true });
              await setDoc(doc(activeDb, 'settings', 'school_data'), updatePayload, { merge: true });
              console.log("Successfully updated classList, classWaliMap, and classWaliNiyMap in Firestore");
            } else {
              await setDoc(doc(activeDb, 'settings', 'school_data'), updatePayload, { merge: true });
            }
          } catch (e) {
            console.error("Failed to update classList/classWaliMap/classWaliNiyMap in Firestore:", e);
            // Fallback to local state if DB update fails (or Toast if critical)
            setClassList(updatedClasses);
          }
        }

        if (excelInputRef.current) excelInputRef.current.value = '';

        if (successCount === 0 && sheetsProcessed.length === 0) {
          // If nothing was parsed successfully
          setImportResult({
            isOpen: true,
            successCount: 0,
            skipCount: skipCount + failCount + emptyCount,
            failCount,
            emptyCount,
            totalParsed,
            error: true,
            errorMessage: 'Format tabel siswa atau kolom nama tidak terdeteksi di lembar kerja Excel mana pun.',
            details: failedRowDetails,
            sheetsProcessed
          });
          showToast('Gagal mengimpor data! Silakan periksa format file Excel.', 'error');
        } else {
          setImportResult({
            isOpen: true,
            successCount,
            skipCount: skipCount + failCount + emptyCount,
            failCount,
            emptyCount,
            totalParsed,
            details: failedRowDetails,
            sheetsProcessed
          });
          
          if (skipCount > 0) {
            showToast(`Berhasil mengimpor ${successCount} siswa. ${skipCount} data duplikat dilewati.`, 'info');
          } else if (failCount > 0) {
            showToast(`Berhasil mengimpor ${successCount} siswa. ${failCount} baris bermasalah.`, 'info');
          } else {
            showToast(`Berhasil mengimpor total ${successCount} siswa dari ${sheetsProcessed.length} lembar kerja.`, 'success');
          }
        }
        resolve();
      } catch (err) {
        console.error(err);
        if (excelInputRef.current) excelInputRef.current.value = '';
        const errMsg = (err as Error).message || '';
        
        let friendlyError = 'Terjadi kesalahan sistem saat membaca file Excel: ' + errMsg;
        if (errMsg.toLowerCase().includes('permission') || errMsg.toLowerCase().includes('missing or insufficient')) {
            friendlyError = 'Sistem menolak menyimpan data (Missing Permissions). Hal ini biasanya terjadi jika Anda menggunakan Database Firebase mandiri namun belum memperbarui "Firestore Rules" menjadi allow read, write: if request.auth != null;';
        }

        setImportResult({
          isOpen: true,
          successCount: 0,
          skipCount: 0,
          failCount: 0,
          emptyCount: 0,
          totalParsed: 0,
          error: true,
          errorMessage: friendlyError
        });
        showToast('Gagal memproses file Excel (Tertolak atau error format).', 'error');
        resolve();
      }
    };

    reader.onerror = () => {
      if (excelInputRef.current) excelInputRef.current.value = '';
      setImportResult({
        isOpen: true,
        successCount: 0,
        skipCount: 0,
        failCount: 0,
        emptyCount: 0,
        totalParsed: 0,
        error: true,
        errorMessage: 'Gagal mengunggah file Excel dari laptop/komputer Anda.'
      });
      showToast('Gagal membaca file.', 'error');
      resolve();
    };

    reader.readAsArrayBuffer(file);
  });
};
