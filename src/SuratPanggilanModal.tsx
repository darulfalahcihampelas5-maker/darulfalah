/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { useState } from 'react';
import { X, Printer, User, Calendar, AlertTriangle, FileText } from 'lucide-react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { kopSuratBase64 } from './kop-surat-b64';

export interface StudentForCall {
  id: string;
  name: string;
  nisn: string;
  class: string;
  waliKelas?: string;
  waliKelasNiy?: string;
  hadir?: number;
  sakit?: number;
  izin?: number;
  alpa?: number;
  dispen?: number;
  total?: number;
  totalTidakHadir?: number;
  persentase?: number;
  rate?: number;
  datesAlpa?: string[];
  datesSakit?: string[];
  datesIzin?: string[];
  datesDispen?: string[];
}

export interface StudentStatsForCall {
  hadir: number;
  sakit: number;
  izin: number;
  alpa: number;
  dispen?: number;
  total: number;
  totalTidakHadir?: number;
  rate: number;
  datesAlpa?: string[];
  datesSakit?: string[];
  datesIzin?: string[];
  datesDispen?: string[];
}

interface SuratPanggilanModalProps {
  isOpen: boolean;
  onClose: () => void;
  student: StudentForCall | null;
  stats?: StudentStatsForCall | null;
  profileData?: Record<string, unknown>;
  showToast?: (message: string, type: 'success' | 'info' | 'error') => void;
}

export const SuratPanggilanModal: React.FC<SuratPanggilanModalProps> = ({
  isOpen,
  onClose,
  student,
  stats,
  profileData,
  showToast
}) => {
  const todayStr = new Date().toLocaleDateString('id-ID', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });

  const [nomorSurat, setNomorSurat] = useState(`421.3/${Math.floor(100 + Math.random() * 900)}/SMA-DF/${new Date().getFullYear()}`);
  const [lampiran, setLampiran] = useState('1 Lembar Rekap Presensi');
  const [perihal, setPerihal] = useState('Pemanggilan Orang Tua / Wali Murid (Pembinaan Presensi)');
  const [tanggalSurat, setTanggalSurat] = useState(todayStr);
  
  const [hariTanggal, setHariTanggal] = useState('');
  const [waktu, setWaktu] = useState('09.00 WIB - Selesai');
  const [tempat, setTempat] = useState('Ruang Guru / Ruang BK SMA Darul Falah');
  const [menemui, setMenemui] = useState('Wali Kelas & Guru BK');

  const [namaWali, setNamaWali] = useState('');
  const [niyWali, setNiyWali] = useState('');
  const [namaKepsek, setNamaKepsek] = useState('');
  const [nipKepsek, setNipKepsek] = useState('');

  const [isGenerating, setIsGenerating] = useState(false);

  // Sync state values when student data changes
  React.useEffect(() => {
    if (student) {
      setNamaWali(student.waliKelas || '');
      setNiyWali(student.waliKelasNiy || '');
      setNamaKepsek((profileData?.namaKepalaSekolah as string) || '');
      setNipKepsek((profileData?.nipKepalaSekolah as string) || '');
    }
  }, [student, profileData]);

  if (!isOpen || !student) return null;

  // Compute presence and absence metrics
  const hadirCount = stats?.hadir ?? student.hadir ?? 0;
  const sakitCount = stats?.sakit ?? student.sakit ?? 0;
  const izinCount = stats?.izin ?? student.izin ?? 0;
  const alpaCount = stats?.alpa ?? student.alpa ?? 0;
  const dispenCount = stats?.dispen ?? student.dispen ?? 0;
  const totalTidakHadir = stats?.totalTidakHadir ?? student.totalTidakHadir ?? (sakitCount + izinCount + alpaCount + dispenCount);
  const totalPertemuan = stats?.total ?? student.total ?? (hadirCount + totalTidakHadir);
  const rawRate = stats?.rate ?? student.rate ?? student.persentase ?? (totalPertemuan > 0 ? (hadirCount / totalPertemuan) * 100 : 0);
  const rateStr = `${typeof rawRate === 'number' ? rawRate.toFixed(1) : rawRate}%`;

  const alpaDates = stats?.datesAlpa || student.datesAlpa || [];
  const sakitDates = stats?.datesSakit || student.datesSakit || [];
  const izinDates = stats?.datesIzin || student.datesIzin || [];
  const dispenDates = stats?.datesDispen || student.datesDispen || [];
  const hasDateNotes = alpaDates.length > 0 || sakitDates.length > 0 || izinDates.length > 0 || dispenDates.length > 0;

  const handleGeneratePDF = () => {
    try {
      setIsGenerating(true);
      const doc = new jsPDF({
        orientation: 'p',
        unit: 'mm',
        format: 'a4'
      });

      const pageWidth = doc.internal.pageSize.getWidth(); // 210
      let currentY = 10;

      // 1. Kop Surat
      let hasKopImage = false;
      if (kopSuratBase64 && kopSuratBase64.length > 100) {
        try {
          doc.addImage(kopSuratBase64, 'JPEG', 14, 8, 182, 30);
          currentY = 41;
          hasKopImage = true;
        } catch {
          hasKopImage = false;
        }
      }

      if (!hasKopImage) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(13);
        doc.text('YAYASAN DARUL FALAH', pageWidth / 2, currentY, { align: 'center' });
        currentY += 5;
        doc.setFontSize(15);
        doc.text('SMA DARUL FALAH CIHAMPELAS', pageWidth / 2, currentY, { align: 'center' });
        currentY += 5;
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        doc.text('Jl. Raya Cihampelas No. 45, Cihampelas, Kab. Bandung Barat, Jawa Barat 40562', pageWidth / 2, currentY, { align: 'center' });
        currentY += 4;
        doc.text('Email: smadarulfalah@gmail.com | NPSN: 20227891', pageWidth / 2, currentY, { align: 'center' });
        currentY += 4;
        
        // Line
        doc.setLineWidth(0.8);
        doc.line(14, currentY, pageWidth - 14, currentY);
        doc.setLineWidth(0.2);
        doc.line(14, currentY + 0.8, pageWidth - 14, currentY + 0.8);
        currentY += 7;
      }

      // 2. Tanggal Surat & Nomor
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9.5);

      const rightX = pageWidth - 18;
      doc.text(`Cihampelas, ${tanggalSurat || todayStr}`, rightX, currentY, { align: 'right' });
      currentY += 5.5;

      const leftMargin = 18;
      doc.text('Nomor', leftMargin, currentY);
      doc.text(':', leftMargin + 18, currentY);
      doc.text(nomorSurat, leftMargin + 21, currentY);
      currentY += 4.5;

      doc.text('Lampiran', leftMargin, currentY);
      doc.text(':', leftMargin + 18, currentY);
      doc.text(lampiran, leftMargin + 21, currentY);
      currentY += 4.5;

      doc.text('Perihal', leftMargin, currentY);
      doc.text(':', leftMargin + 18, currentY);
      doc.setFont('helvetica', 'bold');
      doc.text(perihal, leftMargin + 21, currentY);
      doc.setFont('helvetica', 'normal');
      currentY += 8;

      // 3. Penerima Surat
      doc.setFont('helvetica', 'bold');
      doc.text('Yth. Bapak / Ibu / Wali Murid dari:', leftMargin, currentY);
      currentY += 5.5;

      const labelIndent = leftMargin + 5;
      const colonIndent = leftMargin + 30;
      const valueIndent = leftMargin + 33;

      doc.setFont('helvetica', 'bold');
      doc.text('Nama Siswa', labelIndent, currentY);
      doc.text(':', colonIndent, currentY);
      doc.text(student.name, valueIndent, currentY);
      currentY += 4.5;

      doc.setFont('helvetica', 'normal');
      doc.text('NISN / ID', labelIndent, currentY);
      doc.text(':', colonIndent, currentY);
      doc.text(student.nisn || '-', valueIndent, currentY);
      currentY += 4.5;

      doc.text('Kelas', labelIndent, currentY);
      doc.text(':', colonIndent, currentY);
      doc.text(student.class, valueIndent, currentY);
      currentY += 4.5;

      doc.text('Sekolah', labelIndent, currentY);
      doc.text(':', colonIndent, currentY);
      doc.text('SMA Darul Falah', valueIndent, currentY);
      currentY += 8;

      // 4. Pembuka
      doc.text('Dengan hormat,', leftMargin, currentY);
      currentY += 5;

      const bodyText = `Sehubungan dengan hasil evaluasi kedisiplinan dan rekapitulasi data presensi siswa di SMA Darul Falah, bersama ini kami sampaikan data rincian kehadiran dan catatan ketidakhadiran putra/putri Bapak/Ibu sebagai berikut:`;
      const splitBody = doc.splitTextToSize(bodyText, pageWidth - (leftMargin * 2));
      doc.text(splitBody, leftMargin, currentY);
      currentY += splitBody.length * 4.5 + 3;

      // 5. Tabel Rekapitulasi Presensi Siswa (Total Hadir & Rincian Ketidakhadiran)
      const tableData = [
        ['1', 'Total Hadir', `${hadirCount} Hari / Pertemuan`, 'Hadir mengikuti kegiatan belajar mengajar'],
        ['2', 'Ketidakhadiran - Sakit (S)', `${sakitCount} Hari`, sakitCount > 0 ? 'Keterangan sakit (surat dokter / ortu)' : 'Nihil'],
        ['3', 'Ketidakhadiran - Izin (I)', `${izinCount} Hari`, izinCount > 0 ? 'Keterangan izin resmi orang tua / wali' : 'Nihil'],
        ['4', 'Ketidakhadiran - Alpa (A)', `${alpaCount} Hari`, alpaCount > 0 ? 'TANPA KETERANGAN (PERLU PEMBINAAN)' : 'Nihil'],
      ];

      if (dispenCount > 0) {
        tableData.push(['5', 'Ketidakhadiran - Dispensasi (D)', `${dispenCount} Hari`, 'Dispensasi tugas / kegiatan resmi sekolah']);
      }

      const totalAbsRowText = `TOTAL KETIDAKHADIRAN (S + I + A${dispenCount > 0 ? ' + D' : ''})`;
      tableData.push(
        ['•', totalAbsRowText, `${totalTidakHadir} Hari`, alpaCount >= 3 ? 'KATEGORI KRITIS (Sangat Perlu Pembinaan)' : totalTidakHadir > 5 ? 'Perlu Perhatian & Kerjasama Orang Tua' : 'Catatan Kedisiplinan'],
        ['•', 'TOTAL HARI EFEKTIF BELAJAR', `${totalPertemuan} Pertemuan`, 'Jumlah seluruh sesi / hari presensi berjalan'],
        ['•', 'PERSENTASE KEHADIRAN', rateStr, rawRate < 80 || alpaCount >= 3 ? 'DI BAWAH STANDAR MINIMAL (< 80%)' : 'Memenuhi Standar Minimal']
      );

      autoTable(doc, {
        startY: currentY,
        margin: { left: leftMargin, right: leftMargin },
        head: [['No', 'Kategori Presensi / Keterangan', 'Jumlah', 'Catatan Evaluasi / Keterangan']],
        body: tableData,
        theme: 'grid',
        headStyles: { 
          fillColor: [240, 240, 240], 
          textColor: [0, 0, 0], 
          fontStyle: 'bold', 
          fontSize: 8.5,
          halign: 'center',
          valign: 'middle',
          lineWidth: 0.2,
          lineColor: [120, 120, 120]
        },
        bodyStyles: { 
          fontSize: 8.5, 
          cellPadding: 2.2,
          fillColor: [255, 255, 255],
          textColor: [0, 0, 0],
          lineWidth: 0.2,
          lineColor: [150, 150, 150]
        },
        alternateRowStyles: {
          fillColor: [255, 255, 255],
          textColor: [0, 0, 0]
        },
        columnStyles: {
          0: { cellWidth: 10, halign: 'center', fontStyle: 'bold', textColor: [0, 0, 0] },
          1: { cellWidth: 68, fontStyle: 'bold', textColor: [0, 0, 0] },
          2: { cellWidth: 34, halign: 'center', fontStyle: 'bold', textColor: [0, 0, 0] },
          3: { cellWidth: 62, textColor: [0, 0, 0] }
        }
      });

      currentY = (doc as any).lastAutoTable.finalY + 4;

      // Dates breakdown if available
      if (hasDateNotes) {
        doc.setFontSize(8);
        doc.setFont('helvetica', 'bold');
        doc.text('Rincian Tanggal Ketidakhadiran:', leftMargin, currentY);
        currentY += 3.5;

        doc.setFont('helvetica', 'normal');
        if (alpaDates.length > 0) {
          const alpaText = `• Alpa / Tanpa Keterangan (${alpaDates.length}x): ${alpaDates.join(', ')}`;
          const split = doc.splitTextToSize(alpaText, pageWidth - (leftMargin * 2));
          doc.text(split, leftMargin + 2, currentY);
          currentY += split.length * 3.5;
        }
        if (sakitDates.length > 0) {
          const sakitText = `• Sakit (${sakitDates.length}x): ${sakitDates.join(', ')}`;
          const split = doc.splitTextToSize(sakitText, pageWidth - (leftMargin * 2));
          doc.text(split, leftMargin + 2, currentY);
          currentY += split.length * 3.5;
        }
        if (izinDates.length > 0) {
          const izinText = `• Izin (${izinDates.length}x): ${izinDates.join(', ')}`;
          const split = doc.splitTextToSize(izinText, pageWidth - (leftMargin * 2));
          doc.text(split, leftMargin + 2, currentY);
          currentY += split.length * 3.5;
        }
        if (dispenDates.length > 0) {
          const dispenText = `• Dispensasi (${dispenDates.length}x): ${dispenDates.join(', ')}`;
          const split = doc.splitTextToSize(dispenText, pageWidth - (leftMargin * 2));
          doc.text(split, leftMargin + 2, currentY);
          currentY += split.length * 3.5;
        }
        currentY += 2;
      }

      // 6. Undangan Pertemuan
      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'normal');
      const textInv = `Guna membahas langkah pembinaan dan tindak lanjut kedisiplinan demi kelancaran proses belajar putra/putri Bapak/Ibu, kami mengharapkan kehadiran Bapak/Ibu pada:`;
      const splitInv = doc.splitTextToSize(textInv, pageWidth - (leftMargin * 2));
      doc.text(splitInv, leftMargin, currentY);
      currentY += splitInv.length * 4.5 + 3;

      const invLabelX = leftMargin + 5;
      const invColonX = leftMargin + 35;
      const invValX = leftMargin + 38;

      doc.setFont('helvetica', 'bold');
      doc.text('Hari / Tanggal', invLabelX, currentY);
      doc.text(':', invColonX, currentY);
      doc.text(hariTanggal || '...................., .... .................... 2026', invValX, currentY);
      currentY += 4.5;

      doc.text('Waktu', invLabelX, currentY);
      doc.text(':', invColonX, currentY);
      doc.text(waktu || '09.00 WIB - Selesai', invValX, currentY);
      currentY += 4.5;

      doc.text('Tempat', invLabelX, currentY);
      doc.text(':', invColonX, currentY);
      doc.text(tempat || 'Ruang Guru / BK SMA Darul Falah', invValX, currentY);
      currentY += 4.5;

      doc.text('Menemui', invLabelX, currentY);
      doc.text(':', invColonX, currentY);
      doc.text(menemui || 'Wali Kelas & Guru BK', invValX, currentY);
      currentY += 7;

      // Penutup
      doc.setFont('helvetica', 'normal');
      const textClose = `Mengingat pentingnya hal tersebut, kami sangat mengharapkan kehadiran Bapak/Ibu tepat pada waktunya. Atas perhatian, kesediaan waktu, dan kerja sama Bapak/Ibu, kami sampaikan terima kasih.`;
      const splitClose = doc.splitTextToSize(textClose, pageWidth - (leftMargin * 2));
      doc.text(splitClose, leftMargin, currentY);
      currentY += splitClose.length * 4.5 + 8;

      // 7. Tanda Tangan
      const leftSignerX = leftMargin + 8;
      const rightSignerX = pageWidth - leftMargin - 65;

      // Guard bottom page overflow
      if (currentY > 240) {
        doc.addPage();
        currentY = 25;
      }

      doc.setFont('helvetica', 'normal');
      doc.text('Mengetahui,', leftSignerX, currentY);
      doc.text('Wali Kelas / BK,', rightSignerX, currentY);
      currentY += 4.5;

      doc.setFont('helvetica', 'bold');
      doc.text('Kepala Sekolah', leftSignerX, currentY);
      doc.text(`Wali Kelas ${student.class}`, rightSignerX, currentY);
      currentY += 20;

      doc.text(namaKepsek || '(________________________)', leftSignerX, currentY);
      doc.text(namaWali || '(________________________)', rightSignerX, currentY);
      currentY += 4;

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(nipKepsek ? `NIY. ${nipKepsek}` : '', leftSignerX, currentY);
      doc.text(niyWali ? `NIY. ${niyWali}` : '', rightSignerX, currentY);

      doc.save(`Surat_Pemanggilan_${student.name.replace(/\s+/g, '_')}_${student.class}.pdf`);

      if (showToast) showToast(`Surat Pemanggilan untuk ${student.name} berhasil diunduh.`, 'success');
      onClose();
    } catch (error) {
      console.error('Failed to generate call letter PDF:', error);
      if (showToast) showToast('Gagal mengunduh Surat Pemanggilan PDF.', 'error');
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-fade-in overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-xl overflow-hidden my-8">
        
        {/* Header */}
        <div className="px-6 py-4 bg-gradient-to-r from-rose-700 to-rose-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-white/10 rounded-xl">
              <Printer className="w-5 h-5 text-rose-200" />
            </div>
            <div>
              <h3 className="font-bold text-base leading-snug">Cetak Surat Pemanggilan Orang Tua</h3>
              <p className="text-xs text-rose-200/90">Peringatan Dini & Rekapitulasi Presensi Siswa</p>
            </div>
          </div>
          <button 
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-white/10 text-white/80 hover:text-white transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Student Presence & Absence Breakdown Card */}
        <div className="p-4 bg-rose-50/90 border-b border-rose-100 flex flex-col gap-2.5">
          <div className="flex items-start justify-between gap-2">
            <div className="flex items-center gap-2">
              <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0" />
              <div>
                <p className="font-bold text-sm text-rose-950">{student.name} ({student.class})</p>
                <p className="text-xs text-rose-800 font-mono">NISN: <span className="font-bold">{student.nisn || '-'}</span></p>
              </div>
            </div>
            <span className={`px-2.5 py-1 rounded-full text-xs font-black border ${
              rawRate >= 80 ? 'bg-emerald-100 text-emerald-800 border-emerald-300' : 'bg-rose-100 text-rose-800 border-rose-300'
            }`}>
              Kehadiran: {rateStr}
            </span>
          </div>

          {/* Detailed presence and absence pills */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center text-xs">
            <div className="bg-white/90 p-2 rounded-xl border border-emerald-200 shadow-2xs">
              <span className="text-[10px] text-emerald-800 font-semibold block uppercase">Total Hadir</span>
              <span className="font-black text-sm text-emerald-700">{hadirCount} <span className="text-[10px] font-medium">Hari</span></span>
            </div>
            <div className="bg-white/90 p-2 rounded-xl border border-amber-200 shadow-2xs">
              <span className="text-[10px] text-amber-800 font-semibold block uppercase">Sakit & Izin</span>
              <span className="font-bold text-sm text-amber-700">{sakitCount} S <span className="text-slate-400">/</span> {izinCount} I</span>
            </div>
            <div className="bg-white/90 p-2 rounded-xl border border-rose-200 shadow-2xs">
              <span className="text-[10px] text-rose-800 font-semibold block uppercase">Alpa & Dispen</span>
              <span className="font-bold text-sm text-rose-700">{alpaCount} A <span className="text-slate-400">/</span> {dispenCount} D</span>
            </div>
            <div className="bg-rose-100/90 p-2 rounded-xl border border-rose-300 shadow-2xs">
              <span className="text-[10px] text-rose-900 font-bold block uppercase">Total Tidak Hadir</span>
              <span className="font-black text-sm text-rose-900">{totalTidakHadir} <span className="text-[10px] font-medium">Hari</span></span>
            </div>
          </div>

          {/* Dates notes if available */}
          {hasDateNotes && (
            <div className="bg-white/80 p-2.5 rounded-xl border border-rose-200/60 text-[11px] text-rose-950 space-y-1">
              <div className="font-bold flex items-center gap-1 text-rose-900">
                <FileText className="w-3.5 h-3.5 text-rose-600" />
                <span>Rincian Tanggal Ketidakhadiran:</span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-2 gap-y-0.5 text-[10px] text-slate-700">
                {alpaDates.length > 0 && (
                  <p><span className="font-bold text-rose-700">Alpa ({alpaDates.length}):</span> {alpaDates.join(', ')}</p>
                )}
                {sakitDates.length > 0 && (
                  <p><span className="font-bold text-amber-700">Sakit ({sakitDates.length}):</span> {sakitDates.join(', ')}</p>
                )}
                {izinDates.length > 0 && (
                  <p><span className="font-bold text-blue-700">Izin ({izinDates.length}):</span> {izinDates.join(', ')}</p>
                )}
                {dispenDates.length > 0 && (
                  <p><span className="font-bold text-purple-700">Dispen ({dispenDates.length}):</span> {dispenDates.join(', ')}</p>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Body Form */}
        <div className="p-6 space-y-4 text-xs text-slate-700 max-h-[60vh] overflow-y-auto">
          
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label className="block font-bold text-slate-700 mb-1">Nomor Surat</label>
              <input 
                type="text" 
                value={nomorSurat}
                onChange={e => setNomorSurat(e.target.value)}
                className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-mono font-semibold"
              />
            </div>
            <div>
              <label className="block font-bold text-slate-700 mb-1">Tanggal Surat</label>
              <input 
                type="text" 
                value={tanggalSurat}
                onChange={e => setTanggalSurat(e.target.value)}
                placeholder="misal: 28 Agustus 2026"
                className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-medium"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label className="block font-bold text-slate-700 mb-1">Lampiran</label>
              <input 
                type="text" 
                value={lampiran}
                onChange={e => setLampiran(e.target.value)}
                className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-medium"
              />
            </div>
            <div>
              <label className="block font-bold text-slate-700 mb-1">Perihal</label>
              <input 
                type="text" 
                value={perihal}
                onChange={e => setPerihal(e.target.value)}
                className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-medium"
              />
            </div>
          </div>

          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-3">
            <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider flex items-center gap-1.5">
              <Calendar className="w-4 h-4 text-rose-600" /> Jadwal Undangan Pemanggilan
            </h4>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block font-semibold text-slate-600 mb-1">Hari & Tanggal Pertemuan</label>
                <input 
                  type="text" 
                  value={hariTanggal}
                  onChange={e => setHariTanggal(e.target.value)}
                  placeholder="Contoh: Senin, 31 Agustus 2026"
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-semibold"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-600 mb-1">Waktu Pertemuan</label>
                <input 
                  type="text" 
                  value={waktu}
                  onChange={e => setWaktu(e.target.value)}
                  placeholder="09.00 WIB - Selesai"
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-medium"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block font-semibold text-slate-600 mb-1">Tempat Pertemuan</label>
                <input 
                  type="text" 
                  value={tempat}
                  onChange={e => setTempat(e.target.value)}
                  placeholder="Ruang Guru / BK"
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-medium"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-600 mb-1">Menemui</label>
                <input 
                  type="text" 
                  value={menemui}
                  onChange={e => setMenemui(e.target.value)}
                  placeholder="Wali Kelas & Guru BK"
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all font-medium"
                />
              </div>
            </div>
          </div>

          <div className="p-3.5 bg-emerald-50/80 border border-emerald-200/80 rounded-xl space-y-3">
            <h4 className="font-bold text-emerald-900 text-xs uppercase tracking-wider flex items-center gap-1.5">
              <User className="w-4 h-4 text-emerald-700" /> Penandatangan Surat (Wali Kelas & Kepala Sekolah)
            </h4>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block font-semibold text-emerald-800 mb-1">Nama Wali Kelas {student.class}</label>
                <input 
                  type="text" 
                  value={namaWali}
                  onChange={e => setNamaWali(e.target.value)}
                  placeholder="Contoh: Agan Parta,S.Kom.,Gr."
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all font-medium"
                />
              </div>
              <div>
                <label className="block font-semibold text-emerald-800 mb-1">NIY Wali Kelas</label>
                <input 
                  type="text" 
                  value={niyWali}
                  onChange={e => setNiyWali(e.target.value)}
                  placeholder="Contoh: 198203152009021003"
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all font-medium"
                />
              </div>
            </div>

            <div>
              <label className="block font-semibold text-emerald-800 mb-1">Kepala Sekolah</label>
              <input 
                type="text" 
                value={namaKepsek}
                onChange={e => setNamaKepsek(e.target.value)}
                placeholder="Nama Kepala Sekolah"
                className="w-full p-2.5 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all font-medium"
              />
            </div>
          </div>

        </div>

        {/* Footer actions */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2.5">
          <button 
            type="button"
            onClick={onClose}
            className="px-4 py-2.5 bg-white border border-slate-300 text-slate-700 rounded-xl font-bold hover:bg-slate-100 transition-colors"
          >
            Batal
          </button>
          <button 
            type="button"
            onClick={handleGeneratePDF}
            disabled={isGenerating}
            className="px-5 py-2.5 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white rounded-xl font-bold flex items-center gap-2 transition-all shadow-md active:scale-95"
          >
            {isGenerating ? (
              <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
            ) : (
              <Printer className="w-4 h-4" />
            )}
            {isGenerating ? 'Mengekspor PDF...' : 'Unduh Surat Pemanggilan (PDF)'}
          </button>
        </div>

      </div>
    </div>
  );
};

export default SuratPanggilanModal;

