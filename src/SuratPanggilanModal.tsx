/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { useState } from 'react';
import { X, Printer, User, Calendar, AlertTriangle } from 'lucide-react';
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
}

export interface StudentStatsForCall {
  hadir: number;
  sakit: number;
  izin: number;
  alpa: number;
  dispen?: number;
  total: number;
  rate: number;
  datesAlpa?: string[];
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
          doc.addImage(kopSuratBase64, 'JPEG', 12, 8, 186, 32);
          currentY = 43;
          hasKopImage = true;
        } catch {
          hasKopImage = false;
        }
      }

      if (!hasKopImage) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(14);
        doc.text('YAYASAN DARUL FALAH', pageWidth / 2, currentY, { align: 'center' });
        currentY += 6;
        doc.setFontSize(16);
        doc.text('SMA DARUL FALAH CIHAMPELAS', pageWidth / 2, currentY, { align: 'center' });
        currentY += 5;
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.text('Jl. Raya Cihampelas No. 45, Cihampelas, Kab. Bandung Barat, Jawa Barat 40562', pageWidth / 2, currentY, { align: 'center' });
        currentY += 4;
        doc.text('Email: smadarulfalah@gmail.com | NPSN: 20227891', pageWidth / 2, currentY, { align: 'center' });
        currentY += 4;
        
        // Line
        doc.setLineWidth(0.8);
        doc.line(15, currentY, pageWidth - 15, currentY);
        doc.setLineWidth(0.2);
        doc.line(15, currentY + 1, pageWidth - 15, currentY + 1);
        currentY += 8;
      }

      // 2. Tanggal Surat & Nomor
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(10);

      const rightX = pageWidth - 20;
      doc.text(`Cihampelas, ${tanggalSurat || todayStr}`, rightX, currentY, { align: 'right' });
      currentY += 6;

      doc.text(`Nomor     : ${nomorSurat}`, 20, currentY);
      currentY += 5;
      doc.text(`Lampiran  : ${lampiran}`, 20, currentY);
      currentY += 5;
      doc.text(`Perihal   : ${perihal}`, 20, currentY);
      currentY += 10;

      // 3. Penerima Surat
      doc.setFont('helvetica', 'bold');
      doc.text('Yth. Orang Tua / Wali Murid dari:', 20, currentY);
      currentY += 6;

      doc.setFont('helvetica', 'bold');
      doc.text(`Nama Siswa  : ${student.name}`, 25, currentY);
      currentY += 5;
      doc.setFont('helvetica', 'normal');
      doc.text(`NISN / ID   : ${student.nisn || '-'}`, 25, currentY);
      currentY += 5;
      doc.text(`Kelas       : ${student.class}`, 25, currentY);
      currentY += 5;
      doc.text(`Sekolah     : SMA Darul Falah`, 25, currentY);
      currentY += 10;

      // 4. Pembuka
      doc.text('Dengan hormat,', 20, currentY);
      currentY += 6;

      const bodyText = `Sehubungan dengan hasil evaluasi kedisiplinan dan rekapitulasi presensi harian siswa di SMA Darul Falah, bersama ini kami sampaikan laporan catatan ketidakhadiran putra/putri Bapak/Ibu sebagai berikut:`;
      const splitBody = doc.splitTextToSize(bodyText, pageWidth - 40);
      doc.text(splitBody, 20, currentY);
      currentY += splitBody.length * 5 + 4;

      // 5. Tabel Rekapitulasi Presensi Siswa
      const alpaCount = stats?.alpa || 0;
      const sakitCount = stats?.sakit || 0;
      const izinCount = stats?.izin || 0;
      const hadirCount = stats?.hadir || 0;
      const rateStr = stats ? `${stats.rate}%` : '-';

      const tableData = [
        ['Hadir', `${hadirCount} Pertemuan`, 'Sesuai Jadwal'],
        ['Izin', `${izinCount} Hari`, 'Dengan Surat Izin'],
        ['Sakit', `${sakitCount} Hari`, 'Dengan Surat Dokter/Ortu'],
        ['Tanpa Keterangan (Alpa)', `${alpaCount} Hari`, 'PERLU PEMBINAAN KHUSUS'],
        ['Persentase Kehadiran', rateStr, alpaCount >= 3 || (stats && stats.rate < 80) ? 'DI BAWAH STANDAR (KRITIS)' : 'Cukup']
      ];

      autoTable(doc, {
        startY: currentY,
        margin: { left: 20, right: 20 },
        head: [['Kategori Presensi', 'Jumlah', 'Keterangan Evaluasi']],
        body: tableData,
        theme: 'grid',
        headStyles: { fillColor: [15, 118, 110], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 9 },
        bodyStyles: { fontSize: 9, cellPadding: 2.5 },
        columnStyles: {
          0: { cellWidth: 55, fontStyle: 'bold' },
          1: { cellWidth: 45, halign: 'center' },
          2: { cellWidth: 70 }
        },
        didParseCell: function(data) {
          if (data.row.index === 3) {
            data.cell.styles.fillColor = [254, 242, 242];
            data.cell.styles.textColor = [185, 28, 28];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      });

      currentY = (doc as any).lastAutoTable.finalY + 8;

      if (stats?.datesAlpa && stats.datesAlpa.length > 0) {
        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'italic');
        doc.text(`* Catatan Tanggal Alpa: ${stats.datesAlpa.join(', ')}`, 20, currentY);
        currentY += 6;
      }

      // 6. Undangan Pertemuan
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      const textInv = `Guna membahas langkah pembinaan dan solusi bersama demi kebaikan proses belajar putra/putri Bapak/Ibu, kami mengharapkan kehadiran Bapak/Ibu pada:`;
      const splitInv = doc.splitTextToSize(textInv, pageWidth - 40);
      doc.text(splitInv, 20, currentY);
      currentY += splitInv.length * 5 + 4;

      doc.setFont('helvetica', 'bold');
      doc.text(`Hari / Tanggal  : ${hariTanggal || '...................., .... .................... 2026'}`, 25, currentY); currentY += 5;
      doc.text(`Waktu           : ${waktu || '09.00 WIB - Selesai'}`, 25, currentY); currentY += 5;
      doc.text(`Tempat          : ${tempat || 'Ruang Guru / BK SMA Darul Falah'}`, 25, currentY); currentY += 5;
      doc.text(`Menemui         : ${menemui || 'Wali Kelas & Guru BK'}`, 25, currentY); currentY += 8;

      // Penutup
      doc.setFont('helvetica', 'normal');
      const textClose = `Mengingat pentingnya hal tersebut, kami sangat mengharapkan kehadiran Bapak/Ibu tepat pada waktunya. Atas perhatian, kesediaan, dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.`;
      const splitClose = doc.splitTextToSize(textClose, pageWidth - 40);
      doc.text(splitClose, 20, currentY);
      currentY += splitClose.length * 5 + 12;

      // 7. Tanda Tangan
      const leftSignerX = 25;
      const rightSignerX = 130;

      // Guard bottom page overflow
      if (currentY > 240) {
        doc.addPage();
        currentY = 25;
      }

      doc.setFont('helvetica', 'normal');
      doc.text('Mengetahui,', leftSignerX, currentY);
      doc.text('Wali Kelas / BK,', rightSignerX, currentY);
      currentY += 5;

      doc.setFont('helvetica', 'bold');
      doc.text('Kepala Sekolah', leftSignerX, currentY);
      doc.text(`Wali Kelas ${student.class}`, rightSignerX, currentY);
      currentY += 22;

      doc.text(namaKepsek || '(________________________)', leftSignerX, currentY);
      doc.text(namaWali || '(________________________)', rightSignerX, currentY);
      currentY += 4;

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8.5);
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
              <p className="text-xs text-rose-200/90">Peringatan Dini & Pembinaan Presensi Siswa</p>
            </div>
          </div>
          <button 
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-white/10 text-white/80 hover:text-white transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Student Info Alert */}
        <div className="p-4 bg-rose-50/80 border-b border-rose-100 flex items-start gap-3">
          <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
          <div className="text-xs text-rose-900">
            <p className="font-bold text-sm text-rose-950 mb-0.5">{student.name} ({student.class})</p>
            <p className="font-medium text-rose-800">
              NISN: <span className="font-bold">{student.nisn || '-'}</span> | 
              Alpa: <span className="font-bold text-rose-700">{stats?.alpa || 0} Hari</span> | 
              Kehadiran: <span className="font-bold text-rose-700">{stats?.rate || 0}%</span>
            </p>
          </div>
        </div>

        {/* Body Form */}
        <div className="p-6 space-y-4 text-xs text-slate-700 max-h-[65vh] overflow-y-auto">
          
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

          <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-3">
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

          <div className="p-3 bg-emerald-50/80 border border-emerald-200/80 rounded-xl space-y-3">
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
