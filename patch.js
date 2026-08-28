const fs = require('fs');
let code = fs.readFileSync('src/ReportsView.tsx', 'utf8');

const target = `  const getRightSignerDetails = () => {
    if (reportType === 'daily') {
      if (selectedClass !== 'all') {
        const dailySession = classSessions.find(s => 
          s.date === selectedDailyDate && 
          s.className === selectedClass
        );

        let dailyPiketName = '';
        if (dailySession) {
          const ds = dailySession as { piketTeacherName?: string; recordedByRole?: string; recordedByTeacherName?: string };
          dailyPiketName = ds.piketTeacherName || 
                           (ds.recordedByRole === 'Petugas Piket' ? ds.recordedByTeacherName : '') ||
                           profileData?.namaGuruMapel || 
                           'Petugas Piket';
        } else {
          dailyPiketName = profileData?.namaGuruMapel || 'Petugas Piket';
        }

        const matchedPiket = teachers.find(t => t.name.trim().toLowerCase() === dailyPiketName.trim().toLowerCase());
        
        return {
          label: 'Petugas Piket',
          name: dailyPiketName,
          niy: matchedPiket?.niy ? \`NIY. \${matchedPiket.niy}\` : ''
        };
      } else {
        const koordinatorName = profileData?.namaGuruMapel || profileData?.fullname || 'Koordinator Piket';
        const koordinatorNiy = profileData?.nipGuruMapel ? \`NIY. \${profileData.nipGuruMapel}\` : '';
        return {
          label: 'Koordinator Piket',
          name: koordinatorName,
          niy: koordinatorNiy
        };
      }
    }`;

const replacement = `  const getRightSignerDetails = () => {
    if (reportType === 'daily' || reportType === 'custom') {
      if (selectedClass !== 'all') {
        let relevantSession = null;
        if (reportType === 'daily') {
          relevantSession = classSessions.find(s => 
            s.date === selectedDailyDate && 
            s.className === selectedClass
          );
        } else {
          relevantSession = classSessions.find(s => 
            s.className === selectedClass && 
            (!startDate || s.date >= startDate) && 
            (!endDate || s.date <= endDate)
          );
        }

        let dailyPiketName = '';
        if (relevantSession) {
          const ds = relevantSession as { piketTeacherName?: string; recordedByRole?: string; recordedByTeacherName?: string };
          dailyPiketName = ds.piketTeacherName || 
                           (ds.recordedByRole === 'Petugas Piket' ? ds.recordedByTeacherName : '') ||
                           profileData?.namaGuruMapel || 
                           'Petugas Piket';
        } else {
          dailyPiketName = profileData?.namaGuruMapel || 'Petugas Piket';
        }

        const matchedPiket = teachers.find(t => t.name.trim().toLowerCase() === dailyPiketName.trim().toLowerCase());
        
        return {
          label: 'Petugas Piket',
          name: dailyPiketName,
          niy: (matchedPiket?.niy && matchedPiket.niy.trim() !== '-') ? \`NIY. \${matchedPiket.niy}\` : ''
        };
      } else {
        const koordinatorName = profileData?.namaGuruMapel || profileData?.fullname || 'Koordinator Piket';
        const koordinatorNiy = profileData?.nipGuruMapel ? \`NIY. \${profileData.nipGuruMapel}\` : '';
        return {
          label: 'Koordinator Piket',
          name: koordinatorName,
          niy: koordinatorNiy
        };
      }
    }`;

if (code.includes(target)) {
    code = code.replace(target, replacement);
    fs.writeFileSync('src/ReportsView.tsx', code);
    console.log('patched');
} else {
    console.log('Target not found in file');
}
