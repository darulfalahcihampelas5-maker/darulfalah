const fs = require('fs');
let lines = fs.readFileSync('src/ReportsView.tsx', 'utf8').split('\n');

const startIdx = lines.findIndex(l => l === '  const getRightSignerDetails = () => {');
const endIdx = startIdx + 34; // line 281 + 34 = 315

if (startIdx !== -1) {
  lines.splice(startIdx, endIdx - startIdx + 1, 
`  const getRightSignerDetails = () => {
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
    }`);
  
  fs.writeFileSync('src/ReportsView.tsx', lines.join('\n'));
  console.log('patched successfully');
} else {
  console.log('not found');
}
