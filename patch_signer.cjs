const fs = require('fs');
let code = fs.readFileSync('src/ReportsView.tsx', 'utf8');

code = code.replace(/relevantSession = classSessions\.find\(s => \s*s\.date === selectedDailyDate && \s*s\.className === selectedClass\s*\);/g, 
`const dailySess = classSessions.filter(s => s.date === selectedDailyDate && s.className === selectedClass);
          relevantSession = dailySess[dailySess.length - 1];`);

fs.writeFileSync('src/ReportsView.tsx', code);
console.log('patched right signer');
