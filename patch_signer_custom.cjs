const fs = require('fs');
let code = fs.readFileSync('src/ReportsView.tsx', 'utf8');

code = code.replace(/relevantSession = classSessions\.find\(s => \s*s\.className === selectedClass && \s*\(\!startDate \|\| s\.date >= startDate\) && \s*\(\!endDate \|\| s\.date <= endDate\)\s*\);/g, 
`const customSess = classSessions.filter(s => s.className === selectedClass && (!startDate || s.date >= startDate) && (!endDate || s.date <= endDate));
          relevantSession = customSess[customSess.length - 1];`);

fs.writeFileSync('src/ReportsView.tsx', code);
console.log('patched custom signer');
