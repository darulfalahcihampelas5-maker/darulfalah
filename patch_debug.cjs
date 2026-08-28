const fs = require('fs');
let code = fs.readFileSync('src/ReportsView.tsx', 'utf8');

code = code.replace(/let dailyPiketName = '';/g, 
"let dailyPiketName = '';\n        console.log('RELEVANT SESSION:', relevantSession);");

fs.writeFileSync('src/ReportsView.tsx', code);
console.log('patched debug');
