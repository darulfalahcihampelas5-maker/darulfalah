const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

code = code.replace(/<option\s+key=\{t\.id\}\s+value=\{t\.id\}\s+disabled=\{\!isDutyDay\}\s+className=\{\!isDutyDay \? "text-slate-400 bg-slate-100" : "text-slate-800 font-medium"\}\s*>\s*\{t\.name\}\{t\.dutyDay && t\.dutyDay !== '-' \? \\\` \\(Piket: \\\$\\{t\\.dutyDay\\}\\)\\\` : ' \\(Belum ada jadwal piket\\)'\} \{\!isDutyDay \? ' - Bukan Hari Piket' : ''\}\s*<\/option>/g, 
"<option key={t.id} value={t.id} className={isDutyDay ? 'text-emerald-700 font-bold bg-emerald-50' : 'text-slate-700'}>" +
"{t.name}{t.dutyDay && t.dutyDay !== '-' ? ` (Piket: ${t.dutyDay})` : ' (Belum ada jadwal piket)'} {isDutyDay ? ' - ✨ JADWAL HARI INI ✨' : ''}" + 
"</option>");

fs.writeFileSync('src/App.tsx', code);
console.log('patched dropdown');
