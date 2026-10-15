// מרכיב פרומפטים מוגמרים מקובץ סימולציה (simulation.json), באותו קוד של האתר.
// הרצה: node tools/render_sim.js <simulation.json>  → JSON על stdout
const fs = require('fs');
const path = require('path');
global.window = {};
require(path.join(__dirname, '..', 'js', 'personas.js'));
require(path.join(__dirname, '..', 'js', 'prompts.js'));
const P = window.PROMPTS;

const st = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const n = st.settings.count;
const out = [
  { file: '1 מנהל השיח', box: 'פרומפט מנהל שיח', text: P.strip(P.orchestrator(st)) },
  { file: '2 המנחה', box: 'פרומפט מנחה', text: P.strip(P.mentor(st)) },
  { file: '3 הנחיות משותפות לכל התלמידים', box: 'פרומפט על לכל תלמידי AI', text: P.strip(P.shared(st)) }
];
for (let i = 0; i < n; i++) {
  const s = st.students[i];
  out.push({ file: (4 + i) + ' דמות - ' + s.name, box: 'פרומפט הדמות של ' + s.name + ' (תווית ' + s.label + ')', text: P.strip(P.character(st, i)) });
}
const left = out.filter(o => o.text.includes('⟦')).map(o => o.file);
if (left.length) process.stderr.write('חסרים שדות ב: ' + left.join(', ') + '\n');
process.stdout.write(JSON.stringify(out));
