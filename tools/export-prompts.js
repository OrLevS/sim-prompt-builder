// מייצא את הפרומפטים הקבועים (עם סימוני החלקים המשתנים) ל-JSON, לבניית קובץ ה-ZIP.
// הרצה: node tools/export-prompts.js > tools/out/prompts.json
const path = require('path');
global.window = {};
require(path.join(__dirname, '..', 'js', 'personas.js'));
require(path.join(__dirname, '..', 'js', 'prompts.js'));
const P = window.PROMPTS;
const FIELDS = window.PERSONA_FIELDS;

function blankStudent(g) {
  return { personaId: 'custom', name: '', label: '', g, error: '', opening: '', persona: {} };
}
function ctx(mode, count, g) {
  const students = [];
  for (let i = 0; i < 5; i++) students.push(blankStudent(g || 'f'));
  return {
    settings: { mode, teacher: 'f', count: count || 1 },
    exercise: { title: '', grade: '', learned: '', task: '', figure: '', notYet: '', centralIdea: '', solution: '' },
    stages: [{ name: '', goal: '' }],
    students
  };
}
// מקטע מתוך פרומפט, מכותרת ## ועד הכותרת ## הבאה
function section(text, title) {
  const a = text.indexOf('## ' + title);
  if (a < 0) return '';
  const b = text.indexOf('\n## ', a + 3);
  return text.slice(a, b < 0 ? undefined : b).trim();
}

const out = {};
out.orchestrator_manual = P.orchestrator(ctx('manual'));
out.orchestrator_auto = P.orchestrator(ctx('auto'));
out.mentor = P.mentor(ctx('manual'));
out.shared_manual = P.shared(ctx('manual'));
out.shared_auto = P.shared(ctx('auto'));
out.character_f = P.character(ctx('manual', 1, 'f'), 0);
out.character_m = P.character(ctx('manual', 1, 'm'), 0);
out.helper = P.errorHelper(ctx('manual'), 0);

const g2 = ctx('manual', 2);
out.group_shared = section(P.shared(g2), 'כשיש עוד תלמידים');
out.group_orch = section(P.orchestrator(g2), 'איזה תלמיד/ה');

out.personas = window.PERSONAS.map(p => {
  const c = ctx('manual', 1, p.g);
  c.plainPersona = true;
  const st = c.students[0];
  st.personaId = p.id; st.name = p.name; st.label = p.label;
  FIELDS.forEach(([k]) => { st.persona[k] = P.gender(p[k], p.g); });
  return { id: p.id, name: p.name, label: p.label, g: p.g, tagline: P.gender(p.tagline, p.g), text: P.character(c, 0) };
});

process.stdout.write(JSON.stringify(out));
