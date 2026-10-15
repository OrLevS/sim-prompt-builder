# כותב פרומפטים מוגמרים של סימולציה: טקסט נקי להעתקה + Word לקריאה.
# הרצה: python tools/render_sim.py <תיקיית הסימולציה>   (בתיקייה צריך להיות simulation.json)
import io
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_zip as mz  # noqa: E402

folder = os.path.abspath(sys.argv[1])
js = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'render_sim.js')
res = subprocess.run(['node', js, os.path.join(folder, 'simulation.json')], capture_output=True)
if res.stderr:
    sys.stderr.write(res.stderr.decode('utf-8'))
items = json.loads(res.stdout.decode('utf-8'))
sim = json.load(io.open(os.path.join(folder, 'simulation.json'), encoding='utf-8'))
title = sim['exercise']['title'] + ' — ' + ', '.join(s['name'] for s in sim['students'][:sim['settings']['count']])

txt_dir = os.path.join(folder, 'טקסט נקי להעתקה')
os.makedirs(txt_dir, exist_ok=True)
for it in items:
    with io.open(os.path.join(txt_dir, it['file'] + '.txt'), 'w', encoding='utf-8-sig', newline='\r\n') as f:
        f.write(it['text'])

# Word אחד לקריאה, עם כל הפרומפטים לפי הסדר
doc = mz.new_doc()
mz.T(doc, title, size=20, bold=True, color='0B2545', space_after=2)
mode = 'ידני (המורה כותבת [מעבר לשלב 2])' if sim['settings']['mode'] == 'manual' else 'אוטומטי'
mz.T(doc, 'מעבר שלבים: ' + mode, size=11, color='1D5FA8', space_after=2)
mz.T(doc, 'הקובץ הזה מעוצב לקריאה. להעתקה למערכת: התיקייה "טקסט נקי להעתקה", קובץ לכל תיבה.', size=10, color='667085', space_after=10)
for it in items:
    mz.T(doc, it['file'], size=17, bold=True, color='0B2545', space_before=18, space_after=2)
    mz.T(doc, 'לאן מדביקים: ' + it['box'], size=11, bold=True, color='1D5FA8', space_after=8)
    mz.write_prompt(doc, it['text'])
doc.save(os.path.join(folder, 'כל הפרומפטים - לקריאה.docx'))
print('ok', len(items))
