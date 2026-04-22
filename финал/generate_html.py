import json, os

# Load справочник data
with open('tabel_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

employees = data['employees']
sections  = data['sections']
objects   = data['objects']

# Load template
with open('template_v2.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Inject data
html = html.replace('%%EMPLOYEES%%', json.dumps(employees, ensure_ascii=False))
html = html.replace('%%SECTIONS%%',  json.dumps(sections,  ensure_ascii=False))
html = html.replace('%%OBJECTS%%',   json.dumps(objects,   ensure_ascii=False))

# Write output
out = 'Табель_2026.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)

kb = os.path.getsize(out) // 1024
print(f'OK: {out}  ({kb} KB)')
print(f'   Spivrobitnikiv: {len(employees)}')
print(f'   Rozdiliv:       {len(sections)}')
print(f'   Obiektiv:       {len(objects)}')
