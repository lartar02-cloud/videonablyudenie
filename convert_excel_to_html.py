import pandas as pd
from pathlib import Path

excel_file = 'journal.xlsx'
output_html = 'index.html'
images_dir = 'Images'

# ❗ Листы, которые НЕ попадают в HTML
EXCLUDED_SHEETS = {'СпрСобытий'}

needed_columns = [
    'День нед', 'Дата', 'Время',
    'Событие', 'Объект', 'Груз',
    'Кол-во', 'Примечание'
]

xl = pd.ExcelFile(excel_file)
sheets = [s for s in xl.sheet_names if s not in EXCLUDED_SHEETS]

html = '''
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Журнал видеонаблюдения</title>
<style>
body { font-family: Arial, sans-serif; background:#f5f5f5; }
.tabs { margin-bottom: 10px; }
.tablink {
  padding: 8px 14px;
  border: none;
  background: #ddd;
  cursor: pointer;
  margin-right: 4px;
}
.tablink.active { background: #444; color: white; }

.week-controls {
  margin: 15px 0;
  display: flex;
  align-items: center;
  gap: 10px;
}

table {
  border-collapse: collapse;
  width: 100%;
  background: white;
}
th, td {
  border: 1px solid #ccc;
  padding: 6px;
}
th { background: #eee; }

img.preview {
  max-width: 140px;
  cursor: pointer;
}
</style>
</head>
<body>

<h2>Журнал видеонаблюдения</h2>

<div class="tabs">
'''

# ---------- вкладки ----------
for sheet in sheets:
    html += f'<button class="tablink" onclick="openTab(event, \'{sheet}\')">{sheet}</button>'

html += '''
</div>

<div class="week-controls">
  <button onclick="changeWeek(-1)">⏮ Предыдущая неделя</button>
  <strong id="weekLabel"></strong>
  <button onclick="changeWeek(1)">Следующая неделя ⏭</button>
</div>
'''

# ---------- контент ----------
for sheet in sheets:
    df = xl.parse(sheet).dropna(how='all').fillna('').reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    if 'Дата' not in df.columns:
        continue

    # ISO-дата ТОЛЬКО для JS-фильтра
    df['__date_iso'] = pd.to_datetime(
        df['Дата'], format='%d.%m.%Y', errors='coerce'
    ).dt.strftime('%Y-%m-%d')

    # Красивая дата для таблицы
    df['Дата'] = pd.to_datetime(
        df['Дата'], format='%d.%m.%Y', errors='coerce'
    ).dt.strftime('%d.%m.%Y')

    # ---------- фото ----------
    def make_photo(row):
        link = row.get('Ссылка на фото', '')
        if not link:
            return '—'

        filename = Path(str(link)).name
        src = f'{images_dir}/{sheet}/{filename}'.replace('\\', '/')
        return f'<img class="preview" src="{src}" onclick="openImage(this.src)">'

    df['Фото'] = df.apply(make_photo, axis=1)

    final_cols = [c for c in needed_columns if c in df.columns]
    df = df[final_cols + ['Фото', '__date_iso']]

    html += f'<div id="{sheet}" class="tabcontent" style="display:none">'
    html += f'<h3>{sheet}</h3>'
    html += '<table><thead><tr>'

    for c in df.columns:
        if c != '__date_iso':
            html += f'<th>{c}</th>'
    html += '</tr></thead><tbody>'

    for _, r in df.iterrows():
        html += f'<tr data-date="{r["__date_iso"]}">'
        for c in df.columns:
            if c != '__date_iso':
                html += f'<td>{r[c]}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'

# ---------- JS ----------
html += '''
<script>
let currentMonday = getMonday(new Date());

function getMonday(d) {
  d = new Date(d);
  let day = d.getDay() || 7;
  if (day !== 1) d.setDate(d.getDate() - (day - 1));
  d.setHours(0,0,0,0);
  return d;
}

function formatRU(d) {
  return d.toLocaleDateString('ru-RU', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
}

function renderWeek() {
  let monday = new Date(currentMonday);
  let sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  document.getElementById('weekLabel').innerText =
    formatRU(monday) + ' — ' + formatRU(sunday);

  document.querySelectorAll('tr[data-date]').forEach(row => {
    let d = new Date(row.dataset.date);
    row.style.display = (d >= monday && d <= sunday) ? '' : 'none';
  });
}

function changeWeek(n) {
  currentMonday.setDate(currentMonday.getDate() + n * 7);
  renderWeek();
}

function openTab(evt, name) {
  document.querySelectorAll('.tabcontent').forEach(t => t.style.display = 'none');
  document.querySelectorAll('.tablink').forEach(b => b.classList.remove('active'));
  document.getElementById(name).style.display = 'block';
  evt.currentTarget.classList.add('active');
  renderWeek();
}

function openImage(src) {
  window.open(src, '_blank');
}

document.querySelector('.tablink').click();
</script>

</body>
</html>
'''

Path(output_html).write_text(html, encoding='utf-8')
print('ГОТОВО ✅ Журнал сформирован')
