import pandas as pd
from pathlib import Path

excel_file = 'journal.xlsx'
output_html = 'index.html'
images_dir = 'Images'

EXCLUDED_SHEETS = {'СпрСобытий'}

needed_columns = [
    'День нед', 'Дата', 'Время',
    'Событие', 'Объект', 'Груз',
    'Кол-во', 'Примечание'
]

xl = pd.ExcelFile(excel_file)
sheets = [s for s in xl.sheet_names if s not in EXCLUDED_SHEETS]

html = '''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Журнал видеонаблюдения</title>
<style>
body { font-family: Arial, sans-serif; background:#f5f5f5; margin: 0; padding: 20px 10px; }
.header-sticky {
  position: sticky;
  top: 0;
  background: #f5f5f5;
  z-index: 100;
  padding: 10px 0;
  box-shadow: 0 2px 6px rgba(0,0,0,0.15);
}
h2 { margin: 0 0 10px 0; }
.tabs { margin-bottom: 10px; }
.tablink {
  padding: 8px 14px;
  border: none;
  background: #ddd;
  cursor: pointer;
  margin-right: 4px;
  border-radius: 4px 4px 0 0;
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
  box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
th, td {
  border: 1px solid #ccc;
  padding: 8px;
  text-align: left;
  vertical-align: top;
}
th { background: #eee; }
img.preview {
  max-width: 140px;
  max-height: 100px;
  object-fit: cover;
  cursor: pointer;
  border-radius: 4px;
}
</style>
</head>
<body>

<div class="header-sticky">
  <h2>Журнал видеонаблюдения</h2>
  <div class="tabs">
'''

# Вкладки
for sheet in sheets:
    html += f'<button class="tablink" onclick="openTab(event, \'{sheet}\')">{sheet}</button>'

html += '''
  </div>
</div>

<div class="week-controls">
  <button onclick="changeWeek(-1)">⏮ Предыдущая неделя</button>
  <strong id="weekLabel"></strong>
  <button onclick="changeWeek(1)">Следующая неделя ⏭</button>
</div>
'''

# Контент листов
for sheet in sheets:
    df = xl.parse(sheet).dropna(how='all').fillna('').reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    
    if 'Дата' not in df.columns:
        continue
    
    # Парсим даты
    dates = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
    
    # ISO для фильтра (только валидные)
    df['__date_iso'] = dates.dt.strftime('%Y-%m-%d')
    
    # Красивая дата для отображения (невалидные оставляем как есть)
    df['Дата'] = dates.dt.strftime('%d.%m.%Y').fillna(df['Дата'])
    
    # Фото
    def make_photo(row):
        link = row.get('Ссылка на фото', '')
        if not link:
            return '—'
        filename = Path(str(link)).name
        src = f'{images_dir}/{sheet}/{filename}'.replace('\\', '/')
        return f'<img class="preview" src="{src}" onclick="openImage(this.src)" alt="фото">'
    
    df['Фото'] = df.apply(make_photo, axis=1)
    
    final_cols = [c for c in needed_columns if c in df.columns]
    display_cols = final_cols + ['Фото']
    df = df[display_cols + ['__date_iso']]
    
    html += f'<div id="{sheet}" class="tabcontent" style="display:none;">'
    html += f'<h3>{sheet}</h3>'
    html += '<table><thead><tr>'
    for c in display_cols:
        html += f'<th>{c}</th>'
    html += '</tr></thead><tbody>'
    
    for _, row in df.iterrows():
        iso_date = row['__date_iso'] if pd.notna(row['__date_iso']) else ''
        html += f'<tr data-date="{iso_date}">'
        for c in display_cols:
            html += f'<td>{row[c]}</td>'
        html += '</tr>'
    
    html += '</tbody></table></div>'

# JavaScript
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
  return d.toLocaleDateString('ru-RU', {day: '2-digit', month: '2-digit', year: 'numeric'});
}

function renderWeek() {
  let monday = new Date(currentMonday);
  let sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  document.getElementById('weekLabel').innerText = 
    formatRU(monday) + ' — ' + formatRU(sunday);
  
  document.querySelectorAll('tr[data-date]').forEach(row => {
    let dateStr = row.dataset.date;
    if (!dateStr) {
      row.style.display = 'none';
      return;
    }
    let d = new Date(dateStr + 'T00:00:00');
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
  
  renderWeek();  // Важно: фильтр применяется после открытия вкладки
}

function openImage(src) {
  window.open(src, '_blank');
}

// Открываем первую вкладку и применяем фильтр
if (document.querySelector('.tablink')) {
  document.querySelector('.tablink').click();
}
</script>

</body>
</html>
'''

Path(output_html).write_text(html, encoding='utf-8')
print('ГОТОВО ✅ Журнал сформирован: index.html')