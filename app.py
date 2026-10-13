from pathlib import Path
import math, re, os
from datetime import datetime
from flask import Flask, request, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR=Path(__file__).resolve().parent
UPLOAD_DIR=BASE_DIR/'uploads'; EXPORT_DIR=BASE_DIR/'exports'; CONFIG_DIR=BASE_DIR/'config'
for folder in (UPLOAD_DIR,EXPORT_DIR,CONFIG_DIR): folder.mkdir(exist_ok=True)
app=Flask(__name__)
STATE={'employees':None,'days':None,'filename':None,'status':'idle','error':None}
NO_DEPT='Без подразделения'

def txt(v): return '' if v is None else str(v).strip()
def num(v):
    try: return float(txt(v).replace(' ','').replace(',','.'))
    except: return 0.0
def hm(v):
    m=re.search(r'(?<!\d)(\d{1,2})[:\-.](\d{2})(?!\d)',txt(v))
    return int(m.group(1))*60+int(m.group(2)) if m else None
def fm(v):
    if v is None: return '—'
    sign='-' if v<0 else ''; v=abs(round(v)); return f'{sign}{v//60:02d}:{v%60:02d}'
def fh(v): return fm(v*60)
def fsigned(v): return ('+' if round(v)>0 else '')+fm(v)
def esc(v): return txt(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace(chr(34),'&quot;')

def parse_interval(v):
    m=re.search(r'(\d{1,2}[:\-.]\d{2})\s*[-–—]\s*(\d{1,2}[:\-.]\d{2})',txt(v))
    if not m: return None,None
    start,end=hm(m.group(1)),hm(m.group(2))
    if start is not None and end is not None and end<start: end+=1440
    return start,end

def parse_schedule(v):
    source=txt(v); low=source.lower().replace('ё','е')
    if not source: return {'type':'missing','start':None,'hours':None}
    if 'свободн' in low: return {'type':'free','start':None,'hours':None}
    m=re.search(r'с\s*(\d{1,2})\s*[-:]\s*(\d{2})',low); start=int(m.group(1))*60+int(m.group(2)) if m else None
    m=re.search(r'\((\d+(?:[\.,]\d+)?)\s*час',low); hours=float(m.group(1).replace(',','.')) if m else None
    typ='five_two' if 'пятиднев' in low else ('shift' if 'сменн' in low else 'unknown')
    return {'type':typ,'start':start,'hours':hours}

def format_kind(v):
    low=txt(v).lower().replace('ё','е')
    if 'гибрид' in low: return 'hybrid'
    if 'удален' in low: return 'remote'
    return 'office'

def month_year_grid(grid):
    title=txt(grid[0][1]).lower()
    months={'январ':1,'феврал':2,'март':3,'апрел':4,'май':5,'июн':6,'июл':7,'август':8,'сентябр':9,'октябр':10,'ноябр':11,'декабр':12}
    month=next((v for k,v in months.items() if k in title),7); m=re.search(r'20\d{2}',title)
    return (int(m.group()) if m else datetime.now().year),month

def is_employee_grid(row):
    if row is None: return False
    number=txt(row[1]) if len(row)>1 else ''
    name=txt(row[3]) if len(row)>3 else ''
    personnel=txt(row[4]) if len(row)>4 else ''
    return bool(re.fullmatch(r'\d+',number) and name and name.lower()!='сотрудник' and re.fullmatch(r'\d{4}-\d{5}',personnel))

def strict_day(rule,missing,start,end,rv,skud):
    """Строгий расчёт для дашборда: без допусков, всё в минутах."""
    z={'plan':0,'fact':0,'late':0,'early':0,'late_leave':0,'off':0,'early_arr':0,'remote_rv':0,'free':0}
    has_int = start is not None and end is not None
    presence = max(end-start,0) if has_int else round(skud*60)
    if rule['type']=='free':
        if has_int or skud>0: z['free']=presence
        return z
    if rv<=0:
        if has_int or skud>0: z['off']=presence
        return z
    planm=round(rv*60)
    if not has_int and skud<=0:
        z['remote_rv']=planm
        return z
    if missing:
        z['plan']=planm; z['fact']=round(skud*60) if skud>0 else presence
        return z
    z['plan']=planm
    if rule['start'] is None:
        z['fact']=presence
        return z
    ps=rule['start']; pe=ps+planm
    if has_int:
        worked=max(0,min(end,pe)-max(start,ps))
        extra=max(0,end-max(start,pe))
        z['late']=min(max(0,start-ps),planm); z['early']=min(max(0,pe-end),planm)
        z['early_arr']=max(0,min(end,ps)-start); z['late_leave']=extra
        z['fact']=worked+extra
    else:
        z['fact']=round(skud*60)
    return z

def analyze(path):
    wb=load_workbook(path,data_only=True,read_only=True)
    ws=wb.active
    grid=[row for row in ws.iter_rows(values_only=True)]
    wb.close()
    year,month=month_year_grid(grid)
    max_row=len(grid); employees=[]; days=[]
    for i in range(max_row):
        row=grid[i]
        if not is_employee_grid(row) or i+2>=max_row: continue
        department=txt(row[2])
        work_format=txt(row[8]) or 'Не указан'; fmt=format_kind(work_format); raw_schedule=txt(row[9]); rule=parse_schedule(raw_schedule)
        is_missing_schedule = rule['type']=='missing'
        e={'number':txt(row[1]),'name':txt(row[3]),'department':department,'personnel':txt(row[4]),'work_format':work_format,'schedule':raw_schedule or 'График не указан','rule':rule,'summary':{'work':0,'remote':0,'late':0,'early':0,'anomaly':0,'late_leave':0,'above':0.0,'below':0.0,'off':0,'problems':0,'excluded':False,'termination':is_missing_schedule}}
        data_row=grid[i+2] if i+2<max_row else ()
        for day in range(1,32):
            c=49+(day-1)*2
            if c+1>=len(row) or c+1>=len(data_row): break
            try: date=datetime(year,month,day).strftime('%d.%m.%Y')
            except: continue
            raw=row[c] if c<len(row) else None
            rv=num(data_row[c]) if c<len(data_row) else 0.0
            skud=num(data_row[c+1]) if c+1<len(data_row) else 0.0
            start,end=parse_interval(raw); statuses=[]; kind=''; plan='—'
            m=strict_day(rule,is_missing_schedule,start,end,rv,skud)
            if is_missing_schedule:
                if rv>0:
                    e['summary']['work']+=1
                    if skud>0:
                        d=(skud-rv)*60
                        if d>5: statuses.append('Отклонение длительности СКУД от РВ: +'+fm(d));e['summary']['above']+=skud-rv
                        elif d<-5: statuses.append('Отклонение длительности СКУД от РВ: '+fm(d));e['summary']['below']+=rv-skud
                        else: statuses.append('СКУД соответствует РВ')
                    else:
                        statuses=['Удалённая работа (потенциально)']; kind='remote'; e['summary']['remote']+=1
                else:
                    statuses=[]; kind=''
            elif rule['type']=='free':
                if raw or skud>0: statuses=['Свободное посещение/Подработка']; kind='free'
            elif rv<=0:
                if raw or skud>0: statuses=['Посещение вне графика / переработка'];kind='off_schedule';e['summary']['off']+=1
            elif fmt=='remote' and not raw and skud<=0:
                statuses=['Удалённая работа'];kind='remote';e['summary']['remote']+=1;e['summary']['work']+=1
            elif fmt=='hybrid' and not raw and skud<=0:
                statuses=['Гибрид/удалённая работа'];kind='remote';e['summary']['remote']+=1;e['summary']['work']+=1
            else:
                e['summary']['work']+=1
                if rule['start'] is None:
                    statuses=['Нераспознанный график'];kind='unknown';e['summary']['problems']+=1
                else:
                    planned_end=rule['start']+rv*60; plan=f"{fm(rule['start'])}–{fm(planned_end%1440)}"+(' (+1 день)' if planned_end>=1440 else '')
                    if start is None or end is None:
                        statuses=['Проблема: нет полного интервала'];kind='problem';e['summary']['problems']+=1
                    else:
                        if start<rule['start']-10: statuses.append('Аномально ранний приход '+fm(rule['start']-start));e['summary']['anomaly']+=1
                        elif start>rule['start']: statuses.append('Опоздание '+fm(start-rule['start']));e['summary']['late']+=1
                        elif start<rule['start']: statuses.append('Ранний приход '+fm(rule['start']-start))
                        delta=end-planned_end
                        if delta<0: statuses.append('Ранний уход '+fm(-delta));e['summary']['early']+=1
                        elif delta>5: statuses.append('Поздний уход '+fm(delta));e['summary']['late_leave']+=1
                        d=(skud-rv)*60
                        if d>5: statuses.append('Отклонение длительности СКУД от РВ: +'+fm(d));e['summary']['above']+=skud-rv
                        elif d<-5: statuses.append('Отклонение длительности СКУД от РВ: '+fm(d));e['summary']['below']+=rv-skud
                        else: statuses.append('СКУД соответствует РВ')
            days.append({'number':e['number'],'name':e['name'],'department':department,'work_format':work_format,'date':date,'rv':rv,'skud':skud,'interval':txt(raw) or '—','plan':plan,'statuses':statuses,'kind':kind,'m':m})
        employees.append(e)
    return employees,days

COLUMN_DESCRIPTIONS=[
    ('Рабочие дни', 'Количество дней в месяце, когда по графику было запланировано присутствие (норма РВ больше нуля). Включает дни удалённой и гибридной работы без визита в офис, а также рабочие дни у сотрудников со статусом "Увольнение сотрудника", если в этот день указано РВ больше нуля.'),
    ('Удалённые дни', 'Из общего числа рабочих дней — сколько дней прошло без визита в офис: нет интервала прихода-ухода, и СКУД равен нулю. Не считается проблемой и не влияет на показатели опозданий и уходов.'),
    ('Опоздания', 'Количество дней, когда сотрудник пришёл позже времени начала смены по графику. Считается с первой минуты просрочки, без допуска.'),
    ('Ранние уходы', 'Количество дней, когда сотрудник ушёл раньше планового окончания смены. Считается с первой минуты, без допуска. Не компенсируется тем, что человек пришёл раньше положенного.'),
    ('Аномально ранние приходы', 'Количество дней, когда сотрудник пришёл раньше плана более чем на 10 минут. Приход раньше на 10 минут и меньше считается нормальным и не попадает в этот показатель.'),
    ('Поздние уходы', 'Количество дней, когда сотрудник задержался после планового окончания смены более чем на 5 минут. Задержка до 5 минут включительно считается нейтральной.'),
    ('СКУД выше РВ', 'Суммарное количество часов за месяц, на которое фактическое присутствие по данным пропускной системы (СКУД) превысило норму по графику (РВ). Показывает переработку по продолжительности присутствия.'),
    ('СКУД ниже РВ', 'Суммарное количество часов за месяц, на которое фактическое присутствие оказалось меньше нормы. Показывает недостаточное присутствие на работе.'),
    ('Статус формирования отчёта', '"Сформирован" — у сотрудника есть распознанный график, показатели рассчитаны полностью и сравниваются с планом. "Увольнение сотрудника" — график не указан в исходном файле (обычно люди, принятые или уволенные в середине месяца); дни оцениваются по фактическому присутствию СКУД, но без проверки опозданий и уходов, так как нет времени начала смены для сравнения.'),
]

DASH_DESCRIPTIONS=[
    ('Подразделение', 'Подразделение из исходного файла. Сотрудники без указанного подразделения собраны в строку "Без подразделения".'),
    ('Сотрудники', 'Количество сотрудников подразделения в загруженном файле. Показано только для масштаба.'),
    ('План (РВ), ч', 'Сколько времени нужно было отработать: сумма нормы РВ по рабочим дням сотрудников подразделения. Дни, в которые РВ больше нуля, но нет ни интервала, ни СКУД, считаются потенциально удалёнными и в план не входят (см. серую колонку).'),
    ('Факт, ч', 'Сколько фактически отработано в те же дни: от прихода до ухода. Если сотрудник пришёл раньше планового начала смены, время до начала смены не засчитывается. Время после планового окончания засчитывается.'),
    ('Разница план/факт, ч', 'Факт минус план. Складывается из: поздние уходы (плюс), опоздания и ранние уходы (минус). Раннее появление в офисе на разницу не влияет.'),
    ('Опоздания, ч', 'Суммарное время опозданий: на сколько приход был позже планового начала смены. Считается с первой минуты, без допусков.'),
    ('Ранние уходы, ч', 'Суммарное время ранних уходов: на сколько уход был раньше планового окончания смены. Считается с первой минуты, без допусков.'),
    ('Поздние уходы, ч', 'Суммарное время после планового окончания смены. Считается с первой минуты, без допусков.'),
    ('Вне графика, ч', 'Сколько часов отработано в дни, когда смена по графику не была запланирована (РВ равен нулю). Берётся время от прихода до ухода. В план и разницу не входит.'),
    ('Ранний приход, ч (справочно)', 'Суммарное время прихода раньше планового начала смены. Только для информации: в факт и в разницу план/факт не входит.'),
    ('Потенциально удалённые, ч (не в сравнении)', 'Норма РВ тех дней, когда РВ больше нуля, но нет ни прихода и ухода, ни данных СКУД. Такие дни считаются потенциально удалённой работой (для удалённых и гибридных сотрудников — удалённой). В план и факт эти часы не включены.'),
    ('Свободный график, ч (без нормы)', 'Время присутствия сотрудников с графиком "Свободный график": нормы у такого графика нет, поэтому часы не входят ни в план, ни в факт, ни в "Вне графика".'),
]

def export_xlsx(emps,days,path):
    wb=Workbook();ws=wb.active;ws.title='Анализ данных'
    headers=['№','Сотрудник','Подразделение','Табельный номер','Формат работы','График','Рабочие дни','Удалённые дни','Опоздания','Ранние уходы','Аномально ранние приходы','Поздние уходы >5 мин','СКУД выше РВ, ч','СКУД ниже РВ, ч','Вне графика, дней','Проблемные дни','Статус формирования отчёта']
    ws.append(headers)
    header_fill=PatternFill(start_color='2F6FED',end_color='2F6FED',fill_type='solid')
    header_font=Font(color='FFFFFF',bold=True,size=11)
    thin=Side(style='thin',color='D9DEE3')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    center=Alignment(horizontal='center',vertical='center')
    for cell in ws[1]:
        cell.fill=header_fill; cell.font=header_font; cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=border
    ws.row_dimensions[1].height=32
    green_fill=PatternFill(start_color='E3F6E9',end_color='E3F6E9',fill_type='solid')
    gray_fill=PatternFill(start_color='FDECEC',end_color='FDECEC',fill_type='solid')
    for e in emps:
        s=e['summary']
        status='Увольнение сотрудника' if s['termination'] else 'Сформирован'
        ws.append([e['number'],e['name'],e.get('department',''),e['personnel'],e['work_format'],e['schedule'],s['work'],s['remote'],s['late'],s['early'],s['anomaly'],s['late_leave'],round(s['above'],2),round(s['below'],2),s['off'],s['problems'],status])
        r=ws.max_row
        fill=gray_fill if s['termination'] else green_fill
        ws.cell(r,17).fill=fill
        for c in range(1,18):
            ws.cell(r,c).border=border
            if c not in (2,3,5,6): ws.cell(r,c).alignment=center
    ws.freeze_panes='A2'
    ws.auto_filter.ref=ws.dimensions
    widths={1:8,2:32,3:26,4:16,5:16,6:40,7:11,8:12,9:11,10:11,11:14,12:14,13:12,14:12,15:12,16:12,17:22}
    for c,w in widths.items(): ws.column_dimensions[get_column_letter(c)].width=w
    wd=wb.create_sheet('Дни')
    dheaders=['№','Сотрудник','Подразделение','Формат работы','Дата','РВ','СКУД','Интервал','План','Статусы']
    wd.append(dheaders)
    for cell in wd[1]:
        cell.fill=header_fill; cell.font=header_font; cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=border
    wd.row_dimensions[1].height=28
    bad_fill=PatternFill(start_color='FDECEC',end_color='FDECEC',fill_type='solid')
    warn_fill=PatternFill(start_color='FFF6DF',end_color='FFF6DF',fill_type='solid')
    blue_fill=PatternFill(start_color='EAF1FF',end_color='EAF1FF',fill_type='solid')
    for d in days:
        wd.append([d['number'],d['name'],d.get('department',''),d['work_format'],d['date'],d['rv'],d['skud'],d['interval'],d['plan'],'; '.join(d['statuses'])])
        statuses=d['statuses']
        bad=any(x.startswith(('Опоздание','Ранний уход','Аномально','Проблема','Нераспознанный')) for x in statuses)
        warn=any(x.startswith(('Поздний уход','Отклонение длительности')) for x in statuses)
        fill=bad_fill if bad else (warn_fill if warn else (blue_fill if d['kind'] in ('remote','free','off_schedule') else None))
        if fill:
            r=wd.max_row
            for c in range(1,11): wd.cell(r,c).fill=fill
    wd.freeze_panes='A2'
    wd.auto_filter.ref=wd.dimensions
    dwidths={1:8,2:32,3:26,4:16,5:12,6:9,7:9,8:18,9:20,10:60}
    for c,w in dwidths.items(): wd.column_dimensions[get_column_letter(c)].width=w
    de=wb.create_sheet('Описание значений')
    de.append(['Колонка','Описание'])
    for cell in de[1]:
        cell.fill=header_fill; cell.font=header_font; cell.alignment=Alignment(horizontal='center',vertical='center'); cell.border=border
    de.row_dimensions[1].height=24
    for name,desc in COLUMN_DESCRIPTIONS:
        de.append([name,desc])
        r=de.max_row
        de.cell(r,1).font=Font(bold=True)
        de.cell(r,1).alignment=Alignment(vertical='top',wrap_text=True)
        de.cell(r,2).alignment=Alignment(vertical='top',wrap_text=True)
        de.cell(r,1).border=border; de.cell(r,2).border=border
        de.row_dimensions[r].height=48
    de.column_dimensions['A'].width=28
    de.column_dimensions['B'].width=100
    wb.save(path)

def safe_filename(name):
    cleaned=re.sub(r'[\\/:*?"<>|]','_',name).strip()
    return cleaned or 'export'

def filtered_export_name(selected,dept_selected):
    if dept_selected and not selected:
        return safe_filename(dept_selected)
    return 'Моя выгрузка'

# ---------- Дашборд по подразделениям ----------
DKEYS=['plan','fact','late','early','late_leave','off','early_arr','remote_rv','free']

def dept_name(v): return txt(v) or NO_DEPT

def dashboard_data(emps,days,dept_filter=''):
    agg={}
    for e in emps:
        a=agg.setdefault(dept_name(e.get('department')),{'emp':0,**{k:0 for k in DKEYS}})
        a['emp']+=1
    for x in days:
        a=agg.get(dept_name(x.get('department')))
        m=x.get('m')
        if a is None or not m: continue
        for k in DKEYS: a[k]+=m[k]
    rows=[]
    for name,a in agg.items():
        if dept_filter and name!=dept_filter: continue
        r=dict(a); r['dept']=name; r['diff']=r['fact']-r['plan']; rows.append(r)
    rows.sort(key=lambda r:(r['diff'],r['dept']))
    total={'emp':sum(r['emp'] for r in rows),**{k:sum(r[k] for r in rows) for k in DKEYS}}
    total['diff']=total['fact']-total['plan']
    return rows,total

def export_dashboard_xlsx(emps,days,dept_filter,path):
    rows,total=dashboard_data(emps,days,dept_filter)
    wb=Workbook(); ws=wb.active; ws.title='Дашборд по подразделениям'
    header_fill=PatternFill(start_color='2F6FED',end_color='2F6FED',fill_type='solid')
    header_font=Font(color='FFFFFF',bold=True,size=11)
    thin=Side(style='thin',color='D9DEE3'); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.append([n for n,_ in DASH_DESCRIPTIONS])
    for cell in ws[1]:
        cell.fill=header_fill; cell.font=header_font; cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=border
    ws.row_dimensions[1].height=48
    def hrs(v): return round(v/60,2)
    def line(label,r):
        return [label,r['emp'],hrs(r['plan']),hrs(r['fact']),hrs(r['diff']),hrs(r['late']),hrs(r['early']),hrs(r['late_leave']),hrs(r['off']),hrs(r['early_arr']),hrs(r['remote_rv']),hrs(r['free'])]
    red=PatternFill(start_color='FDECEC',end_color='FDECEC',fill_type='solid')
    grn=PatternFill(start_color='E3F6E9',end_color='E3F6E9',fill_type='solid')
    for r in rows:
        ws.append(line(r['dept'],r)); n=ws.max_row
        if r['diff']<0: ws.cell(n,5).fill=red
        elif r['diff']>0: ws.cell(n,5).fill=grn
    ws.append(line('ИТОГО',total)); n=ws.max_row
    for c in range(1,13): ws.cell(n,c).font=Font(bold=True)
    for r in ws.iter_rows(min_row=2,max_row=ws.max_row):
        for cell in r:
            cell.border=border
            if cell.column>1: cell.alignment=Alignment(horizontal='center')
    ws.freeze_panes='B2'
    ws.column_dimensions['A'].width=34
    for c in range(2,13): ws.column_dimensions[get_column_letter(c)].width=16
    de=wb.create_sheet('Описание значений')
    de.append(['Колонка','Описание'])
    for cell in de[1]:
        cell.fill=header_fill; cell.font=header_font; cell.alignment=Alignment(horizontal='center',vertical='center'); cell.border=border
    for name,desc in DASH_DESCRIPTIONS:
        de.append([name,desc]); n=de.max_row
        de.cell(n,1).font=Font(bold=True)
        de.cell(n,1).alignment=Alignment(vertical='top',wrap_text=True); de.cell(n,2).alignment=Alignment(vertical='top',wrap_text=True)
        de.cell(n,1).border=border; de.cell(n,2).border=border; de.row_dimensions[n].height=48
    de.append([]); de.append(['Примечание','Значения в часах (десятичные доли часа). Допуски не применяются: любое отклонение считается с первой минуты.'])
    de.column_dimensions['A'].width=38; de.column_dimensions['B'].width=100
    wb.save(path)

COLUMN_INFO_MAP = dict(COLUMN_DESCRIPTIONS)
DASH_INFO_MAP = dict(DASH_DESCRIPTIONS)

def th_with_info(label,key):
    tip = COLUMN_INFO_MAP.get(key,'')
    return '<th class="info-th" title="'+esc(tip)+'">'+label+' <span class="info-icon">ⓘ</span></th>'

def th_dash(key):
    return '<th class="info-th" title="'+esc(DASH_INFO_MAP.get(key,''))+'">'+esc(key)+' <span class="info-icon">ⓘ</span></th>'

def render(emps,days,selected,dept_selected):
    selected=str(selected or ''); dept_selected=str(dept_selected or '')
    es=[e for e in emps if not selected or e['name']==selected]
    if dept_selected: es=[e for e in es if e.get('department','')==dept_selected]
    ids={e['number'] for e in es}; ds=[d for d in days if d['number'] in ids]
    all_names=sorted({e['name'] for e in emps})
    name_options=''.join('<option value="'+esc(n)+'">' for n in all_names)
    all_depts=sorted({e.get('department','') for e in emps if e.get('department')})
    dept_options=''.join('<option value="'+esc(dp)+'"'+(' selected' if dp==dept_selected else '')+'>'+esc(dp)+'</option>' for dp in all_depts)
    title=selected if selected else 'Все сотрудники'
    if dept_selected: title+=' — '+dept_selected
    has_filter = bool(selected or dept_selected)
    filter_query = ('employee='+esc(selected)+'&' if selected else '')+('department='+esc(dept_selected) if dept_selected else '')
    sr=[]
    for e in es:
        s=e['summary']
        status='<span class="tag tag-gray">Увольнение сотрудника</span>' if s['termination'] else '<span class="tag tag-green">Сформирован</span>'
        sr.append('<tr><td>'+esc(e['number'])+'</td><td class="name">'+esc(e['name'])+'</td><td>'+esc(e.get('department',''))+'</td><td>'+esc(e['work_format'])+'</td><td class="sched">'+esc(e['schedule'])+'</td><td>'+str(s['work'])+'</td><td>'+str(s['remote'])+'</td><td>'+str(s['late'])+'</td><td>'+str(s['early'])+'</td><td>'+str(s['anomaly'])+'</td><td>'+str(s['late_leave'])+'</td><td>'+fh(s['above'])+'</td><td>'+fh(s['below'])+'</td><td>'+status+'</td></tr>')
    dr=[]
    for d in ds:
        st='; '.join(d['statuses']) or '—'; bad=any(x.startswith(('Опоздание','Ранний уход','Аномально','Проблема','Нераспознанный')) for x in d['statuses']); warn=any(x.startswith(('Поздний уход','Отклонение длительности')) for x in d['statuses'])
        cl='bad' if bad else ('warn' if warn else ('blue' if d['kind'] in ('remote','free','off_schedule') else ''))
        dr.append('<tr class="'+cl+'"><td>'+esc(d['number'])+'</td><td class="name">'+esc(d['name'])+'</td><td>'+esc(d.get('department',''))+'</td><td>'+esc(d['work_format'])+'</td><td>'+d['date']+'</td><td>'+fh(d['rv'])+'</td><td>'+fh(d['skud'])+'</td><td>'+esc(d['interval'])+'</td><td>'+esc(d['plan'])+'</td><td>'+esc(st)+'</td></tr>')
    filters='<div class="card"><form method="get" action="/" class="select-row"><input type="hidden" name="tab" value="employees"><label>Сотрудник</label><input list="names" name="employee" value="'+esc(selected)+'" placeholder="Начните вводить ФИО" autocomplete="off"><datalist id="names">'+name_options+'</datalist><label>Подразделение</label><select name="department"><option value="">Все подразделения</option>'+dept_options+'</select><button class="btn">Показать</button><a class="btn btn-ghost" href="/?tab=employees">Сбросить</a></form></div>'
    export_buttons='<div class="stat"><a class="btn btn-green" href="/export/full">Export Excel (Полная выгрузка)</a></div>'
    if has_filter:
        export_buttons+='<div class="stat"><a class="btn btn-accent" href="/export/filtered?'+filter_query+'">Export Excel по фильтру</a></div>'
    thead='<tr><th>№</th><th>Сотрудник</th><th>Подразделение</th><th>Формат работы</th><th>График</th>'
    thead+=th_with_info('Рабочие дни','Рабочие дни')+th_with_info('Удалённые дни','Удалённые дни')+th_with_info('Опоздания','Опоздания')+th_with_info('Ранние уходы','Ранние уходы')+th_with_info('Аномально ранние приходы','Аномально ранние приходы')+th_with_info('Поздние уходы','Поздние уходы')+th_with_info('СКУД выше РВ','СКУД выше РВ')+th_with_info('СКУД ниже РВ','СКУД ниже РВ')+th_with_info('Статус формирования отчёта','Статус формирования отчёта')
    thead+='</tr>'
    return filters+'<div class="card"><div class="grid"><div class="stat"><div class="metric">'+str(len(es))+'</div><div class="stat-label">сотрудников</div></div><div class="stat"><div class="metric">'+str(len(ds))+'</div><div class="stat-label">дневных записей</div></div>'+export_buttons+'</div></div><div class="card"><h2>Анализ данных: '+esc(title)+'</h2><div class="wrap"><table>'+thead+''.join(sr)+'</table></div></div><div class="card"><h2>Детализация по дням</h2><div class="wrap"><table><tr><th>№</th><th>Сотрудник</th><th>Подразделение</th><th>Формат работы</th><th>Дата</th><th>РВ</th><th>СКУД</th><th>Интервал</th><th>План</th><th>Статусы</th></tr>'+''.join(dr)+'</table></div></div>'

def donut_svg(parts):
    total=sum(v for _,v,_ in parts)
    if total<=0:
        return '<p class="muted">Отклонений нет — строить график не из чего.</p>'
    r=70; C=2*math.pi*r; off=0.0; circles=''; legend=''
    for label,v,color in parts:
        share=v/total*100
        legend+='<tr><td><span class="dot" style="background:'+color+'"></span>'+esc(label)+'</td><td class="num">'+fm(v)+' ч</td><td class="num">'+f'{share:.1f}'+'%</td></tr>'
        if v<=0: continue
        ln=C*v/total
        circles+='<circle cx="100" cy="100" r="70" fill="none" stroke="'+color+'" stroke-width="34" stroke-dasharray="'+f'{ln:.3f} {C-ln:.3f}'+'" stroke-dashoffset="'+f'{-off:.3f}'+'" transform="rotate(-90 100 100)"><title>'+esc(label)+': '+fm(v)+' ч</title></circle>'
        off+=ln
    svg='<svg viewBox="0 0 200 200" width="220" height="220" role="img"><circle cx="100" cy="100" r="70" fill="none" stroke="#eef1f5" stroke-width="34"></circle>'+circles+'<text x="100" y="96" text-anchor="middle" font-size="12" fill="#6b7684">Всего отклонений</text><text x="100" y="116" text-anchor="middle" font-size="16" font-weight="700" fill="#1b2430">'+fm(total)+' ч</text></svg>'
    return '<div class="donut-wrap">'+svg+'<table class="legend">'+legend+'</table></div>'

def render_dashboard(emps,days,dept_selected):
    dept_selected=str(dept_selected or '')
    names=sorted({dept_name(e.get('department')) for e in emps})
    if dept_selected not in names: dept_selected=''
    rows,t=dashboard_data(emps,days,dept_selected)
    opts=''.join('<option value="'+esc(n)+'"'+(' selected' if n==dept_selected else '')+'>'+esc(n)+'</option>' for n in names)
    filters='<div class="card"><form method="get" action="/" class="select-row"><input type="hidden" name="tab" value="dashboard"><label>Подразделение</label><select name="department"><option value="">Все подразделения</option>'+opts+'</select><button class="btn">Показать</button><a class="btn btn-ghost" href="/?tab=dashboard">Сбросить</a></form></div>'
    def sgn(v): return 'neg' if round(v)<0 else ('pos' if round(v)>0 else '')
    cards=('<div class="card"><div class="grid">'
        '<div class="stat"><div class="metric">'+str(len(rows))+'</div><div class="stat-label">подразделений</div></div>'
        '<div class="stat"><div class="metric">'+fm(t['plan'])+'</div><div class="stat-label">план (РВ), ч</div></div>'
        '<div class="stat"><div class="metric">'+fm(t['fact'])+'</div><div class="stat-label">факт, ч</div></div>'
        '<div class="stat"><div class="metric '+sgn(t['diff'])+'">'+fsigned(t['diff'])+'</div><div class="stat-label">разница план/факт, ч</div></div>'
        '<div class="stat"><a class="btn btn-green" href="/export/dashboard?department='+esc(dept_selected)+'">Export Excel (Дашборд)</a></div>'
        '</div></div>')
    parts=[('Опоздания',t['late'],'#e5484d'),('Ранние уходы',t['early'],'#f5a524'),('Поздние уходы',t['late_leave'],'#2f6fed'),('Вне графика',t['off'],'#8b5cf6')]
    chart='<div class="card"><h2>Структура отклонений</h2>'+donut_svg(parts)+'</div>'
    head='<tr>'+''.join(th_dash(n) for n,_ in DASH_DESCRIPTIONS)+'</tr>'
    def line(label,r,cls=''):
        return ('<tr class="'+cls+'"><td class="name">'+esc(label)+'</td><td>'+str(r['emp'])+'</td><td>'+fm(r['plan'])+'</td><td>'+fm(r['fact'])+'</td>'
            '<td class="'+sgn(r['diff'])+' strong">'+fsigned(r['diff'])+'</td><td>'+fm(r['late'])+'</td><td>'+fm(r['early'])+'</td><td>'+fm(r['late_leave'])+'</td><td>'+fm(r['off'])+'</td><td class="grey">'+fm(r['early_arr'])+'</td><td class="grey">'+fm(r['remote_rv'])+'</td><td class="grey">'+fm(r['free'])+'</td></tr>')
    body=''.join(line(r['dept'],r) for r in rows)+line('ИТОГО',t,'total')
    note=('<p class="note">Расчёт без допусков: любое отклонение считается с первой минуты. Время раннего прихода в факт не входит. '
          'Дни, в которые РВ больше нуля, но нет ни интервала, ни СКУД, считаются потенциально удалёнными и в план и факт не входят; присутствие по свободному графику тоже вынесено в серые колонки. '
          'У сотрудников без графика нет времени начала смены, поэтому их отклонение входит только в разницу план/факт. Часы показаны в формате ЧЧ:ММ.</p>')
    table='<div class="card"><h2>Дашборд по подразделениям</h2><div class="wrap"><table>'+head+body+'</table></div>'+note+'</div>'
    return filters+cards+chart+table

CSS=(':root{--bg:#f2f4f7;--card:#ffffff;--ink:#1b2430;--muted:#6b7684;--accent:#2f6fed;--accent-dark:#1f56c4;--green:#1f9d55;--radius:14px}'
     '*{box-sizing:border-box}body{font-family:"Segoe UI",Arial,sans-serif;margin:0;padding:28px 32px;background:linear-gradient(180deg,#eef2f8,#f6f8fb);color:var(--ink)}'
     'h1{font-size:26px;margin:0 0 4px}h2{font-size:17px;margin:0 0 14px;color:var(--ink)}'
     '.subtitle{color:var(--muted);margin:0 0 20px;font-size:14px}'
     '.card{background:var(--card);border-radius:var(--radius);padding:20px 22px;margin:16px 0;box-shadow:0 4px 16px rgba(20,30,60,0.06);border:1px solid #e7ebf1}'
     '.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px}.stat{text-align:center;padding:6px 0}'
     '.metric{font-size:30px;font-weight:700;color:var(--accent-dark)}.stat-label{color:var(--muted);font-size:13px;margin-top:2px}'
     'table{border-collapse:collapse;width:100%;background:var(--card);font-size:13px}'
     'th,td{padding:10px 12px;border-bottom:1px solid #eef1f5;text-align:left;vertical-align:top}'
     'th{background:#f7f9fc;color:#48505c;font-weight:600;position:sticky;top:0;border-bottom:2px solid #e3e8ef}'
     'tr:hover td{background:#f8fafc}.name{font-weight:600}.sched{color:var(--muted);font-size:12px}'
     '.wrap{overflow:auto;max-height:600px;border-radius:10px;border:1px solid #eef1f5}'
     '.bad td{background:#fdecec}.warn td{background:#fff6df}.blue td{background:#eaf1ff}'
     '.btn{background:var(--accent);color:#fff;padding:10px 18px;border:0;border-radius:9px;text-decoration:none;cursor:pointer;font-weight:600;font-size:14px;display:inline-block;transition:background .15s}'
     '.btn:hover{background:var(--accent-dark)}.btn-green{background:var(--green)}.btn-green:hover{background:#177a41}'
     '.btn-ghost{background:#eef1f6;color:var(--muted)}.btn-ghost:hover{background:#e3e7ee}'
     '.btn-accent{background:#8b5cf6}.btn-accent:hover{background:#7440e0}'
     'input[type=file]{padding:8px;border:1px dashed #c7cede;border-radius:8px;background:#fbfcfe}'
     '.select-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}.select-row label{font-weight:600;color:var(--muted);font-size:13px}'
     'select,input[list]{padding:9px 12px;min-width:260px;border-radius:9px;border:1px solid #d7dce4;background:#fff;font-size:13px}'
     '.muted{color:var(--muted);margin:10px 0 0}.spin{color:#b8511e;font-weight:600;margin:10px 0 0}'
     '.tag{display:inline-block;padding:3px 9px;border-radius:999px;font-size:12px;font-weight:600}'
     '.tag-green{background:#e3f6e9;color:#177a41}.tag-gray{background:#eceff3;color:#5b6472}'
     'form[action="/upload"]{display:flex;gap:12px;align-items:center;flex-wrap:wrap}'
     '.info-th{cursor:help;position:relative}.info-icon{font-size:11px;color:#9aa4b2}.info-th:hover{color:var(--accent-dark)}'
     '.tabs{display:flex;gap:8px;margin:18px 0 4px;border-bottom:2px solid #e3e8ef}'
     '.tab{padding:11px 22px;border-radius:10px 10px 0 0;text-decoration:none;font-weight:600;font-size:15px;color:var(--muted);background:#eef1f6;border:1px solid #e3e8ef;border-bottom:none;margin-bottom:-2px}'
     '.tab:hover{color:var(--accent-dark)}.tab.active{background:#fff;color:var(--accent-dark);border-bottom:2px solid #fff}'
     '.pos{color:#177a41}.neg{color:#b42318}.strong{font-weight:700}.grey{color:#8a93a0}'
     'tr.total td{background:#f1f4f9;font-weight:700;border-top:2px solid #d7dce4}'
     '.donut-wrap{display:flex;gap:32px;align-items:center;flex-wrap:wrap}'
     '.legend{width:auto;min-width:320px}.legend td{border:0;padding:6px 14px 6px 0}.legend .num{text-align:right;font-variant-numeric:tabular-nums}'
     '.dot{display:inline-block;width:12px;height:12px;border-radius:50%;margin-right:8px;vertical-align:middle}'
     '.note{color:var(--muted);font-size:12px;margin:12px 0 0;line-height:1.5}')
TOP='<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Анализ присутствия сотрудников</title><style>'+CSS+'</style></head><body>'
BOTTOM='</body></html>'

def form():
    if STATE['status']=='processing':
        note='<p class="spin">⏳ Файл обрабатывается, подождите несколько секунд.</p>'
    elif STATE['error']:
        note='<p class="spin">Ошибка обработки: '+esc(STATE['error'])+'</p>'
    elif STATE['filename']:
        note='<p class="muted">Загружен файл: <b>'+esc(STATE['filename'])+'</b>. Загрузите новый файл, чтобы заменить текущие данные.</p>'
    else:
        note='<p class="muted">Загрузите Excel-файл отчёта о присутствии, чтобы увидеть анализ данных.</p>'
    return '<h1>Анализ присутствия сотрудников</h1><p class="subtitle">План/факт рабочего времени, опоздания, переработки и удалённая работа</p><div class="card"><form method="post" enctype="multipart/form-data" action="/upload"><input type="file" name="file" accept=".xlsx,.xls" required> <button class="btn">Обработать</button></form>'+note+'</div>'

def tabs_bar(tab):
    a='tab active' if tab=='employees' else 'tab'
    b='tab active' if tab=='dashboard' else 'tab'
    return '<div class="tabs"><a class="'+a+'" href="/?tab=employees">Анализ по сотрудникам</a><a class="'+b+'" href="/?tab=dashboard">Дашборд по подразделениям</a></div>'

def build_page(tab,args):
    if STATE['employees'] is None:
        return TOP+form()+BOTTOM
    tab='dashboard' if tab=='dashboard' else 'employees'
    if tab=='dashboard':
        body=render_dashboard(STATE['employees'],STATE['days'],args.get('department',''))
    else:
        body=render(STATE['employees'],STATE['days'],args.get('employee',''),args.get('department',''))
    return TOP+form()+tabs_bar(tab)+body+BOTTOM

@app.route('/',methods=['GET'])
def index():
    return build_page(request.args.get('tab','employees'),request.args)

@app.route('/upload',methods=['GET','POST'])
def upload():
    if request.method=='GET':
        return build_page('employees',{})
    f=request.files.get('file')
    if not f or not f.filename: return build_page('employees',{})
    if Path(f.filename).suffix.lower() not in ('.xlsx','.xls'): return TOP+form()+'<div class="card">Файл должен иметь расширение .xlsx или .xls.</div>'+BOTTOM
    path=UPLOAD_DIR/Path(f.filename).name; f.save(path)
    STATE['status']='processing'
    try:
        emps,days=analyze(path)
        STATE['employees']=emps;STATE['days']=days;STATE['filename']=f.filename;STATE['error']=None
        for old in EXPORT_DIR.glob('*.xlsx'): old.unlink()
    except Exception as ex:
        STATE['error']=str(ex)
    finally:
        STATE['status']='idle'
    return build_page('employees',{})

@app.route('/export/full')
def export_full():
    if STATE['employees'] is None: return TOP+form()+BOTTOM
    target=EXPORT_DIR/'result_full.xlsx'
    if not target.exists(): export_xlsx(STATE['employees'],STATE['days'],target)
    return send_file(target,as_attachment=True,download_name='Полная выгрузка.xlsx')

@app.route('/export/filtered')
def export_filtered():
    if STATE['employees'] is None: return TOP+form()+BOTTOM
    selected=request.args.get('employee',''); dept_selected=request.args.get('department','')
    es=[e for e in STATE['employees'] if not selected or e['name']==selected]
    if dept_selected: es=[e for e in es if e.get('department','')==dept_selected]
    ids={e['number'] for e in es}; ds=[d for d in STATE['days'] if d['number'] in ids]
    fname=filtered_export_name(selected,dept_selected)
    target=EXPORT_DIR/(fname+'.xlsx')
    export_xlsx(es,ds,target)
    return send_file(target,as_attachment=True,download_name=fname+'.xlsx')

@app.route('/export/dashboard')
def export_dashboard():
    if STATE['employees'] is None: return TOP+form()+BOTTOM
    dept=request.args.get('department','')
    fname='Дашборд по подразделениям' if not dept else 'Дашборд - '+safe_filename(dept)
    target=EXPORT_DIR/'dashboard_export.xlsx'
    export_dashboard_xlsx(STATE['employees'],STATE['days'],dept,target)
    return send_file(target,as_attachment=True,download_name=fname+'.xlsx')

if __name__=='__main__':
    port=int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0',port=port,debug=False,threaded=True,use_reloader=False)
