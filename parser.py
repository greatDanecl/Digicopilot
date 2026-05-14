"""
SDC Parser — lee todos los xlsx en /data y genera index.html.
Uso: python parser.py
"""
import pandas as pd, re, json, numpy as np, os, glob, sys
from datetime import timedelta
from collections import defaultdict, Counter
from pathlib import Path

DATA_DIR    = Path(__file__).parent / "data"
OUTPUT_HTML = Path(__file__).parent / "index.html"

MONTH_MAP = {
    'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
    'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12',
    'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
    'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12',
    'Ene':'01','Feb':'02','Mar':'03','Abr':'04','May':'05','Jun':'06',
    'Jul':'07','Ago':'08','Sep':'09','Oct':'10','Nov':'11','Dic':'12',
    'ene':'01','abr':'04','ago':'08','dic':'12',
}
PERIOD_LABELS_MAP = {
    '2025-01':'Ene 2025','2025-02':'Feb 2025','2025-03':'Mar 2025',
    '2025-04':'Abr 2025','2025-05':'May 2025','2025-06':'Jun 2025',
    '2025-07':'Jul 2025','2025-08':'Ago 2025','2025-09':'Sep 2025',
    '2025-10':'Oct 2025','2025-11':'Nov 2025','2025-12':'Dic 2025',
    '2026-01':'Ene 2026','2026-02':'Feb 2026','2026-03':'Mar 2026',
    '2026-04':'Abr 2026','2026-05':'May 2026','2026-06':'Jun 2026',
    '2026-07':'Jul 2026','2026-08':'Ago 2026','2026-09':'Sep 2026',
    '2026-10':'Oct 2026','2026-11':'Nov 2026','2026-12':'Dic 2026',
}

def parse_td(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return 0.0
    if isinstance(val, timedelta): return round(val.total_seconds() / 3600, 2)
    s = str(val)
    m = re.match(r'(\d+) days? (\d+):(\d+)', s)
    if m: return round(int(m.group(1))*24 + int(m.group(2)) + int(m.group(3))/60, 2)
    m2 = re.match(r'(\d+):(\d+)', s)
    if m2: return round(int(m2.group(1)) + int(m2.group(2))/60, 2)
    return 0.0

def detect_period_from_filename(fname):
    fl = fname.lower()
    m = re.search(r'(20\d{2})(0[1-9]|1[0-2])', fl)
    if m: return m.group(1) + '-' + m.group(2)
    m2 = re.search(r'(20\d{2})[-_](0[1-9]|1[0-2])', fl)
    if m2: return m2.group(1) + '-' + m2.group(2)
    month_re = r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|ene|abr|ago|dic)'
    year_re  = r'(20\d{2})'
    m3 = re.search(month_re + r'[_\-\s]*' + year_re, fl)
    if m3:
        code = MONTH_MAP.get(m3.group(1), '00')
        if code != '00': return m3.group(2) + '-' + code
    m4 = re.search(year_re + r'[_\-\s]*' + month_re, fl)
    if m4:
        code = MONTH_MAP.get(m4.group(2), '00')
        if code != '00': return m4.group(1) + '-' + code
    return None

def detect_period_from_df(df):
    try:
        cell = str(df.iloc[1, 2])
        m = re.search(r'(\d{2})(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Ene|Abr|Ago|Dic)(\d{2})', cell, re.I)
        if m:
            mon = m.group(2).capitalize()
            yr  = '20' + m.group(3)
            return yr + '-' + MONTH_MAP.get(mon, '00')
    except: pass
    try:
        for col in range(1, min(df.shape[1], 8)):
            cell = str(df.iloc[8, col])
            m = re.match(r'(\d{2})(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Ene|Abr|Ago|Dic)', cell, re.I)
            if m:
                mon = m.group(2).capitalize()
                for c2 in range(df.shape[1]):
                    yr_m = re.search(r'(20\d{2})', str(df.iloc[1, c2]))
                    if yr_m: return yr_m.group(1) + '-' + MONTH_MAP.get(mon, '00')
                return '2026-' + MONTH_MAP.get(mon, '00')
    except: pass
    return None

def detect_role(fname, sheet_name):
    fl, sl = fname.lower(), sheet_name.lower()
    if 'efec' in fl: return 'actual'
    if 'prog' in fl: return 'programmed'
    if 'efectuado' in fl or 'actual' in fl or 'flown' in fl: return 'actual'
    if 'programado' in fl or 'plan' in fl or 'sched' in fl:  return 'programmed'
    if any(w in sl for w in ['hora', 'actual', 'efect']): return 'actual'
    return 'programmed'

def classify_day(col_vals):
    clean = [v for v in col_vals if v not in ['nan', 'NaT', '']]
    if not clean: return 'BLANCO'
    joined = ' '.join(clean).upper()
    if re.search(r'\bTURNO\d+', joined): return 'TURNO'
    flights = [v for v in clean if re.match(r'^\d{2,4}$', v.strip())]
    if flights: return 'VUELO'
    for code, label in [
        ('HOTEL','HOTEL'),('SIM','SIM'),('ELEAR','ELEAR'),('DH','DH'),
        ('ACT','ACT'),('OFNA2','OFNA2'),('CEMAE','CEMAE'),('LQUIN','LQUIN'),
        ('SINDI','SINDI'),('FVUEL','FVUEL'),('VACAC','VACAC'),('BDAY','BDAY'),
        ('LIBRE','LIBRE'),('FINDE','FINDE'),
    ]:
        if code in joined: return label
    return 'CONT'

def get_day_cols(df, pilot_row, col, n_rows=4):
    col_start = 2
    return [
        str(df.iloc[pilot_row + k, col + col_start]).strip()
        if pilot_row + k < len(df) else ''
        for k in range(n_rows)
    ]

def count_schedule(df, pilot_row, role):
    n_cols = df.shape[1] - 2
    counts = {'turnos': 0, 'vuelos': 0, 'blancos': 0, 'hotel': 0, 'sim': 0, 'elear': 0, 'act': 0, 'dh': 0}
    for col in range(n_cols):
        vals = get_day_cols(df, pilot_row, col)
        tipo = classify_day(vals)
        if tipo == 'TURNO':  counts['turnos'] += 1
        elif tipo == 'VUELO': counts['vuelos'] += 1
        elif tipo == 'BLANCO': counts['blancos'] += 1
        elif tipo == 'HOTEL': counts['hotel'] += 1
        elif tipo == 'SIM':   counts['sim'] += 1
        elif tipo == 'ELEAR': counts['elear'] += 1
        elif tipo == 'ACT':   counts['act'] += 1
        elif tipo == 'DH':    counts['dh'] += 1
    return counts

def find_totals(df, pilot_row, max_look=16):
    cred_h = duty_h = blk_h = 0.0
    for k in range(1, max_look):
        row = pilot_row + k
        if row >= len(df): break
        lbl = str(df.iloc[row, 0]).strip()
        if re.match(r'^[A-Z]{4,5}$', lbl) and k > 5: break
        if lbl == 'Credits':       cred_h = parse_td(df.iloc[row, 1])
        elif lbl == 'Block hours': blk_h  = parse_td(df.iloc[row, 1])
        elif lbl == 'Duty hours':  duty_h = parse_td(df.iloc[row, 1])
    return cred_h, blk_h, duty_h

def block_size(df, pilot_row, max_look=18):
    for k in range(5, max_look):
        row = pilot_row + k
        if row >= len(df): return k
        c0 = str(df.iloc[row, 0]).strip()
        if re.match(r'^[A-Z]{4,5}$', c0): return k
    return 13

def parse_sheet(df, period, role):
    pilots = []
    if len(df) < 10 or df.shape[1] < 2: return pilots
    r9   = str(df.iloc[9, 0]).strip()
    abcd = bool(re.match(r'^[A-H]$', r9))
    i = 9
    while i < len(df):
        c0 = str(df.iloc[i, 0]).strip()
        if abcd:
            if c0 != 'A': i += 1; continue
            code    = str(df.iloc[i,   1]).strip()
            fname_p = str(df.iloc[i+1, 1]).strip() if i+1 < len(df) else ''
            lname   = str(df.iloc[i+2, 1]).strip() if i+2 < len(df) else ''
            rut_pos = str(df.iloc[i+3, 1]).strip() if i+3 < len(df) else ''
            base    = str(df.iloc[i+4, 1]).strip() if i+4 < len(df) else ''
            cred_h  = parse_td(df.iloc[i+5, 2] if i+5 < len(df) else None)
            blk_h   = parse_td(df.iloc[i+6, 2] if i+6 < len(df) else None)
            duty_h  = parse_td(df.iloc[i+7, 2] if i+7 < len(df) else None)
            sched   = [str(v).strip() for v in df.iloc[i, 2:].tolist()]
            pilot_row = i
            i += 8
        else:
            if not re.match(r'^[A-Z]{4,5}$', c0): i += 1; continue
            code    = c0
            fname_p = str(df.iloc[i+1, 0]).strip() if i+1 < len(df) else ''
            lname   = str(df.iloc[i+2, 0]).strip() if i+2 < len(df) else ''
            rut_pos = str(df.iloc[i+3, 0]).strip() if i+3 < len(df) else ''
            base    = str(df.iloc[i+4, 0]).strip() if i+4 < len(df) else ''
            cred_h, blk_h, duty_h = find_totals(df, i)
            sched   = [str(v).strip() for v in df.iloc[i, 2:].tolist()]
            pilot_row = i
            i += block_size(df, i)

        if not re.match(r'^[A-Z]{4,5}$', code): continue
        pos_raw = rut_pos.split(' - ')[-1].strip() if ' - ' in rut_pos else ''
        pos = pos_raw.split(',')[0].strip()
        if not pos or pos in ['nan', 'NaT', '']: continue
        name = (fname_p + ' ' + lname).strip()
        if not name or re.search(r'\b(TEST|PRUEBA)\b', name.upper()): continue

        pg = 'Otro'
        if pos in ['CP', 'CPN', 'C15M']:   pg = 'Capitán'
        elif pos in ['FO', 'FON']:          pg = 'Primer Oficial'
        elif pos in ['INS', 'INST', 'IOA']: pg = 'Instructor'

        vac  = sum(1 for s in sched if any(w in s.upper() for w in ['VACAC','VACAO','VACAP','VACAS']))
        med  = sum(1 for s in sched if any(w in s.upper() for w in ['LM','LICM','LICMED']))
        lib  = sum(1 for s in sched if s in ['LIBRE','FINDE'])
        sim  = sum(1 for s in sched if 'SIM' in s.upper())
        total_days = len([s for s in sched if s not in ['nan','NaT','','None']])
        excl = ((vac + med) / max(total_days, 1)) > 0.35 or blk_h < 5

        sched_counts = count_schedule(df, pilot_row, role)
        if role == 'programmed':
            turnos = sched_counts['turnos']
            vuelos_prog = sched_counts['vuelos']
            vuelos = None
            blancos = None
        else:
            turnos = None
            vuelos_prog = None
            vuelos  = sched_counts['vuelos']
            blancos = sched_counts['blancos']

        pilots.append({
            'period': period, 'role': role, 'code': code, 'name': name,
            'pos': pos, 'pos_group': pg, 'base': base,
            'block_h': blk_h, 'duty_h': duty_h, 'credits_h': cred_h,
            'libre_days': lib, 'vac_days': vac, 'med_days': med, 'sim_days': sim,
            'exclude_from_avg': excl,
            'turnos': turnos,
            'vuelos_prog': vuelos_prog,
            'vuelos': vuelos,
            'blancos': blancos,
        })
    return pilots

def build_dataset():
    xlsx_files = sorted(glob.glob(str(DATA_DIR / '*.xlsx')))
    if not xlsx_files:
        print('ERROR: No se encontraron archivos .xlsx en ' + str(DATA_DIR))
        sys.exit(1)

    config_path = DATA_DIR / 'config.json'
    file_map = {}
    if config_path.exists():
        cfg = json.loads(config_path.read_text(encoding='utf-8'))
        for entry in cfg.get('files', []):
            file_map[entry['filename']] = {'period': entry['period'], 'role': entry['role']}
        print('config.json cargado: ' + str(len(file_map)) + ' entradas\n')
    else:
        print('AVISO: no se encontro config.json en data/ — usando deteccion automatica\n')

    print('Procesando ' + str(len(xlsx_files)) + ' archivos...\n')
    all_records = []

    for fpath in xlsx_files:
        fname = os.path.basename(fpath)
        cfg_entry = file_map.get(fname)
        try:
            xl = pd.ExcelFile(fpath)
        except Exception as e:
            print('  x ' + fname + ': ' + str(e))
            continue

        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(fpath, sheet_name=sheet_name, header=None)
                if cfg_entry:
                    period = cfg_entry['period']
                else:
                    period = detect_period_from_filename(fname) or detect_period_from_df(df)
                if not period:
                    print('  ? ' + fname + '/' + sheet_name + ': periodo no detectado, omitiendo')
                    continue
                if cfg_entry:
                    sl = sheet_name.lower()
                    if any(w in sl for w in ['hora','actual','efect']):
                        role = 'actual'
                    else:
                        role = cfg_entry['role']
                else:
                    role = detect_role(fname, sheet_name)
                recs = parse_sheet(df, period, role)
                all_records.extend(recs)
                lbl = PERIOD_LABELS_MAP.get(period, period)
                print('  ok ' + fname + '/' + sheet_name + ': ' + lbl + ' ' + role + ' ' + str(len(recs)) + 'p')
            except Exception as e:
                print('  x ' + fname + '/' + sheet_name + ': ' + str(e))

    summary = {}
    for r in all_records:
        key = (r['period'], r['code'])
        if key not in summary:
            summary[key] = {
                'period': r['period'], 'code': r['code'], 'name': r['name'],
                'pos': r['pos'], 'pos_group': r['pos_group'], 'base': r['base'],
                'libre_days': r['libre_days'], 'vac_days': r['vac_days'],
                'med_days': r['med_days'], 'sim_days': r['sim_days'],
                'exclude_from_avg': r['exclude_from_avg'],
                'block_h_programmed': None, 'duty_h_programmed': None, 'credits_h_programmed': None,
                'block_h_actual':     None, 'duty_h_actual':     None, 'credits_h_actual':     None,
                'turnos_programados': None, 'vuelos_programados': None,
                'vuelos_efectuados':  None, 'dias_blancos':       None,
            }
        rk = r['role']
        for metric in ['block_h', 'duty_h', 'credits_h']:
            if r[metric] > 0:
                summary[key][metric + '_' + rk] = r[metric]
        if r['exclude_from_avg']:
            summary[key]['exclude_from_avg'] = True
        if rk == 'programmed' and r.get('turnos') is not None:
            summary[key]['turnos_programados'] = r['turnos']
        if rk == 'programmed' and r.get('vuelos_prog') is not None:
            summary[key]['vuelos_programados'] = r['vuelos_prog']
        if rk == 'actual' and r.get('vuelos') is not None:
            summary[key]['vuelos_efectuados'] = r['vuelos']
        if rk == 'actual' and r.get('blancos') is not None:
            summary[key]['dias_blancos'] = r['blancos']

    records = list(summary.values())
    periods = sorted(set(r['period'] for r in records))

    from collections import defaultdict
    names_by_grp = defaultdict(set)
    for r in records:
        names_by_grp[r['pos_group']].add(r['name'])

    print('\n' + '-'*50)
    print('Total registros: ' + str(len(records)))
    print('Periodos: ' + str([PERIOD_LABELS_MAP.get(p,p) for p in periods]))
    for g, ns in sorted(names_by_grp.items()):
        print('  ' + g + ': ' + str(len(ns)) + ' pilotos')

    return records, periods


def generate_html(records, periods):
    period_labels = {p: PERIOD_LABELS_MAP.get(p, p) for p in periods}
    DATA_JS    = json.dumps(records,       ensure_ascii=False, default=str)
    PERIODS_JS = json.dumps(periods,       ensure_ascii=False)
    LABELS_JS  = json.dumps(period_labels, ensure_ascii=False)

    CSS = (
        ':root{'
        '--purple:#671E77;--purple-l:#9B44B8;--purple-xl:#C480E0;'
        '--purple-dim:rgba(103,30,119,0.18);--purple-dim2:rgba(103,30,119,0.08);'
        '--green:#26D800;--green-l:#5CF200;--green-dim:rgba(38,216,0,0.15);--green-dim2:rgba(38,216,0,0.07);'
        '--violet:#8B35A8;--teal:#00C89B;--teal-dim:rgba(0,200,155,0.12);'
        '--danger:#E53E3E;--danger-dim:rgba(229,62,62,0.12);'
        '--warn:#C46AE0;--warn-dim:rgba(196,106,224,0.12);'
        '--bg:#F8F7FC;--surface:#FFFFFF;--s2:#F0EBF7;--s3:#E8DFF5;'
        '--border:rgba(103,30,119,0.18);--border2:rgba(103,30,119,0.35);'
        '--text:#2A1240;--text2:#5A3878;--muted:#8B6FA8;--dim:#B09CC8;'
        '--tooltip-bg:#2D1B45;--tooltip-border:rgba(155,68,184,0.5);'
        '--tooltip-title:#F0E8F8;--tooltip-body:#D4B8EE;--tooltip-accent:#A7F3D0;'
        '--r:10px;--r2:14px;'
        '--shadow:0 1px 4px rgba(0,0,0,.08),0 4px 20px rgba(103,30,119,.10);'
        '--shadow2:0 2px 12px rgba(0,0,0,.12),0 8px 32px rgba(103,30,119,.18);'
        "--font:'DM Sans',sans-serif;--display:'Playfair Display',serif;--mono:'DM Mono',monospace;"
        '}'
        '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}'
        'html{font-size:14px}'
        'body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.5}'
        '.shell{display:grid;grid-template-columns:230px 1fr;min-height:100vh}'
        '.sidebar{background:#FFFFFF;display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;border-right:1px solid var(--border)}'
        '.sidebar-top{padding:0 0 14px;border-bottom:1px solid rgba(103,30,119,.15)}'
        '.logo-wrap{width:100%;background:#FFFFFF;display:flex;align-items:center;justify-content:center;padding:14px 18px}'
        '.logo-wrap img{width:100%;max-width:192px;height:auto;display:block}'
        '.brand-sub-line{font-size:9px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase;text-align:center;padding:5px 0 0;font-family:var(--mono)}'
        '.filters{padding:14px 16px;display:flex;flex-direction:column;gap:11px;border-bottom:1px solid rgba(103,30,119,.15)}'
        '.f-block{display:flex;flex-direction:column;gap:5px}'
        '.f-label{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono)}'
        '.f-select{appearance:none;background:rgba(103,30,119,.05);border:1px solid rgba(103,30,119,.25);border-radius:8px;color:var(--text);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;background-image:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'12\' height=\'12\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%238B6FA8\' stroke-width=\'2\'%3E%3Cpolyline points=\'6 9 12 15 18 9\'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 9px center}'
        '.f-select:focus,.f-select:hover{border-color:var(--green);box-shadow:0 0 0 2px rgba(38,216,0,.15)}'
        '.f-select option{background:#FFFFFF;color:var(--text)}'
        '.sidebar-nav{padding:10px 8px;flex:1}'
        '.nav-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:var(--text2);cursor:pointer;transition:all .15s;margin-bottom:2px;border-left:2px solid transparent;user-select:none}'
        '.nav-item:hover{color:var(--purple);background:rgba(103,30,119,.08);border-left-color:rgba(103,30,119,.4)}'
        '.nav-item.active{color:var(--purple);background:rgba(103,30,119,.12);border-left-color:var(--purple)}'
        '.nav-item svg{width:14px;height:14px;flex-shrink:0}'
        '.sidebar-footer{padding:12px 16px;border-top:1px solid rgba(103,30,119,.15)}'
        '.pilot-badge{display:flex;align-items:center;gap:10px}'
        '.pilot-avatar{width:34px;height:34px;border-radius:50%;background:var(--purple);border:1.5px solid var(--purple-l);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}'
        '.pilot-name-s{font-size:11px;font-weight:500;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
        '.pilot-pos-s{font-size:10px;color:var(--muted);font-family:var(--mono)}'
        '.main{display:flex;flex-direction:column;min-height:100vh}'
        '.topbar{background:#FFFFFF;border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;box-shadow:0 1px 8px rgba(103,30,119,.06)}'
        '.page-title{font-family:var(--display);font-size:17px;color:var(--text)}'
        '.page-title span{color:var(--purple)}'
        '.page-sub{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.topbar-right{display:flex;align-items:center;gap:8px}'
        '.pill{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--s2);color:var(--text2)}'
        '.dot{width:6px;height:6px;border-radius:50%;background:var(--green);box-shadow:0 0 7px var(--green)}'
        '.content{padding:18px 26px;display:flex;flex-direction:column;gap:13px;flex:1}'
        '.view-section{display:none;flex-direction:column;gap:16px}'
        '.view-section.active{display:flex}'
        '.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}'
        '.kpi-grid-6{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}'
        '.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 15px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:box-shadow .2s,transform .15s,border-color .2s}'
        '.kpi:hover{box-shadow:var(--shadow2);transform:translateY(-1px);border-color:var(--border2)}'
        ".kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;border-radius:var(--r2) var(--r2) 0 0}"
        '.kpi.k-p1::before{background:var(--purple-l)}'
        '.kpi.k-p2::before{background:var(--violet)}'
        '.kpi.k-g1::before{background:var(--green)}'
        '.kpi.k-g2::before{background:var(--teal)}'
        '.kpi.k-g3::before{background:var(--green-l)}'
        '.kpi.k-r1::before{background:var(--danger)}'
        '.kpi-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;font-family:var(--mono)}'
        '.kpi-val{font-size:24px;font-weight:600;color:var(--text);font-family:var(--mono);letter-spacing:-.03em;line-height:1}'
        '.kpi-unit{font-size:12px;font-weight:400;color:var(--muted);margin-left:2px}'
        '.kpi-footer{display:flex;align-items:center;justify-content:space-between;margin-top:7px}'
        '.kpi-vs{font-size:10px;color:var(--muted)}.kpi-vs b{color:var(--text2);font-weight:500}'
        '.delta{font-size:10px;font-family:var(--mono);padding:2px 6px;border-radius:4px}'
        '.d-up{background:var(--green-dim);color:#1A9900}'
        '.d-down{background:var(--danger-dim);color:var(--danger)}'
        '.d-neu{background:var(--s3);color:var(--muted)}'
        '.d-warn{background:var(--warn-dim);color:#8B22AA}'
        '.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow)}'
        '.card-head{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}'
        '.card-title{font-size:13px;font-weight:500;color:var(--text)}'
        '.card-sub{font-size:10px;color:var(--muted);margin-top:2px;font-family:var(--mono)}'
        '.legend{display:flex;gap:12px;align-items:center;font-size:10px;color:var(--muted);font-family:var(--mono);flex-wrap:wrap}'
        '.leg{display:flex;align-items:center;gap:5px}'
        '.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}'
        '.chart-wrap{position:relative;height:220px}'
        '.chart-wrap-lg{position:relative;height:300px}'
        '.comp-table{width:100%;border-collapse:collapse;font-size:12px}'
        '.comp-table th{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--text2);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--s2)}'
        '.comp-table td{padding:8px 10px;border-bottom:1px solid rgba(103,30,119,.2)}'
        '.comp-table tr:last-child td{border-bottom:none}'
        '.comp-table tr:hover td{background:var(--s2)}'
        '.bottom-row{display:grid;grid-template-columns:1fr 300px;gap:14px}'
        '.prog-list{display:flex;flex-direction:column;gap:13px}'
        '.prog-head{display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px}'
        '.prog-lbl{color:var(--text2)}.prog-num{font-family:var(--mono);font-size:11px}'
        '.prog-track{height:5px;background:var(--s3);border-radius:3px;overflow:hidden}'
        '.prog-fill{height:100%;border-radius:3px}'
        '.prog-note{font-size:9px;color:var(--dim);margin-top:2px;font-family:var(--mono)}'
        '.alert-list{display:flex;flex-direction:column;gap:7px}'
        '.alert{display:flex;align-items:flex-start;gap:9px;padding:9px 11px;border-radius:8px;border:1px solid}'
        '.alert.ok{background:rgba(38,216,0,.07);border-color:rgba(38,216,0,.3)}'
        '.alert.warn{background:rgba(196,106,224,.10);border-color:rgba(139,53,168,.3)}'
        '.alert.danger{background:rgba(229,62,62,.10);border-color:rgba(229,62,62,.35)}'
        '.alert-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px}'
        '.alert.ok .alert-dot{background:var(--green)}'
        '.alert.warn .alert-dot{background:#8B22AA}'
        '.alert.danger .alert-dot{background:var(--danger)}'
        '.alert-title{font-size:11px;font-weight:500;color:var(--text)}'
        '.alert-desc{font-size:10px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.excl-note{display:flex;align-items:center;gap:6px;padding:8px 11px;border-radius:7px;background:var(--s2);border:1px solid var(--border);font-size:10px;color:var(--muted);margin-top:10px}'
        '.excl-note svg{width:12px;height:12px;flex-shrink:0}'
        '.dan-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}'
        '.dan-card{border-radius:var(--r2);padding:20px;border:2px solid;position:relative;overflow:hidden}'
        '.dan-card.ok{background:rgba(38,216,0,.06);border-color:rgba(38,216,0,.35)}'
        '.dan-card.warn{background:rgba(196,106,224,.10);border-color:rgba(139,53,168,.35)}'
        '.dan-card.danger{background:rgba(229,62,62,.10);border-color:rgba(229,62,62,.40)}'
        '.dan-label{font-size:10px;font-family:var(--mono);text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:6px}'
        '.dan-val{font-size:28px;font-weight:700;font-family:var(--mono);letter-spacing:-.04em;line-height:1;margin-bottom:4px}'
        '.dan-card.ok .dan-val{color:#1A7A00}'
        '.dan-card.warn .dan-val{color:#6B1A8A}'
        '.dan-card.danger .dan-val{color:#C0392B}'
        '.dan-limit{font-size:11px;color:var(--muted);font-family:var(--mono)}'
        '.dan-bar-wrap{margin-top:10px;height:6px;background:rgba(0,0,0,.08);border-radius:3px;overflow:hidden}'
        '.dan-bar-fill{height:100%;border-radius:3px}'
        '.dan-card.ok .dan-bar-fill{background:var(--green)}'
        '.dan-card.warn .dan-bar-fill{background:#8B22AA}'
        '.dan-card.danger .dan-bar-fill{background:var(--danger)}'
        '.stat-row{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}'
        '::-webkit-scrollbar{width:4px}'
        '::-webkit-scrollbar-thumb{background:rgba(103,30,119,.3);border-radius:2px}'
        '@keyframes fadeUp{from{opacity:0;transform:translateY(7px)}to{opacity:1;transform:none}}'
        '.kpi,.card,.dan-card{animation:fadeUp .28s ease both}'
        '.hamburger{display:none;position:fixed;top:12px;left:12px;z-index:200;width:38px;height:38px;border-radius:8px;border:1px solid var(--border2);background:var(--s2);cursor:pointer;align-items:center;justify-content:center;flex-direction:column;gap:5px;padding:9px}'
        '.hamburger span{display:block;width:16px;height:1.5px;background:var(--text2);border-radius:2px;transition:all .2s}'
        '.hamburger.open span:nth-child(1){transform:translateY(6.5px) rotate(45deg)}'
        '.hamburger.open span:nth-child(2){opacity:0}'
        '.hamburger.open span:nth-child(3){transform:translateY(-6.5px) rotate(-45deg)}'
        '.sidebar-overlay{display:none;position:fixed;inset:0;background:rgba(40,20,60,.5);z-index:155;backdrop-filter:blur(2px)}'
        '.sidebar-overlay.open{display:block}'
        '@media(max-width:1024px){.kpi-grid{grid-template-columns:repeat(3,1fr)}.kpi-grid-6{grid-template-columns:repeat(3,1fr)}.charts-row{grid-template-columns:1fr}.bottom-row{grid-template-columns:1fr}.dan-grid{grid-template-columns:1fr}.stat-row{grid-template-columns:repeat(2,1fr)}}'
        '@media(max-width:768px){.shell{grid-template-columns:1fr}.sidebar{position:fixed;left:-240px;top:0;height:100vh;width:230px;z-index:160;transition:left .25s cubic-bezier(.4,0,.2,1)}.sidebar.open{left:0;box-shadow:6px 0 32px rgba(46,36,22,.3)}.hamburger{display:flex}.topbar{padding:12px 16px 12px 58px}.content{padding:14px 16px}.kpi-grid{grid-template-columns:repeat(2,1fr);gap:8px}.kpi-grid-6{grid-template-columns:repeat(2,1fr);gap:8px}.charts-row{grid-template-columns:1fr;gap:10px}.bottom-row{grid-template-columns:1fr;gap:10px}.chart-wrap{height:190px}.chart-wrap-lg{height:240px}.page-title{font-size:14px}.page-sub{font-size:10px;margin-top:0}.card{padding:14px}.card-head{flex-direction:column;gap:8px;align-items:flex-start}.legend{gap:8px;font-size:9px}#compTableWrap{overflow-x:auto;-webkit-overflow-scrolling:touch}.comp-table th,.comp-table td{padding:6px 8px}.dan-grid{grid-template-columns:1fr}.stat-row{grid-template-columns:1fr 1fr}}'
        '@media(max-width:420px){.kpi-grid{grid-template-columns:1fr 1fr}.kpi-grid-6{grid-template-columns:1fr 1fr}.kpi-val{font-size:21px}.kpi{padding:12px 12px}.chart-wrap{height:165px}.chart-wrap-lg{height:200px}.content{padding:10px 12px}}'
    )
    return CSS


def build_js(DATA_JS, PERIODS_JS, LABELS_JS):
    return (
        'const RAW = ' + DATA_JS + ';\n'
        'const PERIODS = ' + PERIODS_JS + ';\n'
        'const PERIOD_LABELS = ' + LABELS_JS + ';\n'
        '\n'
        "document.getElementById('periodPill').textContent = Object.values(PERIOD_LABELS).join(' \u00b7 ');\n"
        "document.getElementById('periodsHint').textContent = 'Per\u00edodos: ' + Object.values(PERIOD_LABELS).join(' \u00b7 ');\n"
        '\n'
        '// ── TOOLTIP DEFAULTS ────────────────────────────────\n'
        'const TOOLTIP_DEFAULTS = {\n'
        "  backgroundColor:'#FFFFFF',\n"
        "  borderColor:'rgba(103,30,119,0.3)',\n"
        "  borderWidth:1,\n"
        "  titleColor:'#2A1240',\n"
        "  bodyColor:'#5A3878',\n"
        "  footerColor:'#671E77',\n"
        "  padding:12,\n"
        "  titleFont:{family:\"'DM Sans',sans-serif\",size:12,weight:600},\n"
        "  bodyFont:{family:\"'DM Mono',monospace\",size:11},\n"
        "  footerFont:{family:\"'DM Mono',monospace\",size:11,weight:500},\n"
        "  boxShadow:'0 4px 20px rgba(103,30,119,0.15)',\n"
        "  cornerRadius:8,\n"
        '};\n'
        '\n'
        'let blockChartInst = null, compareChartInst = null, dutyChartInst = null, dutyCompInst = null;\n'
        'let currentView = "resumen";\n'
        'const selGroup = document.getElementById("selGroup");\n'
        'const selPilot = document.getElementById("selPilot");\n'
        'const selMonth = document.getElementById("selMonth");\n'
        '\n'
        '// ── NAV ──────────────────────────────────────────────\n'
        'document.querySelectorAll(".nav-item").forEach(item => {\n'
        '  item.addEventListener("click", () => {\n'
        '    if (!selPilot.value) return;\n'
        '    const view = item.dataset.view;\n'
        '    if (!view || view === currentView) return;\n'
        '    currentView = view;\n'
        '    document.querySelectorAll(".nav-item").forEach(n => n.classList.remove("active"));\n'
        '    item.classList.add("active");\n'
        '    document.querySelectorAll(".view-section").forEach(s => s.classList.remove("active"));\n'
        '    document.getElementById("view-" + view).classList.add("active");\n'
        '    const pt = {resumen:"Resumen",bloque:"Block Hours",deber:"Duty Hours",dan:"DAN 121"};\n'
        '    document.getElementById("pageTitle").innerHTML = "<span>" + selPilot.value.split(" ").slice(0,2).join(" ") + "</span> \u00b7 " + (pt[view]||view);\n'
        '    if (window.innerWidth <= 768) closeMenu();\n'
        '  });\n'
        '});\n'
        '\n'
        'selGroup.addEventListener("change", () => {\n'
        '  const g = selGroup.value;\n'
        '  const names = [...new Set(RAW.filter(r => r.pos_group === g).map(r => r.name))].sort((a,b) => a.localeCompare(b, "es"));\n'
        '  selPilot.innerHTML = \'<option value="">— Seleccionar tripulante —</option>\';\n'
        '  names.forEach(n => { const o = document.createElement("option"); o.value = o.textContent = n; selPilot.appendChild(o); });\n'
        '  selPilot.disabled = false;\n'
        '  selMonth.innerHTML = \'<option value="">— Seleccione un tripulante —</option>\';\n'
        '  selMonth.disabled = true;\n'
        '  document.getElementById("placeholder").style.display = "flex";\n'
        '  document.getElementById("dashboard").style.display = "none";\n'
        '});\n'
        '\n'
        'selPilot.addEventListener("change", () => {\n'
        '  if (selPilot.value) render(selPilot.value, selGroup.value);\n'
        '});\n'
        '\n'
        'selMonth.addEventListener("change", () => {\n'
        '  if (selMonth.value && selPilot.value) {\n'
        '    renderKPIs(selPilot.value, selGroup.value, selMonth.value);\n'
        '    renderDAN(selPilot.value, selGroup.value, selMonth.value);\n'
        '  }\n'
        '});\n'
        '\n'
        'function fmt(v, d) { d = d === undefined ? 1 : d; if (v == null || +v === 0) return "\u2014"; return (+v).toFixed(d); }\n'
        'function avg(arr) { const v = arr.filter(x => x != null && x > 0); return v.length ? v.reduce((a,b) => a+b, 0)/v.length : 0; }\n'
        'function dc(d) { return d > 2 ? "d-up" : d < -2 ? "d-down" : "d-neu"; }\n'
        'function ds(d) { return (d >= 0 ? "+" : "") + d.toFixed(1) + "%"; }\n'
        'function bestBlock(r) { return (r.block_h_actual && r.block_h_actual > 0) ? r.block_h_actual : (r.block_h_programmed || 0); }\n'
        'function isProgrammedOnly(r) { return !(r.block_h_actual && r.block_h_actual > 0) && (r.block_h_programmed && r.block_h_programmed > 0); }\n'
        'function makeGrad(ctx, ca, c1, c2) {\n'
        '  if (!ca) return "transparent";\n'
        '  const g = ctx.createLinearGradient(0, ca.top, 0, ca.bottom);\n'
        '  g.addColorStop(0, c1); g.addColorStop(1, c2); return g;\n'
        '}\n'
        '\n'
        'function render(pilotName, group) {\n'
        '  document.getElementById("placeholder").style.display = "none";\n'
        '  document.getElementById("dashboard").style.display = "flex";\n'
        '  const pr = RAW.filter(r => r.name === pilotName);\n'
        '  const gr = RAW.filter(r => r.pos_group === group);\n'
        '  const latest = pr.filter(r => r.block_h_actual > 0).sort((a,b) => b.period.localeCompare(a.period))[0] || pr.sort((a,b) => b.period.localeCompare(a.period))[0];\n'
        '  const init = pilotName.split(" ").filter((_,i) => i < 2).map(w => w[0]).join("");\n'
        '  document.getElementById("sideAvatar").textContent = init;\n'
        '  document.getElementById("sideName").textContent = pilotName.split(" ").slice(0,2).join(" ");\n'
        '  document.getElementById("sidePos").textContent = (latest ? latest.pos : group) + " \u00b7 " + (latest ? latest.base : "");\n'
        '  document.getElementById("pageTitle").innerHTML = "<span>" + pilotName.split(" ").slice(0,2).join(" ") + "</span> \u00b7 Resumen";\n'
        '  document.getElementById("pageSub").textContent = (latest ? latest.pos_group : group) + " \u00b7 " + (latest ? latest.base : "") + " \u00b7 " + Object.values(PERIOD_LABELS).join(" \u00b7 ");\n'
        '  // Reset nav to resumen\n'
        '  currentView = "resumen";\n'
        '  document.querySelectorAll(".nav-item").forEach(n => n.classList.remove("active"));\n'
        '  document.querySelector(\'[data-view="resumen"]\').classList.add("active");\n'
        '  document.querySelectorAll(".view-section").forEach(s => s.classList.remove("active"));\n'
        '  document.getElementById("view-resumen").classList.add("active");\n'
        '\n'
        '  // Populate month dropdown\n'
        '  const pilotPeriods = [...new Set(pr.map(r => r.period))].sort().reverse();\n'
        '  selMonth.innerHTML = "";\n'
        '  pilotPeriods.forEach(p => {\n'
        '    const o = document.createElement("option");\n'
        '    o.value = p;\n'
        '    const r = pr.find(x => x.period === p);\n'
        '    const hasBoth = r && r.block_h_actual > 0 && r.block_h_programmed > 0;\n'
        '    const hasAct  = r && r.block_h_actual > 0;\n'
        '    const tag = hasBoth ? " (prog+ef)" : hasAct ? " (ef)" : " (prog)";\n'
        '    o.textContent = (PERIOD_LABELS[p] || p) + tag;\n'
        '    selMonth.appendChild(o);\n'
        '  });\n'
        '  const defaultPeriod = (latest ? latest.period : pilotPeriods[0]);\n'
        '  selMonth.value = defaultPeriod;\n'
        '  selMonth.disabled = false;\n'
        '\n'
        '  renderCharts(pilotName, group, pr, gr);\n'
        '  renderKPIs(pilotName, group, defaultPeriod);\n'
        '  renderDutyView(pilotName, group, pr, gr);\n'
        '  renderDAN(pilotName, group, defaultPeriod);\n'
        '}\n'
        '\n'
        '// ── VIEW: RESUMEN KPIs ───────────────────────────────\n'
        'function renderKPIs(pilotName, group, selectedPeriod) {\n'
        '  const pr = RAW.filter(r => r.name === pilotName);\n'
        '  const gr = RAW.filter(r => r.pos_group === group);\n'
        '  const sel = pr.find(r => r.period === selectedPeriod) || pr[0];\n'
        '  const lp  = selectedPeriod;\n'
        '  const ga = gr.filter(r => r.period === lp && r.name !== pilotName && !r.exclude_from_avg && bestBlock(r) > 0);\n'
        '  const ab = avg(ga.map(r => r.block_h_actual || 0).filter(v => v > 0));\n'
        '  const ad = avg(ga.map(r => r.duty_h_actual  || 0).filter(v => v > 0));\n'
        '  const al = avg(ga.map(r => r.libre_days     || 0).filter(v => v > 0));\n'
        '  const mb = sel ? (sel.block_h_actual || sel.block_h_programmed || 0) : 0;\n'
        '  const md = sel ? (sel.duty_h_actual  || sel.duty_h_programmed  || 0) : 0;\n'
        '  const ml = sel ? (sel.libre_days     || 0) : 0;\n'
        '  const isProg = sel && !(sel.block_h_actual > 0);\n'
        '  const bd = ab > 0 ? (mb-ab)/ab*100 : 0;\n'
        '  const dd = ad > 0 ? (md-ad)/ad*100 : 0;\n'
        '  const actP  = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);\n'
        '  const accB  = actP.reduce((s,r) => s + (r.block_h_actual || 0), 0);\n'
        '  const excl  = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '  const turnos  = sel ? (sel.turnos_programados || null) : null;\n'
        '  const vuelos  = sel ? (sel.vuelos_efectuados  || null) : null;\n'
        '  const vProg   = sel ? (sel.vuelos_programados || null) : null;\n'
        '  const blancos = sel ? (sel.dias_blancos        || null) : null;\n'
        '  const progTag = isProg ? \' <span style="font-size:9px;color:var(--muted);font-family:var(--mono)">(prog.)</span>\' : "";\n'
        '\n'
        '  document.getElementById("kpiRow").innerHTML =\n'
        '    \'<div class="kpi k-p1"><div class="kpi-label">Block Hours \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + fmt(mb) + \'<span class="kpi-unit">h</span>\' + progTag + \'</div><div class="kpi-footer"><span class="kpi-vs">Group avg: <b>\' + fmt(ab) + \'h</b></span><span class="delta \' + dc(bd) + \'">\' + ds(bd) + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-p2"><div class="kpi-label">Duty Hours \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + fmt(md) + \'<span class="kpi-unit">h</span>\' + progTag + \'</div><div class="kpi-footer"><span class="kpi-vs">Group avg: <b>\' + fmt(ad) + \'h</b></span><span class="delta \' + dc(dd) + \'">\' + ds(dd) + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-g1"><div class="kpi-label">D\u00edas libres \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + ml + \'<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Group avg: <b>\' + fmt(al,0) + \'d</b></span><span class="delta \' + dc(ml-al) + \'">\' + (ml-al>=0?"+":"") + (ml-al).toFixed(0) + \'d</span></div></div>\' +\n'
        '    \'<div class="kpi k-g2"><div class="kpi-label">Turnos prog. \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + (turnos !== null ? turnos : "\\u2014") + \'<span class="kpi-unit">\' + (vuelos !== null ? " / "+vuelos+" ef." : "") + \'</span></div><div class="kpi-footer"><span class="kpi-vs">\' + (vProg !== null ? vProg+" vuelos prog." : "Sin datos prog.") + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-g3"><div class="kpi-label">Block Hours acum. YTD</div><div class="kpi-val">\' + fmt(accB,0) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">\' + actP.length + \' meses activos</span><span class="delta d-neu">/\' + PERIODS.length + \'m</span></div></div>\' +\n'
        '    \'<div class="kpi k-p2"><div class="kpi-label">D\u00edas blancos \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + (blancos !== null ? blancos : "\\u2014") + \'<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Sin asignaci\u00f3n</span><span class="delta \' + (blancos > 5 ? "d-warn" : "d-up") + \'">\' + (blancos !== null ? (blancos > 5 ? "\\u26a0 revisar" : "\\u2713 ok") : "\\u2014") + \'</span></div></div>\';\n'
        '\n'
        '  const pct1 = Math.min(accB/1000*100, 100);\n'
        '  const avgM = actP.length ? accB/actP.length : 0;\n'
        '  const proj = avgM * 12;\n'
        '  const pctP = Math.min(proj/1000*100, 100);\n'
        '  const totL = pr.reduce((s,r) => s+(r.libre_days||0), 0);\n'
        '  const avgL = pr.length ? totL/pr.length : 0;\n'
        "  document.getElementById('progList').innerHTML =\n"
        '    \'<div><div class="prog-head"><span class="prog-lbl">Block Hours acumuladas</span><span class="prog-num" style="color:var(--purple)">\' + accB.toFixed(0) + \'h</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pct1 + \'%;background:var(--purple)"></div></div><div class="prog-note">L\u00edmite DAN 121: 1.000h/a\u00f1o \u00b7 \' + (100-pct1).toFixed(1) + \'% disponible</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Proyecci\u00f3n 12 meses</span><span class="prog-num" style="color:var(--violet)">~\' + proj.toFixed(0) + \'h est.</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pctP + \'%;background:linear-gradient(90deg,var(--green),var(--teal))"></div></div><div class="prog-note">Prom. \' + avgM.toFixed(1) + \'h/mes en meses activos</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Descanso promedio</span><span class="prog-num" style="color:var(--teal)">\' + avgL.toFixed(1) + \' d/mes</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + Math.min(avgL/20*100,100) + \'%;background:var(--teal)"></div></div><div class="prog-note">M\u00ednimo reglamentario DAN 121: 8 d\u00edas/mes</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Meses activos</span><span class="prog-num">\' + actP.length + \' / \' + PERIODS.length + \'</span></div><div style="display:flex;gap:3px;margin-top:4px"><div style="height:5px;border-radius:2px 0 0 2px;background:var(--teal);flex:\' + actP.length + \'"></div><div style="height:5px;border-radius:0 2px 2px 0;background:rgba(103,30,119,.25);flex:\' + Math.max(PERIODS.length-actP.length,0) + \'"></div></div><div class="prog-note">\' + (excl.length?excl.map(p=>PERIOD_LABELS[p]||p).join(", ")+" excluidos":"Sin ausencias prolongadas") + \'</div></div>\';\n'
        '}\n'
    )


def build_js_part2():
    return (
        '\n// ── VIEW: RESUMEN CHARTS ────────────────────────────\n'
        'function renderCharts(pilotName, group, pr, gr) {\n'
        '  const excl  = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '  const progOnlyIdx = PERIODS.map((p,i) => { const r = pr.find(x => x.period===p); return (r && isProgrammedOnly(r)) ? i : -1; }).filter(i => i>=0);\n'
        '  const pData = PERIODS.map(p => { const r = pr.find(x => x.period===p); return r ? bestBlock(r) : null; });\n'
        '  const gData = PERIODS.map(p => {\n'
        '    const peers = gr.filter(r => r.name !== pilotName && !r.exclude_from_avg && bestBlock(r) > 0);\n'
        '    const inPeriod = peers.filter(r => r.period === p);\n'
        '    return inPeriod.length ? avg(inPeriod.map(r => bestBlock(r))) : null;\n'
        '  });\n'
        "  const bc = document.getElementById('blockChart').getContext('2d');\n"
        '  if (blockChartInst) blockChartInst.destroy();\n'
        '  blockChartInst = new Chart(bc, {\n'
        "    type: 'line',\n"
        '    data: { labels: PERIODS.map(p => PERIOD_LABELS[p]||p), datasets: [\n'
        "      { label:'Piloto', data:pData, borderColor:'#26D800',\n"
        "        backgroundColor(c) { return makeGrad(bc, c.chart.chartArea, 'rgba(38,216,0,.15)', 'rgba(38,216,0,.01)'); },\n"
        '        borderWidth:2.5,\n'
        '        pointRadius(c)          { return excl.includes(PERIODS[c.dataIndex]) ? 6 : 4; },\n'
        "        pointStyle(c)           { return excl.includes(PERIODS[c.dataIndex]) ? 'triangle' : progOnlyIdx.includes(c.dataIndex) ? 'rectRot' : 'circle'; },\n"
        "        pointBackgroundColor(c) { return excl.includes(PERIODS[c.dataIndex]) ? '#5CF200' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#26D800'; },\n"
        "        pointBorderColor(c)     { return excl.includes(PERIODS[c.dataIndex]) ? '#5CF200' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#26D800'; },\n"
        '        pointHoverRadius:7, tension:.35, fill:true, spanGaps:true, order:1 },\n'
        "      { label:'Group avg', data:gData, borderColor:'#9B44B8', borderWidth:1.5, borderDash:[5,4],\n"
        "        pointBackgroundColor:'#9B44B8', pointRadius:3, pointHoverRadius:5,\n"
        '        tension:.35, fill:false, spanGaps:false, order:2 }\n'
        '    ]},\n'
        '    options: { responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},\n'
        '      plugins: { legend:{display:false}, tooltip:{\n'
        '        ...TOOLTIP_DEFAULTS,\n'
        '        callbacks:{\n'
        '          title(i)     { const p=PERIODS[i[0].dataIndex]; const ex=excl.includes(p); const po=progOnlyIdx.includes(i[0].dataIndex); return (PERIOD_LABELS[p]||p)+(ex?" \u00b7 \u26a0 excluido del prom.":po?" \u00b7 solo programado":""); },\n'
        '          label(i)     { if(i.raw==null||i.raw===0)return null; const po=progOnlyIdx.includes(i.dataIndex)&&i.datasetIndex===0; return "  "+i.dataset.label+(po?" (prog.)":"")+": "+i.raw.toFixed(1)+"h"; },\n'
        '          afterBody(i) { const p=PERIODS[i[0].dataIndex]; const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex]; if(av==null||my==null||my===0)return[]; const d=my-av; return["  vs group avg: "+(d>=0?"+":"")+d.toFixed(1)+"h"]; }\n'
        '        }\n'
        '      }},\n'
        "      scales:{x:{grid:{color:'rgba(103,30,119,.12)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},y:{min:0,grid:{color:'rgba(103,30,119,.12)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"},callback:function(v){return v+'h';}},border:{display:false}}}\n"
        '    }\n'
        '  });\n'
        '\n'
        "  const en = document.getElementById('exclNote');\n"
        '  const ep = excl.map(p => PERIOD_LABELS[p]||p).filter(Boolean);\n'
        "  if (ep.length) { en.style.display='flex'; document.getElementById('exclText').textContent='Meses excluidos del promedio: '+ep.join(', ')+'. Mostrados como tri\u00e1ngulo en el gr\u00e1fico.'; }\n"
        "  else en.style.display='none';\n"
        '\n'
        '  // Bar chart\n'
        '  const prog = PERIODS.map(p => { const r=pr.find(x=>x.period===p); return r?(r.block_h_programmed||0):0; });\n'
        '  const act  = PERIODS.map(p => { const r=pr.find(x=>x.period===p); return r?(r.block_h_actual||0):0; });\n'
        "  const cc = document.getElementById('compareChart').getContext('2d');\n"
        '  if (compareChartInst) compareChartInst.destroy();\n'
        '  compareChartInst = new Chart(cc, {\n'
        "    type:'bar',\n"
        '    data:{labels:PERIODS.map(p=>PERIOD_LABELS[p]||p),datasets:[\n'
        "      {label:'Programado',data:prog,backgroundColor:'rgba(155,68,184,.45)',borderColor:'#9B44B8',borderWidth:1,borderRadius:5,borderSkipped:false},\n"
        "      {label:'Efectuado', data:act, backgroundColor:'rgba(38,216,0,.4)',borderColor:'#26D800',borderWidth:1,borderRadius:5,borderSkipped:false}\n"
        '    ]},\n'
        '    options:{responsive:true,maintainAspectRatio:false,\n'
        '      plugins:{legend:{display:false},tooltip:{\n'
        '        ...TOOLTIP_DEFAULTS,\n'
        '        callbacks:{\n'
        '          label(i){return "  "+i.dataset.label+": "+i.raw.toFixed(1)+"h";},\n'
        '          afterBody(i){const idx=i[0].dataIndex;const d=act[idx]-prog[idx];if(prog[idx]===0&&act[idx]===0)return["  Sin datos"];const w=d>.5?"\u2191 Efectuado mayor":d<-.5?"\u2193 Programado mayor":"\u2248 Similares";return["\u0394: "+(d>=0?"+":"")+d.toFixed(1)+"h  "+w];}\n'
        '        }\n'
        '      }},\n'
        "      scales:{x:{grid:{display:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},y:{min:0,grid:{color:'rgba(103,30,119,.12)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"},callback:function(v){return v+'h';}},border:{display:false}}}\n"
        '    }\n'
        '  });\n'
        '\n'
        '  // Comparison table\n'
        '  let tbl = \'<table class="comp-table"><thead><tr><th>Per\u00edodo</th><th>Block prog.</th><th>Block ef.</th><th>\u0394 Block</th><th>Turnos</th><th>Vuelos prog.</th><th>Vuelos ef.</th><th>Blancos</th></tr></thead><tbody>\';\n'
        '  PERIODS.forEach(p => {\n'
        '    const r = pr.find(x => x.period===p); if(!r) return;\n'
        '    const pg=r.block_h_programmed||0, ac=r.block_h_actual||0, d=ac-pg;\n'
        '    const tp=r.turnos_programados, vp=r.vuelos_programados;\n'
        '    const ve=r.vuelos_efectuados,  bl=r.dias_blancos;\n'
        '    const ex = r.exclude_from_avg ? \'<span style="color:var(--warn);font-size:9px"> \u2731</span>\' : "";\n'
        '    const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\u2014";\n'
        '    const dclr = d >= 0 ? "var(--teal)" : "var(--danger)";\n'
        '    const blCell = bl !== null ? (bl > 5 ? \'<span style="color:var(--danger)">\'+bl+\'</span>\' : bl) : "\u2014";\n'
        '    tbl += \'<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">\' + (PERIOD_LABELS[p]||p) + ex + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (pg>0?pg.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ac>0?ac.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + dclr + \'">\' + dstr + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (tp !== null ? tp : "\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (vp !== null ? vp : "\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ve !== null ? ve : "\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + blCell + \'</td>\'\n'
        '         + \'</tr>\';\n'
        '  });\n'
        '  if (excl.length) tbl += \'<tr><td colspan="8" style="font-size:9px;color:var(--muted);font-family:var(--mono);padding:6px 10px">\u2731 Excluido del promedio</td></tr>\';\n'
        '  tbl += "</tbody></table>";\n'
        "  document.getElementById('compTableWrap').innerHTML = tbl;\n"
        '}\n'
    )


def build_js_part3():
    return (
        '\n// ── VIEW: BLOCK HOURS ───────────────────────────────\n'
        '// (renderCharts ya construye los gráficos del resumen;\n'
        '// la vista Bloque reutiliza datos pero con canvas propios)\n'
        'function renderBlockView(pilotName, group, pr, gr) {\n'
        '  const excl = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '  const progOnlyIdx = PERIODS.map((p,i) => { const r = pr.find(x => x.period===p); return (r && isProgrammedOnly(r)) ? i : -1; }).filter(i => i>=0);\n'
        '  const pData = PERIODS.map(p => { const r = pr.find(x => x.period===p); return r ? bestBlock(r) : null; });\n'
        '  const gData = PERIODS.map(p => {\n'
        '    const peers = gr.filter(r => r.name !== pilotName && !r.exclude_from_avg && bestBlock(r) > 0);\n'
        '    const inP = peers.filter(r => r.period === p);\n'
        '    return inP.length ? avg(inP.map(r => bestBlock(r))) : null;\n'
        '  });\n'
        '  // Percentile band (25th–75th) across group\n'
        '  const p25 = PERIODS.map(p => {\n'
        '    const peers = gr.filter(r => r.name !== pilotName && !r.exclude_from_avg && bestBlock(r) > 0 && r.period === p).map(r => bestBlock(r)).sort((a,b)=>a-b);\n'
        '    return peers.length >= 4 ? peers[Math.floor(peers.length*0.25)] : null;\n'
        '  });\n'
        '  const p75 = PERIODS.map(p => {\n'
        '    const peers = gr.filter(r => r.name !== pilotName && !r.exclude_from_avg && bestBlock(r) > 0 && r.period === p).map(r => bestBlock(r)).sort((a,b)=>a-b);\n'
        '    return peers.length >= 4 ? peers[Math.floor(peers.length*0.75)] : null;\n'
        '  });\n'
        "  const ctx = document.getElementById('blockViewChart').getContext('2d');\n"
        '  if (window._blockViewInst) window._blockViewInst.destroy();\n'
        '  window._blockViewInst = new Chart(ctx, {\n'
        "    type:'line',\n"
        '    data:{ labels: PERIODS.map(p => PERIOD_LABELS[p]||p), datasets:[\n'
        "      { label:'Piloto', data:pData, borderColor:'#26D800',\n"
        "        backgroundColor(c){return makeGrad(ctx,c.chart.chartArea,'rgba(38,216,0,.18)','rgba(38,216,0,.01)');},\n"
        '        borderWidth:2.5, tension:.35, fill:true, spanGaps:true,\n'
        '        pointRadius(c){return excl.includes(PERIODS[c.dataIndex])?7:5;},\n'
        "        pointStyle(c){return excl.includes(PERIODS[c.dataIndex])?'triangle':progOnlyIdx.includes(c.dataIndex)?'rectRot':'circle';},\n"
        "        pointBackgroundColor(c){return excl.includes(PERIODS[c.dataIndex])?'#5CF200':progOnlyIdx.includes(c.dataIndex)?'#9B7EC8':'#26D800';},\n"
        '        pointHoverRadius:8, order:1 },\n'
        "      { label:'Group avg', data:gData, borderColor:'#9B44B8', borderWidth:2, borderDash:[6,4],\n"
        "        pointBackgroundColor:'#9B44B8', pointRadius:4, pointHoverRadius:6,\n"
        '        tension:.35, fill:false, spanGaps:false, order:2 },\n'
        "      { label:'P75', data:p75, borderColor:'rgba(155,68,184,.25)', borderWidth:1, borderDash:[2,3],\n"
        "        pointRadius:0, tension:.35, fill:false, spanGaps:true, order:3 },\n"
        "      { label:'P25', data:p25, borderColor:'rgba(155,68,184,.25)', borderWidth:1, borderDash:[2,3],\n"
        "        backgroundColor:'rgba(155,68,184,.06)', pointRadius:0, tension:.35, fill:'-1', spanGaps:true, order:4 }\n"
        '    ]},\n'
        '    options:{ responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},\n'
        '      plugins:{ legend:{display:false}, tooltip:{\n'
        '        ...TOOLTIP_DEFAULTS,\n'
        '        callbacks:{\n'
        '          title(i){ const p=PERIODS[i[0].dataIndex]; return (PERIOD_LABELS[p]||p)+(excl.includes(p)?" \u00b7 \u26a0 excluido":""); },\n'
        '          label(i){ if(i.datasetIndex>1)return null; if(i.raw==null||i.raw===0)return null; const po=progOnlyIdx.includes(i.dataIndex)&&i.datasetIndex===0; return "  "+i.dataset.label+(po?" (prog.)":"")+": "+i.raw.toFixed(1)+"h"; },\n'
        '          afterBody(i){ const p=PERIODS[i[0].dataIndex]; const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex],lo=p25[i[0].dataIndex],hi=p75[i[0].dataIndex]; const lines=[]; if(av!=null&&my!=null&&my>0){const d=my-av;lines.push("  vs avg: "+(d>=0?"+":"")+d.toFixed(1)+"h");} if(lo!=null&&hi!=null)lines.push("  rango P25\u2013P75: "+lo.toFixed(1)+"\u2013"+hi.toFixed(1)+"h"); return lines; }\n'
        '        }\n'
        '      }},\n'
        "      scales:{x:{grid:{color:'rgba(103,30,119,.10)'},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},y:{min:0,grid:{color:'rgba(103,30,119,.10)'},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"},callback:v=>v+'h'},border:{display:false}}}\n"
        '    }\n'
        '  });\n'
        '\n'
        '  // Block stats cards\n'
        '  const actVals = pData.filter((v,i)=>v!=null&&!excl.includes(PERIODS[i]));\n'
        '  const blockMax = actVals.length ? Math.max(...actVals) : 0;\n'
        '  const blockMin = actVals.length ? Math.min(...actVals) : 0;\n'
        '  const blockAvg = actVals.length ? actVals.reduce((a,b)=>a+b,0)/actVals.length : 0;\n'
        '  const actB = pr.filter(r=>!r.exclude_from_avg&&r.block_h_actual>0).reduce((s,r)=>s+(r.block_h_actual||0),0);\n'
        '  const prog12 = blockAvg*12;\n'
        "  document.getElementById('blockStats').innerHTML =\n"
        '    \'<div class="kpi k-g1"><div class="kpi-label">Block Hours \u00b7 Promedio</div><div class="kpi-val">\' + fmt(blockAvg) + \'<span class="kpi-unit">h/mes</span></div><div class="kpi-footer"><span class="kpi-vs">Meses activos sin ausencias</span></div></div>\' +\n'
        '    \'<div class="kpi k-p1"><div class="kpi-label">Block Hours \u00b7 M\u00e1ximo</div><div class="kpi-val">\' + fmt(blockMax) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Mes de mayor actividad</span></div></div>\' +\n'
        '    \'<div class="kpi k-p2"><div class="kpi-label">Block Hours \u00b7 M\u00ednimo</div><div class="kpi-val">\' + fmt(blockMin) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Mes de menor actividad</span></div></div>\' +\n'
        '    \'<div class="kpi k-g3"><div class="kpi-label">Acumulado YTD</div><div class="kpi-val">\' + fmt(actB,0) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">L\u00edmite anual: <b>1.000h</b></span><span class="delta \' + (actB>900?"d-down":actB>750?"d-warn":"d-up") + \'">\' + (1000-actB).toFixed(0) + \'h disp.</span></div></div>\';\n'
        '}\n'
    )


def build_js_part4():
    return (
        '\n// ── VIEW: DUTY HOURS ────────────────────────────────\n'
        'function renderDutyView(pilotName, group, pr, gr) {\n'
        '  const excl = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '  const dData = PERIODS.map(p => { const r=pr.find(x=>x.period===p); return r?(r.duty_h_actual||r.duty_h_programmed||null):null; });\n'
        '  const gDuty = PERIODS.map(p => {\n'
        '    const peers = gr.filter(r => r.name!==pilotName && !r.exclude_from_avg && (r.duty_h_actual||0)>0 && r.period===p);\n'
        '    return peers.length ? avg(peers.map(r=>r.duty_h_actual||0)) : null;\n'
        '  });\n'
        '  // Ratio duty/block per period\n'
        '  const ratioData = PERIODS.map(p => {\n'
        '    const r=pr.find(x=>x.period===p);\n'
        '    if(!r) return null;\n'
        '    const bh=bestBlock(r), dh=r.duty_h_actual||r.duty_h_programmed||0;\n'
        '    return bh>0&&dh>0 ? parseFloat((dh/bh).toFixed(2)) : null;\n'
        '  });\n'
        '  const gRatio = PERIODS.map(p => {\n'
        '    const peers = gr.filter(r=>r.name!==pilotName&&!r.exclude_from_avg&&bestBlock(r)>0&&(r.duty_h_actual||0)>0&&r.period===p);\n'
        '    return peers.length ? parseFloat(avg(peers.map(r=>(r.duty_h_actual||0)/bestBlock(r))).toFixed(2)) : null;\n'
        '  });\n'
        '\n'
        "  const dc1 = document.getElementById('dutyChart').getContext('2d');\n"
        '  if (dutyChartInst) dutyChartInst.destroy();\n'
        '  dutyChartInst = new Chart(dc1, {\n'
        "    type:'line',\n"
        '    data:{ labels:PERIODS.map(p=>PERIOD_LABELS[p]||p), datasets:[\n'
        "      { label:'Duty Hours', data:dData, borderColor:'#9B44B8',\n"
        "        backgroundColor(c){return makeGrad(dc1,c.chart.chartArea,'rgba(155,68,184,.18)','rgba(155,68,184,.01)');},\n"
        '        borderWidth:2.5, tension:.35, fill:true, spanGaps:true,\n'
        "        pointBackgroundColor:'#9B44B8', pointRadius:4, pointHoverRadius:7, order:1 },\n"
        "      { label:'Group avg', data:gDuty, borderColor:'#00C89B', borderWidth:1.5, borderDash:[5,4],\n"
        "        pointBackgroundColor:'#00C89B', pointRadius:3, pointHoverRadius:5,\n"
        '        tension:.35, fill:false, spanGaps:false, order:2 }\n'
        '    ]},\n'
        '    options:{ responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},\n'
        '      plugins:{ legend:{display:false}, tooltip:{\n'
        '        ...TOOLTIP_DEFAULTS,\n'
        '        callbacks:{\n'
        '          title(i){ return PERIOD_LABELS[PERIODS[i[0].dataIndex]]||PERIODS[i[0].dataIndex]; },\n'
        '          label(i){ if(i.raw==null||i.raw===0)return null; return "  "+i.dataset.label+": "+i.raw.toFixed(1)+"h"; },\n'
        '          afterBody(i){ const my=dData[i[0].dataIndex],av=gDuty[i[0].dataIndex]; if(!av||!my)return[]; const d=my-av; return["  vs avg: "+(d>=0?"+":"")+d.toFixed(1)+"h"]; }\n'
        '        }\n'
        '      }},\n'
        "      scales:{x:{grid:{color:'rgba(103,30,119,.10)'},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},y:{min:0,grid:{color:'rgba(103,30,119,.10)'},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"},callback:v=>v+'h'},border:{display:false}}}\n"
        '    }\n'
        '  });\n'
        '\n'
        "  const dc2 = document.getElementById('dutyRatioChart').getContext('2d');\n"
        '  if (dutyCompInst) dutyCompInst.destroy();\n'
        '  dutyCompInst = new Chart(dc2, {\n'
        "    type:'bar',\n"
        '    data:{ labels:PERIODS.map(p=>PERIOD_LABELS[p]||p), datasets:[\n'
        "      { label:'Ratio Duty/Block \u00b7 Piloto', data:ratioData, backgroundColor:'rgba(155,68,184,.5)', borderColor:'#9B44B8', borderWidth:1, borderRadius:5 },\n"
        "      { label:'Ratio Duty/Block \u00b7 Grupo', data:gRatio, backgroundColor:'rgba(0,200,155,.4)', borderColor:'#00C89B', borderWidth:1, borderRadius:5 }\n"
        '    ]},\n'
        '    options:{ responsive:true, maintainAspectRatio:false,\n'
        '      plugins:{ legend:{display:false}, tooltip:{\n'
        '        ...TOOLTIP_DEFAULTS,\n'
        '        callbacks:{\n'
        '          title(i){ return PERIOD_LABELS[PERIODS[i[0].dataIndex]]||PERIODS[i[0].dataIndex]; },\n'
        '          label(i){ if(i.raw==null)return null; return "  "+i.dataset.label+": "+i.raw.toFixed(2)+"x"; },\n'
        '          afterBody(i){ const rp=ratioData[i[0].dataIndex],rg=gRatio[i[0].dataIndex]; if(!rp||!rg)return[]; const d=rp-rg; return["  vs avg: "+(d>=0?"+":"")+d.toFixed(2)+"x"]; }\n'
        '        }\n'
        '      }},\n'
        "      scales:{x:{grid:{display:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},y:{min:0,grid:{color:'rgba(103,30,119,.10)'},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}}}\n"
        '    }\n'
        '  });\n'
        '\n'
        '  // Duty stats\n'
        '  const dutyVals = dData.filter((v,i)=>v!=null&&!excl.includes(PERIODS[i]));\n'
        '  const dutyAvg  = dutyVals.length ? dutyVals.reduce((a,b)=>a+b,0)/dutyVals.length : 0;\n'
        '  const dutyMax  = dutyVals.length ? Math.max(...dutyVals) : 0;\n'
        '  const dutyMin  = dutyVals.length ? Math.min(...dutyVals) : 0;\n'
        '  const ratioAvg = ratioData.filter(v=>v!=null);\n'
        '  const rAvg = ratioAvg.length ? ratioAvg.reduce((a,b)=>a+b,0)/ratioAvg.length : 0;\n'
        "  document.getElementById('dutyStats').innerHTML =\n"
        '    \'<div class="kpi k-p2"><div class="kpi-label">Duty Hours \u00b7 Promedio</div><div class="kpi-val">\' + fmt(dutyAvg) + \'<span class="kpi-unit">h/mes</span></div><div class="kpi-footer"><span class="kpi-vs">Meses activos</span></div></div>\' +\n'
        '    \'<div class="kpi k-p1"><div class="kpi-label">Duty Hours \u00b7 M\u00e1ximo</div><div class="kpi-val">\' + fmt(dutyMax) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Mes de mayor actividad</span></div></div>\' +\n'
        '    \'<div class="kpi k-g2"><div class="kpi-label">Duty Hours \u00b7 M\u00ednimo</div><div class="kpi-val">\' + fmt(dutyMin) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Mes de menor actividad</span></div></div>\' +\n'
        '    \'<div class="kpi k-g3"><div class="kpi-label">Ratio Duty / Block</div><div class="kpi-val">\' + fmt(rAvg,2) + \'<span class="kpi-unit">x</span></div><div class="kpi-footer"><span class="kpi-vs">Promedio del per\u00edodo</span><span class="delta \' + (rAvg>1.8?"d-warn":"d-neu") + \'">\' + (rAvg>1.8?"\u26a0 alto":"normal") + \'</span></div></div>\';\n'
        '}\n'
    )


def build_js_part5():
    return (
        '\n// ── VIEW: DAN 121 ───────────────────────────────────\n'
        'function renderDAN(pilotName, group, selectedPeriod) {\n'
        '  const pr  = RAW.filter(r => r.name === pilotName);\n'
        '  const gr  = RAW.filter(r => r.pos_group === group);\n'
        '  const sel = pr.find(r => r.period === selectedPeriod) || pr[0];\n'
        '  const lp  = selectedPeriod;\n'
        '\n'
        '  const actP = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);\n'
        '  const accB = actP.reduce((s,r) => s + (r.block_h_actual||0), 0);\n'
        '  const avgM = actP.length ? accB/actP.length : 0;\n'
        '  const proj12 = avgM * 12;\n'
        '\n'
        '  const mb = sel ? (sel.block_h_actual || sel.block_h_programmed || 0) : 0;\n'
        '  const md = sel ? (sel.duty_h_actual  || sel.duty_h_programmed  || 0) : 0;\n'
        '  const ml = sel ? (sel.libre_days || 0) : 0;\n'
        '  const excl = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '\n'
        '  function danStatus(val, warn, danger) { return val >= danger ? "danger" : val >= warn ? "warn" : "ok"; }\n'
        '  function danStatusLow(val, warn, danger) { return val <= danger ? "danger" : val <= warn ? "warn" : "ok"; }\n'
        '\n'
        '  const s1 = danStatus(mb, 85, 100);\n'
        '  const s2 = danStatus(accB, 750, 900);\n'
        '  const s3 = danStatus(proj12, 800, 950);\n'
        '  const s4 = danStatusLow(ml, 10, 8);\n'
        '  const s5 = danStatus(md, 105, 130);\n'
        '  const s6 = danStatus(accB/Math.max(actP.length,1), 85, 100);\n'
        '\n'
        '  function danCard(status, label, val, unit, limit, pct, note) {\n'
        '    const icons = {ok:\'<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M20 6L9 17l-5-5"/></svg>\',\n'
        '                   warn:\'<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 9v4m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>\',\n'
        '                   danger:\'<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>\'};\n'
        '    return \'<div class="dan-card \'+status+\'">\'\n'
        '           +\'<div style="display:flex;justify-content:space-between;align-items:flex-start">\'\n'
        '           +\'<div class="dan-label">\'+label+\'</div>\'\n'
        '           +\'<div style="color:\'+{ok:"#1A7A00",warn:"#6B1A8A",danger:"#C0392B"}[status]+\'">\'+icons[status]+\'</div></div>\'\n'
        '           +\'<div class="dan-val">\'+val+\'<span style="font-size:14px;font-weight:400;margin-left:3px">\'+unit+\'</span></div>\'\n'
        '           +\'<div class="dan-limit">\'+limit+\'</div>\'\n'
        '           +\'<div class="dan-bar-wrap"><div class="dan-bar-fill" style="width:\'+Math.min(pct,100)+\'%"></div></div>\'\n'
        '           +\'<div style="font-size:10px;color:var(--muted);margin-top:8px;font-family:var(--mono)">\'+note+\'</div>\'\n'
        '           +\'</div>\';\n'
        '  }\n'
        '\n'
        "  document.getElementById('danCards').innerHTML =\n"
        '    danCard(s1,"Block Hours \u00b7 " + (PERIOD_LABELS[lp]||lp), fmt(mb),"h","L\u00edmite mensual: 100h",mb/100*100, mb>100?"Excede el l\u00edmite DAN 121 \u00b7 Art. 121.500":mb>85?"Cercano al l\u00edmite mensual de 100h":"Dentro del rango permitido") +\n'
        '    danCard(s2,"Block Hours acum. YTD", fmt(accB,0),"h","L\u00edmite anual: 1.000h",accB/1000*100, accB>900?"Muy pr\u00f3ximo al l\u00edmite anual":accB>750?"Supera el 75% del cupo anual":"Cupo restante: "+(1000-accB).toFixed(0)+"h") +\n'
        '    danCard(s3,"Proyecci\u00f3n 12 meses", "~"+fmt(proj12,0),"h est.","Proyectado sobre avg mensual",proj12/1000*100, proj12>950?"Proyecci\u00f3n supera el l\u00edmite anual":proj12>800?"Proyecci\u00f3n sobre el 80% del cupo":"Proyecci\u00f3n dentro del cupo anual") +\n'
        '    danCard(s4,"D\u00edas libres \u00b7 " + (PERIOD_LABELS[lp]||lp), ml,"d","M\u00ednimo reglamentario: 8d/mes",(ml/20)*100, ml<8?"Bajo el m\u00ednimo DAN 121 \u00b7 Art. 121.485":ml<10?"Sobre el m\u00ednimo pero bajo el promedio del cargo":"Descanso adecuado seg\u00fan DAN 121") +\n'
        '    danCard(s5,"Duty Hours \u00b7 " + (PERIOD_LABELS[lp]||lp), fmt(md),"h","Referencia: 130h/mes",md/130*100, md>130?"Duty hours muy elevadas, revisar FDPs":md>105?"Sobre el promedio del cargo":"Dentro del rango normal") +\n'
        '    danCard(s6,"Block Hours prom. mensual", fmt(avgM),"h/mes","Promedio meses activos",avgM/100*100, avgM>95?"Promedio mensual muy alto, vigilar acumulado":avgM>80?"Nivel sostenido alto":"Nivel de actividad normal");\n'
        '\n'
        '  // History table\n'
        '  let htbl = \'<table class="comp-table"><thead><tr>\'\n'
        '    + \'<th>Per\u00edodo</th><th>Block ef.</th><th>Duty ef.</th><th>D. libres</th><th>D. blancos</th><th>Estado</th></tr></thead><tbody>\';\n'
        '  PERIODS.forEach(p => {\n'
        '    const r = pr.find(x => x.period===p); if(!r) return;\n'
        '    const bh = r.block_h_actual || r.block_h_programmed || 0;\n'
        '    const dh = r.duty_h_actual  || r.duty_h_programmed  || 0;\n'
        '    const lib = r.libre_days || 0;\n'
        '    const bl  = r.dias_blancos;\n'
        '    const ex  = r.exclude_from_avg;\n'
        '    const isProg = !(r.block_h_actual > 0) && bh > 0;\n'
        '    const bSt = bh > 100 ? "danger" : bh > 85 ? "warn" : "ok";\n'
        '    const lSt = lib < 8 ? "danger" : lib < 10 ? "warn" : "ok";\n'
        '    const overall = (bSt==="danger"||lSt==="danger") ? "danger" : (bSt==="warn"||lSt==="warn") ? "warn" : "ok";\n'
        '    const dot = {ok:\'<span style="color:#1A7A00">&#9679;</span>\',warn:\'<span style="color:#6B1A8A">&#9679;</span>\',danger:\'<span style="color:#C0392B">&#9679;</span>\'};\n'
        '    const tag = isProg ? \' <span style="font-size:9px;color:var(--muted)">(prog.)</span>\' : "";\n'
        '    const exTag = ex ? \' <span style="font-size:9px;color:var(--warn)">\u2731</span>\' : "";\n'
        '    htbl += \'<tr>\'\n'
        '         + \'<td style="font-family:var(--mono);font-size:11px">\' + (PERIOD_LABELS[p]||p) + exTag + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + (bh>100?"var(--danger)":bh>85?"#6B1A8A":"var(--text2)") + \'">\' + (bh>0?bh.toFixed(1)+"h":"\u2014") + tag + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (dh>0?dh.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + (lib<8?"var(--danger)":lib<10?"#6B1A8A":"var(--text2)") + \'">\' + lib + \'d</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (bl!==null?bl+\'d\':"\u2014") + \'</td>\'\n'
        '         + \'<td>\' + dot[overall] + \'</td>\'\n'
        '         + \'</tr>\';\n'
        '  });\n'
        '  htbl += "</tbody></table>";\n'
        "  document.getElementById('danHistory').innerHTML = htbl;\n"
        '  if(excl.length){ document.getElementById("danExclNote").style.display="flex"; document.getElementById("danExclText").textContent="Meses excluidos del promedio: "+excl.map(p=>PERIOD_LABELS[p]||p).join(", ")+"."; }\n'
        '  else document.getElementById("danExclNote").style.display="none";\n'
        '}\n'
        '\n'
        '// ── MOBILE NAV ───────────────────────────────────────\n'
        'const menuBtn = document.getElementById("menuBtn");\n'
        'const sidebar  = document.getElementById("sidebar");\n'
        'const overlay  = document.getElementById("overlay");\n'
        'function openMenu()  { sidebar.classList.add("open"); overlay.classList.add("open"); menuBtn.classList.add("open"); document.body.style.overflow="hidden"; }\n'
        'function closeMenu() { sidebar.classList.remove("open"); overlay.classList.remove("open"); menuBtn.classList.remove("open"); document.body.style.overflow=""; }\n'
        'menuBtn.addEventListener("click", () => sidebar.classList.contains("open") ? closeMenu() : openMenu());\n'
        'overlay.addEventListener("click", closeMenu);\n'
        'document.getElementById("selPilot").addEventListener("change", () => { if(window.innerWidth <= 768) closeMenu(); });\n'
    )


def build_html(CSS, JS, LOGO_B64=""):
    # Logo placeholder — usar el base64 real del logo en producción
    LOGO = LOGO_B64 if LOGO_B64 else "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 200 60'%3E%3Crect width='200' height='60' fill='%23671E77' rx='8'/%3E%3Ctext x='100' y='38' text-anchor='middle' fill='white' font-size='18' font-family='Arial' font-weight='bold'%3ESPSKY%3C/text%3E%3C/svg%3E"

    return (
        '<!DOCTYPE html>\n'
        '<html lang="es">\n'
        '<head>\n'
        '<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        '<title>SPSKY Digital Copilot \u00b7 Productividad de Tripulaci\u00f3n</title>\n'
        '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
        '<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">\n'
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>\n'
        '<style>\n' + CSS + '\n</style>\n'
        '</head>\n'
        '<body>\n'
        '<div class="shell">\n'
        '<button class="hamburger" id="menuBtn" aria-label="Abrir men\u00fa"><span></span><span></span><span></span></button>\n'
        '<div class="sidebar-overlay" id="overlay"></div>\n'
        '<div class="sidebar" id="sidebar">\n'
        '  <div class="sidebar-top">\n'
        '    <div class="logo-wrap"><img src="' + LOGO + '" alt="SPSKY Pilotos"></div>\n'
        '    <div class="brand-sub-line">Digital Copilot</div>\n'
        '  </div>\n'
        '  <div class="filters">\n'
        '    <div class="f-block"><div class="f-label">Cargo</div>\n'
        '      <select class="f-select" id="selGroup">\n'
        '        <option value="">\u2014 Seleccionar cargo \u2014</option>\n'
        '        <option value="Capit\u00e1n">Capit\u00e1n</option>\n'
        '        <option value="Primer Oficial">Primer Oficial</option>\n'
        '        <option value="Instructor">Instructor</option>\n'
        '      </select>\n'
        '    </div>\n'
        '    <div class="f-block"><div class="f-label">Tripulante</div>\n'
        '      <select class="f-select" id="selPilot" disabled>\n'
        '        <option value="">\u2014 Seleccione un cargo primero \u2014</option>\n'
        '      </select>\n'
        '    </div>\n'
        '    <div class="f-block"><div class="f-label">Mes (KPIs)</div>\n'
        '      <select class="f-select" id="selMonth" disabled>\n'
        '        <option value="">\u2014 Seleccione un tripulante \u2014</option>\n'
        '      </select>\n'
        '    </div>\n'
        '  </div>\n'
        '  <nav class="sidebar-nav">\n'
        '    <div class="nav-item active" data-view="resumen"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>Resumen</div>\n'
        '    <div class="nav-item" data-view="bloque"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>Block Hours</div>\n'
        '    <div class="nav-item" data-view="deber"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>Duty Hours</div>\n'
        '    <div class="nav-item" data-view="dan"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>DAN 121</div>\n'
        '  </nav>\n'
        '  <div class="sidebar-footer">\n'
        '    <div class="pilot-badge">\n'
        '      <div class="pilot-avatar" id="sideAvatar">\u2014</div>\n'
        '      <div><div class="pilot-name-s" id="sideName">Sin selecci\u00f3n</div><div class="pilot-pos-s" id="sidePos">\u2014</div></div>\n'
        '    </div>\n'
        '  </div>\n'
        '</div>\n'
        '<div class="main">\n'
        '  <div class="topbar">\n'
        '    <div>\n'
        '      <div class="page-title" id="pageTitle">Seleccione un <span>tripulante</span></div>\n'
        '      <div class="page-sub" id="pageSub">SDC \u00b7 Productividad de Tripulaci\u00f3n</div>\n'
        '    </div>\n'
        '    <div class="topbar-right">\n'
        '      <div class="pill"><span class="dot"></span>Sistema activo</div>\n'
        '      <div class="pill" id="periodPill">\u2014</div>\n'
        '    </div>\n'
        '  </div>\n'
        '  <div class="content">\n'

        # ── PLACEHOLDER ──
        '    <div id="placeholder" style="display:flex;flex-direction:column;align-items:center;justify-content:center;flex:1;gap:14px;color:var(--dim);padding:60px 0;">\n'
        '      <svg width="44" height="44" fill="none" stroke="currentColor" stroke-width="1.2" viewBox="0 0 24 24" style="stroke:var(--border2)"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/></svg>\n'
        '      <div style="font-family:var(--display);font-size:18px;color:var(--text2)">SDC \u00b7 SPSKY Digital Copilot</div>\n'
        '      <div style="font-size:12px;text-align:center;max-width:300px;line-height:1.7;color:var(--muted)">Seleccione un cargo y un tripulante para visualizar sus indicadores de productividad.</div>\n'
        '      <div style="font-size:10px;font-family:var(--mono);color:var(--dim);margin-top:4px" id="periodsHint"></div>\n'
        '    </div>\n'

        # ── DASHBOARD WRAPPER ──
        '    <div id="dashboard" style="display:none;flex-direction:column;gap:16px;">\n'

        # ── VIEW: RESUMEN ──
        '      <div class="view-section active" id="view-resumen">\n'
        '        <div class="kpi-grid" id="kpiRow"></div>\n'
        '        <div class="card">\n'
        '          <div class="card-head">\n'
        '            <div><div class="card-title">Block Hours \u00b7 Evoluci\u00f3n mensual</div><div class="card-sub">Piloto vs. promedio del cargo (meses activos)</div></div>\n'
        '            <div class="legend">\n'
        '              <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#26D800" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="#26D800"/></svg><span>Efectuado</span></div>\n'
        '              <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#9B7EC8" stroke-width="1.5" stroke-dasharray="2 2"/><rect x="5.5" y="1.5" width="5" height="5" transform="rotate(45 9 4)" fill="#9B7EC8"/></svg><span style="color:var(--muted)">Solo programado</span></div>\n'
        '              <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#9B44B8" stroke-width="1.5" stroke-dasharray="5 4"/><circle cx="9" cy="4" r="2.5" fill="#9B44B8"/></svg><span>Group avg</span></div>\n'
        '              <div class="leg"><svg width="14" height="12"><polygon points="7,1 13,11 1,11" fill="none" stroke="#9B44B8" stroke-width="1.5"/></svg><span style="color:var(--muted)">Excluido prom.</span></div>\n'
        '            </div>\n'
        '          </div>\n'
        '          <div class="chart-wrap"><canvas id="blockChart"></canvas></div>\n'
        '          <div class="excl-note" id="exclNote" style="display:none"><svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg><span id="exclText"></span></div>\n'
        '        </div>\n'
        '        <div class="charts-row">\n'
        '          <div class="card">\n'
        '            <div class="card-head"><div><div class="card-title">Block Hours \u00b7 Programado vs. Efectuado</div><div class="card-sub">Por per\u00edodo</div></div><div class="legend"><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:rgba(155,68,184,.5);display:inline-block"></span><span>Programado</span></div><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:rgba(38,216,0,.4);display:inline-block"></span><span>Efectuado</span></div></div></div>\n'
        '            <div class="chart-wrap"><canvas id="compareChart"></canvas></div>\n'
        '          </div>\n'
        '          <div class="card"><div class="card-head"><div class="card-title">Comparativo por Per\u00edodo</div><div class="card-sub">Block prog. vs. ef. \u00b7 \u0394 horas</div></div><div id="compTableWrap" style="overflow-x:auto"></div></div>\n'
        '        </div>\n'
        '        <div class="bottom-row">\n'
        '          <div class="card"><div class="card-head"><div class="card-title">Acumulado &amp; Proyecci\u00f3n</div><div class="card-sub">Basado en meses activos</div></div><div class="prog-list" id="progList"></div></div>\n'
        '          <div class="card"><div class="card-head"><div class="card-title">Cumplimiento DAN 121</div><div class="card-sub">Mes seleccionado</div></div><div class="alert-list" id="alertListResumen"></div></div>\n'
        '        </div>\n'
        '      </div>\n'

        # ── VIEW: BLOCK HOURS ──
        '      <div class="view-section" id="view-bloque">\n'
        '        <div class="kpi-grid" id="blockStats"></div>\n'
        '        <div class="card">\n'
        '          <div class="card-head">\n'
        '            <div><div class="card-title">Block Hours \u00b7 Evoluci\u00f3n completa</div><div class="card-sub">Con banda percentil P25\u2013P75 del cargo</div></div>\n'
        '            <div class="legend">\n'
        '              <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#26D800" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="#26D800"/></svg><span>Piloto</span></div>\n'
        '              <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#9B44B8" stroke-width="2" stroke-dasharray="6 4"/><circle cx="9" cy="4" r="3" fill="#9B44B8"/></svg><span>Group avg</span></div>\n'
        '              <div class="leg"><span style="width:28px;height:8px;background:rgba(155,68,184,.15);border:1px dashed rgba(155,68,184,.4);display:inline-block;border-radius:2px"></span><span style="color:var(--muted)">Rango P25\u2013P75</span></div>\n'
        '            </div>\n'
        '          </div>\n'
        '          <div class="chart-wrap-lg"><canvas id="blockViewChart"></canvas></div>\n'
        '        </div>\n'
        '        <div class="card"><div class="card-head"><div class="card-title">Detalle por Per\u00edodo</div><div class="card-sub">Block prog. vs. ef. \u00b7 comparativo</div></div><div id="blockDetailTable" style="overflow-x:auto"></div></div>\n'
        '      </div>\n'

        # ── VIEW: DUTY HOURS ──
        '      <div class="view-section" id="view-deber">\n'
        '        <div class="kpi-grid" id="dutyStats"></div>\n'
        '        <div class="charts-row">\n'
        '          <div class="card">\n'
        '            <div class="card-head"><div><div class="card-title">Duty Hours \u00b7 Evoluci\u00f3n mensual</div><div class="card-sub">Piloto vs. promedio del cargo</div></div>\n'
        '              <div class="legend"><div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#9B44B8" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="#9B44B8"/></svg><span>Duty Hours</span></div><div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="#00C89B" stroke-width="1.5" stroke-dasharray="5 4"/><circle cx="9" cy="4" r="3" fill="#00C89B"/></svg><span>Group avg</span></div></div>\n'
        '            </div>\n'
        '            <div class="chart-wrap"><canvas id="dutyChart"></canvas></div>\n'
        '          </div>\n'
        '          <div class="card">\n'
        '            <div class="card-head"><div><div class="card-title">Ratio Duty / Block Hours</div><div class="card-sub">Por per\u00edodo \u00b7 Piloto vs. grupo</div></div>\n'
        '              <div class="legend"><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:rgba(155,68,184,.5);display:inline-block"></span><span>Piloto</span></div><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:rgba(0,200,155,.4);display:inline-block"></span><span>Grupo</span></div></div>\n'
        '            </div>\n'
        '            <div class="chart-wrap"><canvas id="dutyRatioChart"></canvas></div>\n'
        '          </div>\n'
        '        </div>\n'
        '        <div class="card"><div class="card-head"><div class="card-title">Detalle Duty Hours por Per\u00edodo</div><div class="card-sub">Efectuado vs. programado</div></div><div id="dutyDetailTable" style="overflow-x:auto"></div></div>\n'
        '      </div>\n'

        # ── VIEW: DAN 121 ──
        '      <div class="view-section" id="view-dan">\n'
        '        <div style="font-size:12px;color:var(--muted);padding:4px 2px 8px;font-family:var(--mono)">Mes de referencia para l\u00edmites mensuales: <b id="danMonthLabel" style="color:var(--text2)"></b> \u00b7 Cambia con el selector de mes en el panel lateral.</div>\n'
        '        <div class="dan-grid" id="danCards"></div>\n'
        '        <div class="card">\n'
        '          <div class="card-head"><div class="card-title">Historial DAN 121 por Per\u00edodo</div><div class="card-sub">Block Hours ef. \u00b7 Duty Hours \u00b7 D\u00edas libres</div></div>\n'
        '          <div id="danHistory" style="overflow-x:auto"></div>\n'
        '          <div class="excl-note" id="danExclNote" style="display:none;margin-top:10px"><svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg><span id="danExclText"></span></div>\n'
        '        </div>\n'
        '        <div style="padding:10px 12px;background:var(--s2);border-radius:8px;font-size:10px;color:var(--muted);line-height:1.6;font-family:var(--mono)">Alertas indicativas. El c\u00e1lculo oficial de FDP y l\u00edmites es responsabilidad de Operaciones. Referencia: DAN 121 Art. 121.485 y 121.500.</div>\n'
        '      </div>\n'

        '    </div>\n'  # end dashboard
        '  </div>\n'   # end content
        '</div>\n'     # end main
        '</div>\n'     # end shell
        '<script>\n' + JS + '</script>\n'
        '</body>\n'
        '</html>\n'
    )


def generate_html(records, periods):
    period_labels = {p: PERIOD_LABELS_MAP.get(p, p) for p in periods}
    DATA_JS    = json.dumps(records,       ensure_ascii=False, default=str)
    PERIODS_JS = json.dumps(periods,       ensure_ascii=False)
    LABELS_JS  = json.dumps(period_labels, ensure_ascii=False)

    # Get CSS
    css_fn = generate_html.__globals__.get('_get_css')
    CSS = _get_css()

    # Build all JS parts
    JS_base  = build_js(DATA_JS, PERIODS_JS, LABELS_JS)
    JS_p2    = build_js_part2()
    JS_p3    = build_js_part3()
    JS_p4    = build_js_part4()
    JS_p5    = build_js_part5()

    # Alert list for resumen (reuse DAN logic inline)
    JS_alerts = (
        '\n// ── RESUMEN ALERTS (sidebar) ───────────────────────\n'
        'function renderResumenAlerts(pilotName, group, selectedPeriod) {\n'
        '  const pr  = RAW.filter(r => r.name === pilotName);\n'
        '  const sel = pr.find(r => r.period === selectedPeriod) || pr[0];\n'
        '  const mb  = sel ? (sel.block_h_actual || sel.block_h_programmed || 0) : 0;\n'
        '  const md  = sel ? (sel.duty_h_actual  || sel.duty_h_programmed  || 0) : 0;\n'
        '  const ml  = sel ? (sel.libre_days || 0) : 0;\n'
        '  const actP = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);\n'
        '  const accB = actP.reduce((s,r) => s+(r.block_h_actual||0), 0);\n'
        '  function alrt(t,title,desc){ return \'<div class="alert \'+t+\'"><div class="alert-dot"></div><div><div class="alert-title">\'+title+\'</div><div class="alert-desc">\'+desc+\'</div></div></div>\'; }\n'
        '  let alerts = "";\n'
        '  alerts += alrt(mb>100?"danger":mb>85?"warn":"ok","Block Hours \u00b7 "+fmt(mb)+"h",mb>100?"Supera l\u00edmite DAN 121 (100h/mes)":mb>85?"Cercano al l\u00edmite mensual":"Dentro del l\u00edmite (100h/mes)");\n'
        '  alerts += alrt(accB>900?"danger":accB>750?"warn":"ok","Acumulado YTD \u00b7 "+accB.toFixed(0)+"h",accB>900?"Muy cerca del l\u00edmite anual de 1.000h":accB>750?"Supera el 75% del l\u00edmite anual":"Sin riesgo l\u00edmite anual");\n'
        '  alerts += alrt(ml<8?"danger":ml<10?"warn":"ok","D\u00edas libres \u00b7 "+ml+"d",ml<8?"Bajo el m\u00ednimo reglamentario (8d/mes)":ml<10?"Dentro del m\u00ednimo, bajo el promedio":"Descanso adecuado seg\u00fan DAN 121");\n'
        '  alerts += alrt(md>130?"danger":md>105?"warn":"ok","Duty Hours \u00b7 "+fmt(md)+"h",md>130?"Duty hours muy elevadas, revisar FDPs":md>105?"Sobre el promedio del cargo":"Dentro de rango normal");\n'
        '  alerts += \'<div style="margin-top:6px;padding:9px 11px;background:var(--s2);border-radius:7px;font-size:10px;color:var(--muted);line-height:1.5;font-family:var(--mono)">Indicativo solamente. C\u00e1lculo oficial es responsabilidad de Operaciones.</div>\';\n'
        '  document.getElementById("alertListResumen").innerHTML = alerts;\n'
        '}\n'
        '\n// Patch selMonth listener to also update resumen alerts\n'
        'selMonth.addEventListener("change", () => {\n'
        '  if (selMonth.value && selPilot.value) {\n'
        '    renderResumenAlerts(selPilot.value, selGroup.value, selMonth.value);\n'
        '    document.getElementById("danMonthLabel").textContent = PERIOD_LABELS[selMonth.value] || selMonth.value;\n'
        '  }\n'
        '});\n'
        '\n// Patch render to also call resumen alerts and block detail table\n'
        'const _origRender = render;\n'
        'render = function(pilotName, group) {\n'
        '  _origRender(pilotName, group);\n'
        '  const pr = RAW.filter(r => r.name === pilotName);\n'
        '  const gr = RAW.filter(r => r.pos_group === group);\n'
        '  renderBlockView(pilotName, group, pr, gr);\n'
        '  // Block detail table (same as comp table)\n'
        '  let tbl = \'<table class="comp-table"><thead><tr><th>Per\u00edodo</th><th>Block prog.</th><th>Block ef.</th><th>\u0394 Block</th><th>Turnos</th><th>Vuelos prog.</th><th>Vuelos ef.</th><th>Blancos</th></tr></thead><tbody>\';\n'
        '  PERIODS.forEach(p => {\n'
        '    const r = pr.find(x => x.period===p); if(!r) return;\n'
        '    const pg=r.block_h_programmed||0, ac=r.block_h_actual||0, d=ac-pg;\n'
        '    const ex = r.exclude_from_avg ? \'<span style="color:var(--warn);font-size:9px"> \u2731</span>\' : "";\n'
        '    const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\u2014";\n'
        '    const dclr = d >= 0 ? "var(--teal)" : "var(--danger)";\n'
        '    const bl  = r.dias_blancos;\n'
        '    const blCell = bl !== null ? (bl > 5 ? \'<span style="color:var(--danger)">\'+bl+\'</span>\' : bl) : "\u2014";\n'
        '    tbl += \'<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">\' + (PERIOD_LABELS[p]||p) + ex + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (pg>0?pg.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ac>0?ac.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + dclr + \'">\' + dstr + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (r.turnos_programados!==null?r.turnos_programados:"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (r.vuelos_programados!==null?r.vuelos_programados:"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (r.vuelos_efectuados!==null?r.vuelos_efectuados:"\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + blCell + \'</td></tr>\';\n'
        '  });\n'
        '  tbl += "</tbody></table>";\n'
        '  document.getElementById("blockDetailTable").innerHTML = tbl;\n'
        '  // Duty detail table\n'
        '  let dtbl = \'<table class="comp-table"><thead><tr><th>Per\u00edodo</th><th>Duty prog.</th><th>Duty ef.</th><th>\u0394 Duty</th><th>Block ef.</th><th>Ratio D/B</th></tr></thead><tbody>\';\n'
        '  PERIODS.forEach(p => {\n'
        '    const r = pr.find(x => x.period===p); if(!r) return;\n'
        '    const dp=r.duty_h_programmed||0, da=r.duty_h_actual||0, dd=da-dp;\n'
        '    const bef=r.block_h_actual||0;\n'
        '    const ratio = bef>0&&da>0 ? (da/bef).toFixed(2) : "\u2014";\n'
        '    const dstr = dp > 0 ? ((dd>=0?"+":"")+dd.toFixed(1)+"h") : "\u2014";\n'
        '    const ex = r.exclude_from_avg ? \'<span style="color:var(--warn);font-size:9px"> \u2731</span>\' : "";\n'
        '    dtbl += \'<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">\' + (PERIOD_LABELS[p]||p) + ex + \'</td>\'\n'
        '          + \'<td style="font-family:var(--mono)">\' + (dp>0?dp.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '          + \'<td style="font-family:var(--mono)">\' + (da>0?da.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '          + \'<td style="font-family:var(--mono);color:\' + (dd>=0?"var(--teal)":"var(--danger)") + \'">\' + dstr + \'</td>\'\n'
        '          + \'<td style="font-family:var(--mono)">\' + (bef>0?bef.toFixed(1)+"h":"\u2014") + \'</td>\'\n'
        '          + \'<td style="font-family:var(--mono)">\' + ratio + \'</td></tr>\';\n'
        '  });\n'
        '  dtbl += "</tbody></table>";\n'
        '  document.getElementById("dutyDetailTable").innerHTML = dtbl;\n'
        '  // Initial alerts for resumen\n'
        '  const defaultPeriod = selMonth.value;\n'
        '  if(defaultPeriod) renderResumenAlerts(pilotName, group, defaultPeriod);\n'
        '  document.getElementById("danMonthLabel").textContent = PERIOD_LABELS[selMonth.value] || selMonth.value;\n'
        '};\n'
    )

    JS = JS_base + JS_p2 + JS_p3 + JS_p4 + JS_p5 + JS_alerts

    # Use real logo from original parser if embedded, else placeholder
    return build_html(CSS, JS)


def _get_css():
    # Import the CSS build from generate_html's local function
    # We call it via the module-level CSS builder
    return _build_css_string()

def _build_css_string():
    return (
        ':root{'
        '--purple:#671E77;--purple-l:#9B44B8;--purple-xl:#C480E0;'
        '--purple-dim:rgba(103,30,119,0.18);--purple-dim2:rgba(103,30,119,0.08);'
        '--green:#26D800;--green-l:#5CF200;--green-dim:rgba(38,216,0,0.15);--green-dim2:rgba(38,216,0,0.07);'
        '--violet:#8B35A8;--teal:#00C89B;--teal-dim:rgba(0,200,155,0.12);'
        '--danger:#E53E3E;--danger-dim:rgba(229,62,62,0.12);'
        '--warn:#C46AE0;--warn-dim:rgba(196,106,224,0.12);'
        '--bg:#F8F7FC;--surface:#FFFFFF;--s2:#F0EBF7;--s3:#E8DFF5;'
        '--border:rgba(103,30,119,0.18);--border2:rgba(103,30,119,0.35);'
        '--text:#2A1240;--text2:#5A3878;--muted:#8B6FA8;--dim:#B09CC8;'
        '--r:10px;--r2:14px;'
        '--shadow:0 1px 4px rgba(0,0,0,.08),0 4px 20px rgba(103,30,119,.10);'
        '--shadow2:0 2px 12px rgba(0,0,0,.12),0 8px 32px rgba(103,30,119,.18);'
        "--font:'DM Sans',sans-serif;--display:'Playfair Display',serif;--mono:'DM Mono',monospace;"
        '}'
        '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}'
        'html{font-size:14px}'
        'body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.5}'
        '.shell{display:grid;grid-template-columns:230px 1fr;min-height:100vh}'
        '.sidebar{background:#FFFFFF;display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;border-right:1px solid var(--border)}'
        '.sidebar-top{padding:0 0 14px;border-bottom:1px solid rgba(103,30,119,.15)}'
        '.logo-wrap{width:100%;background:#FFFFFF;display:flex;align-items:center;justify-content:center;padding:14px 18px}'
        '.logo-wrap img{width:100%;max-width:192px;height:auto;display:block}'
        '.brand-sub-line{font-size:9px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase;text-align:center;padding:5px 0 0;font-family:var(--mono)}'
        '.filters{padding:14px 16px;display:flex;flex-direction:column;gap:11px;border-bottom:1px solid rgba(103,30,119,.15)}'
        '.f-block{display:flex;flex-direction:column;gap:5px}'
        '.f-label{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono)}'
        '.f-select{appearance:none;background:rgba(103,30,119,.05);border:1px solid rgba(103,30,119,.25);border-radius:8px;color:var(--text);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;background-image:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'12\' height=\'12\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%238B6FA8\' stroke-width=\'2\'%3E%3Cpolyline points=\'6 9 12 15 18 9\'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 9px center}'
        '.f-select:focus,.f-select:hover{border-color:var(--green);box-shadow:0 0 0 2px rgba(38,216,0,.15)}'
        '.f-select option{background:#FFFFFF;color:var(--text)}'
        '.sidebar-nav{padding:10px 8px;flex:1}'
        '.nav-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:var(--text2);cursor:pointer;transition:all .15s;margin-bottom:2px;border-left:2px solid transparent;user-select:none}'
        '.nav-item:hover{color:var(--purple);background:rgba(103,30,119,.08);border-left-color:rgba(103,30,119,.4)}'
        '.nav-item.active{color:var(--purple);background:rgba(103,30,119,.12);border-left-color:var(--purple);font-weight:500}'
        '.nav-item svg{width:14px;height:14px;flex-shrink:0}'
        '.sidebar-footer{padding:12px 16px;border-top:1px solid rgba(103,30,119,.15)}'
        '.pilot-badge{display:flex;align-items:center;gap:10px}'
        '.pilot-avatar{width:34px;height:34px;border-radius:50%;background:var(--purple);border:1.5px solid var(--purple-l);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}'
        '.pilot-name-s{font-size:11px;font-weight:500;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
        '.pilot-pos-s{font-size:10px;color:var(--muted);font-family:var(--mono)}'
        '.main{display:flex;flex-direction:column;min-height:100vh}'
        '.topbar{background:#FFFFFF;border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;box-shadow:0 1px 8px rgba(103,30,119,.06)}'
        '.page-title{font-family:var(--display);font-size:17px;color:var(--text)}'
        '.page-title span{color:var(--purple)}'
        '.page-sub{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.topbar-right{display:flex;align-items:center;gap:8px}'
        '.pill{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--s2);color:var(--text2)}'
        '.dot{width:6px;height:6px;border-radius:50%;background:var(--green);box-shadow:0 0 7px var(--green)}'
        '.content{padding:18px 26px;display:flex;flex-direction:column;gap:13px;flex:1}'
        '.view-section{display:none;flex-direction:column;gap:16px}'
        '.view-section.active{display:flex}'
        '.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}'
        '.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 15px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:box-shadow .2s,transform .15s,border-color .2s}'
        '.kpi:hover{box-shadow:var(--shadow2);transform:translateY(-1px);border-color:var(--border2)}'
        ".kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;border-radius:var(--r2) var(--r2) 0 0}"
        '.kpi.k-p1::before{background:var(--purple-l)}.kpi.k-p2::before{background:var(--violet)}'
        '.kpi.k-g1::before{background:var(--green)}.kpi.k-g2::before{background:var(--teal)}'
        '.kpi.k-g3::before{background:var(--green-l)}.kpi.k-r1::before{background:var(--danger)}'
        '.kpi-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;font-family:var(--mono)}'
        '.kpi-val{font-size:24px;font-weight:600;color:var(--text);font-family:var(--mono);letter-spacing:-.03em;line-height:1}'
        '.kpi-unit{font-size:12px;font-weight:400;color:var(--muted);margin-left:2px}'
        '.kpi-footer{display:flex;align-items:center;justify-content:space-between;margin-top:7px}'
        '.kpi-vs{font-size:10px;color:var(--muted)}.kpi-vs b{color:var(--text2);font-weight:500}'
        '.delta{font-size:10px;font-family:var(--mono);padding:2px 6px;border-radius:4px}'
        '.d-up{background:rgba(38,216,0,.15);color:#1A7A00}'
        '.d-down{background:rgba(229,62,62,.12);color:#C0392B}'
        '.d-neu{background:var(--s3);color:var(--muted)}'
        '.d-warn{background:rgba(196,106,224,.15);color:#6B1A8A}'
        '.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow)}'
        '.card-head{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}'
        '.card-title{font-size:13px;font-weight:500;color:var(--text)}'
        '.card-sub{font-size:10px;color:var(--muted);margin-top:2px;font-family:var(--mono)}'
        '.legend{display:flex;gap:12px;align-items:center;font-size:10px;color:var(--muted);font-family:var(--mono);flex-wrap:wrap}'
        '.leg{display:flex;align-items:center;gap:5px}'
        '.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}'
        '.chart-wrap{position:relative;height:220px}'
        '.chart-wrap-lg{position:relative;height:300px}'
        '.comp-table{width:100%;border-collapse:collapse;font-size:12px}'
        '.comp-table th{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--text2);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--s2)}'
        '.comp-table td{padding:8px 10px;border-bottom:1px solid rgba(103,30,119,.12)}'
        '.comp-table tr:last-child td{border-bottom:none}'
        '.comp-table tr:hover td{background:var(--s2)}'
        '.bottom-row{display:grid;grid-template-columns:1fr 300px;gap:14px}'
        '.prog-list{display:flex;flex-direction:column;gap:13px}'
        '.prog-head{display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px}'
        '.prog-lbl{color:var(--text2)}.prog-num{font-family:var(--mono);font-size:11px}'
        '.prog-track{height:5px;background:var(--s3);border-radius:3px;overflow:hidden}'
        '.prog-fill{height:100%;border-radius:3px}'
        '.prog-note{font-size:9px;color:var(--dim);margin-top:2px;font-family:var(--mono)}'
        '.alert-list{display:flex;flex-direction:column;gap:7px}'
        '.alert{display:flex;align-items:flex-start;gap:9px;padding:9px 11px;border-radius:8px;border:1px solid}'
        '.alert.ok{background:rgba(38,216,0,.07);border-color:rgba(38,216,0,.3)}'
        '.alert.warn{background:rgba(139,53,168,.08);border-color:rgba(139,53,168,.3)}'
        '.alert.danger{background:rgba(229,62,62,.08);border-color:rgba(229,62,62,.35)}'
        '.alert-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px}'
        '.alert.ok .alert-dot{background:var(--green)}'
        '.alert.warn .alert-dot{background:#8B22AA}'
        '.alert.danger .alert-dot{background:var(--danger)}'
        '.alert-title{font-size:11px;font-weight:500;color:var(--text)}'
        '.alert-desc{font-size:10px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.excl-note{display:flex;align-items:center;gap:6px;padding:8px 11px;border-radius:7px;background:var(--s2);border:1px solid var(--border);font-size:10px;color:var(--muted)}'
        '.excl-note svg{width:12px;height:12px;flex-shrink:0}'
        '.dan-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}'
        '.dan-card{border-radius:var(--r2);padding:20px;border:2px solid;position:relative}'
        '.dan-card.ok{background:rgba(38,216,0,.06);border-color:rgba(38,216,0,.35)}'
        '.dan-card.warn{background:rgba(139,53,168,.08);border-color:rgba(139,53,168,.35)}'
        '.dan-card.danger{background:rgba(229,62,62,.08);border-color:rgba(229,62,62,.40)}'
        '.dan-label{font-size:10px;font-family:var(--mono);text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:6px}'
        '.dan-val{font-size:28px;font-weight:700;font-family:var(--mono);letter-spacing:-.04em;line-height:1;margin-bottom:4px}'
        '.dan-card.ok .dan-val{color:#1A7A00}.dan-card.warn .dan-val{color:#6B1A8A}.dan-card.danger .dan-val{color:#C0392B}'
        '.dan-limit{font-size:11px;color:var(--muted);font-family:var(--mono)}'
        '.dan-bar-wrap{margin-top:10px;height:6px;background:rgba(0,0,0,.08);border-radius:3px;overflow:hidden}'
        '.dan-bar-fill{height:100%;border-radius:3px}'
        '.dan-card.ok .dan-bar-fill{background:var(--green)}'
        '.dan-card.warn .dan-bar-fill{background:#8B22AA}'
        '.dan-card.danger .dan-bar-fill{background:var(--danger)}'
        '::-webkit-scrollbar{width:4px}::-webkit-scrollbar-thumb{background:rgba(103,30,119,.3);border-radius:2px}'
        '@keyframes fadeUp{from{opacity:0;transform:translateY(7px)}to{opacity:1;transform:none}}'
        '.kpi,.card,.dan-card{animation:fadeUp .28s ease both}'
        '.hamburger{display:none;position:fixed;top:12px;left:12px;z-index:200;width:38px;height:38px;border-radius:8px;border:1px solid var(--border2);background:var(--s2);cursor:pointer;align-items:center;justify-content:center;flex-direction:column;gap:5px;padding:9px}'
        '.hamburger span{display:block;width:16px;height:1.5px;background:var(--text2);border-radius:2px;transition:all .2s}'
        '.hamburger.open span:nth-child(1){transform:translateY(6.5px) rotate(45deg)}'
        '.hamburger.open span:nth-child(2){opacity:0}'
        '.hamburger.open span:nth-child(3){transform:translateY(-6.5px) rotate(-45deg)}'
        '.sidebar-overlay{display:none;position:fixed;inset:0;background:rgba(40,20,60,.5);z-index:155;backdrop-filter:blur(2px)}'
        '.sidebar-overlay.open{display:block}'
        '@media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(3,1fr)}.dan-grid{grid-template-columns:repeat(2,1fr)}}'
        '@media(max-width:1024px){.charts-row{grid-template-columns:1fr}.bottom-row{grid-template-columns:1fr}}'
        '@media(max-width:768px){.shell{grid-template-columns:1fr}.sidebar{position:fixed;left:-240px;top:0;height:100vh;width:230px;z-index:160;transition:left .25s cubic-bezier(.4,0,.2,1)}.sidebar.open{left:0;box-shadow:6px 0 32px rgba(46,36,22,.3)}.hamburger{display:flex}.topbar{padding:12px 16px 12px 58px}.content{padding:14px 16px}.kpi-grid{grid-template-columns:repeat(2,1fr);gap:8px}.charts-row{grid-template-columns:1fr;gap:10px}.bottom-row{grid-template-columns:1fr;gap:10px}.chart-wrap{height:190px}.chart-wrap-lg{height:240px}.page-title{font-size:14px}.card{padding:14px}.card-head{flex-direction:column;gap:8px}.legend{gap:8px;font-size:9px}.dan-grid{grid-template-columns:1fr}}'
        '@media(max-width:420px){.kpi-grid{grid-template-columns:1fr 1fr}.kpi-val{font-size:21px}.kpi{padding:12px 12px}.chart-wrap{height:165px}.chart-wrap-lg{height:200px}.content{padding:10px 12px}}'
    )


if __name__ == '__main__':
    records, periods = build_dataset()
    html = generate_html(records, periods)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print('\nDashboard generado: ' + str(OUTPUT_HTML))
    print('Tamaño: ' + str(len(html)//1024) + ' KB')
