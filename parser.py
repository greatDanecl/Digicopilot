"""
SDC Parser — lee todos los xlsx en /data y genera index.html.
Uso: python parser.py
"""
import pandas as pd, re, json, numpy as np, os, glob, sys
from datetime import timedelta
from collections import defaultdict, Counter
from pathlib import Path

# ── CONFIGURACIÓN ─────────────────────────────────────────
DATA_DIR    = Path(__file__).parent / "data"
OUTPUT_HTML = Path(__file__).parent / "index.html"

MONTH_MAP = {
    # Inglés mayúscula/minúscula
    'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
    'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12',
    'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
    'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12',
    # Español mayúscula/minúscula — todos los meses
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

# ── HELPERS ───────────────────────────────────────────────
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
    if not clean:
        return 'BLANCO'
    joined = ' '.join(clean).upper()
    if re.search(r'\bTURNO\d+', joined):
        return 'TURNO'
    flights = [v for v in clean if re.match(r'^\d{2,4}$', v.strip())]
    if flights:
        return 'VUELO'
    for code, label in [
        ('HOTEL',  'HOTEL'), ('SIM',    'SIM'), ('ELEAR',  'ELEAR'),
        ('DH',     'DH'),    ('ACT',    'ACT'), ('OFNA2',  'OFNA2'),
        ('CEMAE',  'CEMAE'), ('LQUIN',  'LQUIN'),('SINDI',  'SINDI'),
        ('FVUEL',  'FVUEL'), ('VACAC',  'VACAC'),('BDAY',   'BDAY'),
        ('LIBRE',  'LIBRE'), ('FINDE',  'FINDE'),
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
    counts = {
        'turnos': 0, 'vuelos': 0, 'blancos': 0,
        'hotel': 0, 'sim': 0, 'elear': 0, 'act': 0, 'dh': 0
    }
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

# ── PROCESO PRINCIPAL ──────────────────────────────────────
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
                if cfg_entry: period = cfg_entry['period']
                else: period = detect_period_from_filename(fname) or detect_period_from_df(df)

                if not period:
                    print('  ? ' + fname + '/' + sheet_name + ': periodo no detectado, omitiendo')
                    continue

                if cfg_entry:
                    sl = sheet_name.lower()
                    if any(w in sl for w in ['hora','actual','efect']): role = 'actual'
                    else: role = cfg_entry['role']
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
                'vuelos_efectuados':  None, 'dias_blancos':        None,
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

    names_by_grp = defaultdict(set)
    for r in records:
        names_by_grp[r['pos_group']].add(r['name'])

    print('\n' + '-'*50)
    print('Total registros: ' + str(len(records)))
    print('Periodos: ' + str([PERIOD_LABELS_MAP.get(p,p) for p in periods]))
    for g, ns in sorted(names_by_grp.items()):
        print('  ' + g + ': ' + str(len(ns)) + ' pilotos')

    return records, periods

# ── GENERAR HTML ───────────────────────────────────────────

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SDC \u00b7 Productividad de Tripulaci\u00f3n</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
:root{
  --purple:#671E77;--purple-l:#9B44B8;--purple-xl:#C480E0;
  --purple-dim:rgba(103,30,119,0.18);--purple-dim2:rgba(103,30,119,0.08);
  --green:#26D800;--green-l:#5CF200;--green-dim:rgba(38,216,0,0.15);--green-dim2:rgba(38,216,0,0.07);
  --violet:#8B35A8;--teal:#00C89B;--teal-dim:rgba(0,200,155,0.12);
  --danger:#FF4466;--danger-dim:rgba(255,68,102,0.12);
  --warn:#C46AE0;--warn-dim:rgba(196,106,224,0.12);
  --bg:#F8F7FC;--surface:#FFFFFF;--s2:#F0EBF7;--s3:#E8DFF5;
  --border:rgba(103,30,119,0.18);--border2:rgba(103,30,119,0.35);
  --text:#2A1240;--text2:#5A3878;--muted:#8B6FA8;--dim:#B09CC8;
  --r:10px;--r2:14px;
  --shadow:0 1px 4px rgba(0,0,0,.4),0 4px 20px rgba(103,30,119,.15);
  --shadow2:0 2px 12px rgba(0,0,0,.5),0 8px 32px rgba(103,30,119,.25);
  --font:'DM Sans',sans-serif;--display:'Playfair Display',serif;--mono:'DM Mono',monospace;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{font-size:14px}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.5}
.shell{display:grid;grid-template-columns:230px 1fr;min-height:100vh}
.sidebar{background:#FFFFFF;display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;border-right:1px solid var(--border)}
.sidebar-top{padding:0 0 14px;border-bottom:1px solid rgba(103,30,119,.15)}
.logo-wrap{width:100%;background:#FFFFFF;display:flex;align-items:center;justify-content:center;padding:24px 18px 10px}
.brand-sub-line{font-size:9px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase;text-align:center;padding:5px 0 0;font-family:var(--mono)}
.filters{padding:14px 16px;display:flex;flex-direction:column;gap:11px;border-bottom:1px solid rgba(103,30,119,.15)}
.f-block{display:flex;flex-direction:column;gap:5px}
.f-label{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono)}
.f-select{appearance:none;background:rgba(103,30,119,.05);border:1px solid rgba(103,30,119,.25);border-radius:8px;color:var(--text);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238B6FA8' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 9px center}
.f-select:focus,.f-select:hover{border-color:var(--green);box-shadow:0 0 0 2px rgba(38,216,0,.15)}
.f-select option{background:#FFFFFF;color:var(--text)}
.sidebar-nav{padding:10px 8px;flex:1}
.nav-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:var(--text2);cursor:pointer;transition:all .15s;margin-bottom:2px;border-left:2px solid transparent}
.nav-item:hover{color:var(--purple);background:rgba(103,30,119,.08);border-left-color:rgba(103,30,119,.4)}
.nav-item.active{color:var(--purple);background:rgba(103,30,119,.12);border-left-color:var(--purple);font-weight:500;}
.nav-item svg{width:14px;height:14px;flex-shrink:0}
.sidebar-footer{padding:12px 16px;border-top:1px solid rgba(103,30,119,.15)}
.pilot-badge{display:flex;align-items:center;gap:10px}
.pilot-avatar{width:34px;height:34px;border-radius:50%;background:var(--purple);border:1.5px solid var(--purple-l);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}
.pilot-name-s{font-size:11px;font-weight:500;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pilot-pos-s{font-size:10px;color:var(--muted);font-family:var(--mono)}
.main{display:flex;flex-direction:column;min-height:100vh}
.topbar{background:#FFFFFF;border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;box-shadow:0 1px 8px rgba(103,30,119,.08)}
.page-title{font-family:var(--display);font-size:17px;color:var(--text)}
.page-title span{color:var(--purple)}
.page-sub{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}
.topbar-right{display:flex;align-items:center;gap:8px}
.pill{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--s2);color:var(--text2)}
.dot{width:6px;height:6px;border-radius:50%;background:var(--green);box-shadow:0 0 7px var(--green)}
.content{padding:18px 26px;display:flex;flex-direction:column;gap:13px;flex:1}
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 15px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:box-shadow .2s,transform .15s,border-color .2s}
.kpi:hover{box-shadow:var(--shadow2);transform:translateY(-1px);border-color:var(--border2)}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;border-radius:var(--r2) var(--r2) 0 0}
.kpi.k-p1::before{background:var(--purple-l)}
.kpi.k-p2::before{background:var(--violet)}
.kpi.k-g1::before{background:var(--green)}
.kpi.k-g2::before{background:var(--teal)}
.kpi.k-g3::before{background:var(--green-l)}
.kpi-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;font-family:var(--mono)}
.kpi-val{font-size:24px;font-weight:600;color:var(--text);font-family:var(--mono);letter-spacing:-.03em;line-height:1}
.kpi-unit{font-size:12px;font-weight:400;color:var(--muted);margin-left:2px}
.kpi-footer{display:flex;align-items:center;justify-content:space-between;margin-top:7px}
.kpi-vs{font-size:10px;color:var(--muted)}.kpi-vs b{color:var(--text2);font-weight:500}
.delta{font-size:10px;font-family:var(--mono);padding:2px 6px;border-radius:4px}
.d-up{background:var(--green-dim);color:var(--green-l)}
.d-down{background:var(--danger-dim);color:var(--danger)}
.d-neu{background:var(--s3);color:var(--muted)}
.d-warn{background:var(--warn-dim);color:var(--warn)}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow); display:flex; flex-direction:column; justify-content:space-between;}
.card-head{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}
.card-title{font-size:13px;font-weight:500;color:var(--text)}
.card-sub{font-size:10px;color:var(--muted);margin-top:2px;font-family:var(--mono)}
.legend{display:flex;gap:12px;align-items:center;font-size:10px;color:var(--muted);font-family:var(--mono);flex-wrap:wrap}
.leg{display:flex;align-items:center;gap:5px}
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.chart-wrap{position:relative;height:220px}
.comp-table{width:100%;border-collapse:collapse;font-size:12px}
.comp-table th{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--s2);color:var(--text2)}
.comp-table td{padding:8px 10px;border-bottom:1px solid rgba(103,30,119,.2)}
.comp-table tr:last-child td{border-bottom:none}
.comp-table tr:hover td{background:var(--s2)}
.bottom-row{display:grid;grid-template-columns:1fr 300px;gap:14px}
.prog-list{display:flex;flex-direction:column;gap:13px}
.prog-head{display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px}
.prog-lbl{color:var(--text2)}.prog-num{font-family:var(--mono);font-size:11px}
.prog-track{height:5px;background:var(--s3);border-radius:3px;overflow:hidden}
.prog-fill{height:100%;border-radius:3px}
.prog-note{font-size:9px;color:var(--dim);margin-top:2px;font-family:var(--mono)}
.alert-list{display:flex;flex-direction:column;gap:7px}
.alert{display:flex;align-items:flex-start;gap:9px;padding:9px 11px;border-radius:8px;border:1px solid}
.alert.ok{background:var(--green-dim2);border-color:rgba(38,216,0,.25)}
.alert.warn{background:var(--warn-dim);border-color:rgba(196,106,224,.3)}
.alert.danger{background:var(--danger-dim);border-color:rgba(255,68,102,.3)}
.alert-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px}
.alert.ok .alert-dot{background:var(--green)}
.alert.warn .alert-dot{background:var(--warn)}
.alert.danger .alert-dot{background:var(--danger)}
.alert-title{font-size:11px;font-weight:500;color:var(--text)}
.alert-desc{font-size:10px;color:var(--muted);margin-top:1px;font-family:var(--mono)}
.excl-note{display:flex;align-items:center;gap:6px;padding:8px 11px;border-radius:7px;background:var(--s2);border:1px solid var(--border);font-size:10px;color:var(--muted);margin-top:10px}
.excl-note svg{width:12px;height:12px;flex-shrink:0}
::-webkit-scrollbar{width:4px}
::-webkit-scrollbar-thumb{background:rgba(103,30,119,.3);border-radius:2px}
@keyframes fadeUp{from{opacity:0;transform:translateY(7px)}to{opacity:1;transform:none}}
.kpi,.card{animation:fadeUp .28s ease both}
.turnos-bar{display:flex;align-items:center;gap:8px;margin-top:6px}
.tbar-track{flex:1;height:7px;background:var(--s3);border-radius:4px;overflow:hidden;position:relative}
.tbar-prog{height:100%;background:var(--purple-l);border-radius:4px;position:absolute;left:0;top:0}
.tbar-act{height:100%;background:var(--green);border-radius:4px;position:absolute;left:0;top:0;opacity:.85}
.tbar-label{font-size:10px;font-family:var(--mono);color:var(--muted);white-space:nowrap}
.hamburger{display:none;position:fixed;top:12px;left:12px;z-index:200;width:38px;height:38px;border-radius:8px;border:1px solid var(--border2);background:var(--s2);cursor:pointer;align-items:center;justify-content:center;flex-direction:column;gap:5px;padding:9px}
.hamburger span{display:block;width:16px;height:1.5px;background:var(--text2);border-radius:2px;transition:all .2s}
.hamburger.open span:nth-child(1){transform:translateY(6.5px) rotate(45deg)}
.hamburger.open span:nth-child(2){opacity:0}
.hamburger.open span:nth-child(3){transform:translateY(-6.5px) rotate(-45deg)}
.sidebar-overlay{display:none;position:fixed;inset:0;background:rgba(40,20,60,.5);z-index:155;backdrop-filter:blur(2px)}
.sidebar-overlay.open{display:block}
@media(max-width:1024px){
.kpi-grid{grid-template-columns:repeat(3,1fr)}
.charts-row{grid-template-columns:1fr}
.bottom-row{grid-template-columns:1fr}
}
@media(max-width:768px){
.shell{grid-template-columns:1fr}
.sidebar{position:fixed;left:-240px;top:0;height:100vh;width:230px;z-index:160;transition:left .25s cubic-bezier(.4,0,.2,1)}
.sidebar.open{left:0;box-shadow:6px 0 32px rgba(46,36,22,.3)}
.hamburger{display:flex}
.topbar{padding:12px 16px 12px 58px}
.content{padding:14px 16px}
.kpi-grid{grid-template-columns:repeat(2,1fr);gap:8px}
.charts-row{grid-template-columns:1fr;gap:10px}
.bottom-row{grid-template-columns:1fr;gap:10px}
.chart-wrap{height:190px !important}
.page-title{font-size:14px}
.page-sub{font-size:10px;margin-top:0}
.card{padding:14px}
.card-head{flex-direction:column;gap:8px;align-items:flex-start}
.legend{gap:8px;font-size:9px}
#compTableWrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
.comp-table th,.comp-table td{padding:6px 8px}
}
@media(max-width:420px){
.kpi-grid{grid-template-columns:1fr 1fr}
.kpi-val{font-size:21px}
.kpi{padding:12px 12px}
.chart-wrap{height:165px !important}
.content{padding:10px 12px}
}
</style>
</head>
<body>
<div class="shell">
<button class="hamburger" id="menuBtn" aria-label="Abrir menú"><span></span><span></span><span></span></button>
<div class="sidebar-overlay" id="overlay"></div>
<div class="sidebar" id="sidebar">
  <div class="sidebar-top">
    <div class="logo-wrap">
      <svg viewBox="0 0 500 120" xmlns="http://www.w3.org/2000/svg" style="width: 100%; height: auto; display: block;">
        <!-- Ala Izquierda -->
        <polygon points="10,15 130,15 120,40 25,40" fill="#671E77"/>
        <polygon points="30,48 120,48 110,73 45,73" fill="#671E77"/>
        <polygon points="50,81 110,81 100,106 65,106" fill="#5CF200"/>
        <!-- Ala Derecha -->
        <polygon points="370,15 490,15 475,40 380,40" fill="#671E77"/>
        <polygon points="380,48 470,48 455,73 390,73" fill="#671E77"/>
        <polygon points="390,81 450,81 435,106 400,106" fill="#5CF200"/>
        <!-- Textos -->
        <text x="250" y="38" font-family="'DM Sans', sans-serif" font-weight="bold" font-size="30" fill="#671E77" text-anchor="middle" letter-spacing="1">SINDICATO</text>
        <text x="250" y="74" font-family="'DM Sans', sans-serif" font-weight="bold" font-size="38" fill="#671E77" text-anchor="middle" letter-spacing="1.5">PILOTOS</text>
        <text x="250" y="108" font-family="'DM Sans', sans-serif" font-weight="bold" font-size="32" text-anchor="middle" letter-spacing="0.5"><tspan fill="#5CF200">SKY </tspan><tspan fill="#671E77">AIRLINE</tspan></text>
      </svg>
    </div>
    <div class="brand-sub-line">Digital Copilot</div>
  </div>
  <div class="filters">
    <div class="f-block"><div class="f-label">Cargo</div>
      <select class="f-select" id="selGroup">
        <option value="">— Seleccionar cargo —</option>
        <option value="Capitán">Capitán</option>
        <option value="Primer Oficial">Primer Oficial</option>
        <option value="Instructor">Instructor</option>
      </select>
    </div>
    <div class="f-block"><div class="f-label">Tripulante</div>
      <select class="f-select" id="selPilot" disabled>
        <option value="">— Seleccione cargo primero —</option>
      </select>
    </div>
    <div class="f-block"><div class="f-label">Mes (KPIs)</div>
      <select class="f-select" id="selMonth" disabled>
        <option value="">— Seleccione tripulante —</option>
      </select>
    </div>
  </div>
  <nav class="sidebar-nav">
    <div class="nav-item active" data-tab="resumen"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>Resumen</div>
    <div class="nav-item" data-tab="block"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>Block hours</div>
    <div class="nav-item" data-tab="duty"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>Duty hours</div>
    <div class="nav-item" data-tab="dan121"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>DAN 121</div>
  </nav>
  <div class="sidebar-footer">
    <div class="pilot-badge">
      <div class="pilot-avatar" id="sideAvatar">—</div>
      <div><div class="pilot-name-s" id="sideName">Sin selección</div><div class="pilot-pos-s" id="sidePos">—</div></div>
    </div>
  </div>
</div>
<div class="main">
  <div class="topbar">
    <div>
      <div class="page-title" id="pageTitle">Seleccione un <span>tripulante</span></div>
      <div class="page-sub" id="pageSub">SDC \u00b7 Productividad de Tripulación</div>
    </div>
    <div class="topbar-right">
      <div class="pill"><span class="dot"></span>Sistema activo</div>
      <div class="pill" id="periodPill">—</div>
    </div>
  </div>
  <div class="content">
    <div id="placeholder" style="display:flex;flex-direction:column;align-items:center;justify-content:center;flex:1;gap:14px;color:var(--dim);padding:60px 0;">
      <svg width="44" height="44" fill="none" stroke="currentColor" stroke-width="1.2" viewBox="0 0 24 24" style="stroke:var(--border2)"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/></svg>
      <div style="font-family:var(--display);font-size:18px;color:var(--text2)">SDC \u00b7 SPSKY Digital Copilot</div>
      <div style="font-size:12px;text-align:center;max-width:300px;line-height:1.7;color:var(--muted)">Seleccione un cargo y un tripulante para visualizar sus indicadores de productividad.</div>
      <div style="font-size:10px;font-family:var(--mono);color:var(--dim);margin-top:4px" id="periodsHint"></div>
    </div>
    <div id="dashboard" style="display:none;flex-direction:column;gap:16px;">
      <div class="kpi-grid" id="kpiRow"></div>
      <div id="tabContent" style="display:flex;flex-direction:column;gap:16px;">
        <!-- Se rellena con JS -->
      </div>
    </div>
  </div>
</div>
</div>
<script>
__JS_CONTENT__
</script>
</body>
</html>
"""

JS_TEMPLATE = r"""
const RAW = __RAW__;
const PERIODS = __PERIODS__;
const PERIOD_LABELS = __LABELS__;

document.getElementById('periodPill').textContent = Object.values(PERIOD_LABELS).join(' \u00b7 ');
document.getElementById('periodsHint').textContent = 'Per\u00edodos: ' + Object.values(PERIOD_LABELS).join(' \u00b7 ');

let blockChartInst = null, compareChartInst = null, dutyChartInst = null;
let currentTab = 'resumen';

const selGroup = document.getElementById("selGroup");
const selPilot = document.getElementById("selPilot");
const selMonth = document.getElementById("selMonth");

// ── UTILS ──
function fmt(v, d=1) { return (v == null || +v === 0) ? "\u2014" : (+v).toFixed(d); }
function avg(arr) { const v = arr.filter(x => x != null && x > 0); return v.length ? v.reduce((a,b) => a+b, 0)/v.length : 0; }
function dc(d) { return d > 2 ? "d-up" : d < -2 ? "d-down" : "d-neu"; }
function ds(d) { return (d >= 0 ? "+" : "") + d.toFixed(1) + "%"; }
function bestBlock(r) { return (r.block_h_actual && r.block_h_actual > 0) ? r.block_h_actual : (r.block_h_programmed || 0); }
function bestDuty(r) { return (r.duty_h_actual && r.duty_h_actual > 0) ? r.duty_h_actual : (r.duty_h_programmed || 0); }
function isProgOnly(r) { return !(r.block_h_actual && r.block_h_actual > 0) && (r.block_h_programmed && r.block_h_programmed > 0); }
function makeGrad(ctx, ca, c1, c2) {
  if (!ca) return "transparent";
  const g = ctx.createLinearGradient(0, ca.top, 0, ca.bottom);
  g.addColorStop(0, c1); g.addColorStop(1, c2); return g;
}

const tooltipStyle = {
    backgroundColor: '#FFFFFF',
    titleColor: '#2A1240',
    bodyColor: '#2A1240',
    borderColor: 'rgba(103,30,119,0.25)',
    borderWidth: 1,
    padding: 12,
    boxPadding: 4,
    usePointStyle: true,
    titleFont: {family: "'DM Sans',sans-serif", size: 13, weight: 600},
    bodyFont: {family: "'DM Mono',monospace", size: 12}
};

// ── EVENTS ──
selGroup.addEventListener("change", () => {
  const g = selGroup.value;
  const names = [...new Set(RAW.filter(r => r.pos_group === g).map(r => r.name))].sort((a,b) => a.localeCompare(b, "es"));
  selPilot.innerHTML = '<option value="">\u2014 Seleccionar tripulante \u2014</option>';
  names.forEach(n => { const o = document.createElement("option"); o.value = o.textContent = n; selPilot.appendChild(o); });
  selPilot.disabled = false;
  selMonth.innerHTML = '<option value="">\u2014 Seleccione un tripulante \u2014</option>';
  selMonth.disabled = true;
  document.getElementById("placeholder").style.display = "flex";
  document.getElementById("dashboard").style.display = "none";
});

selPilot.addEventListener("change", () => {
    if(!selPilot.value) return;
    const pr = RAW.filter(r => r.name === selPilot.value);
    const pilotPeriods = [...new Set(pr.map(r => r.period))].sort().reverse();
    selMonth.innerHTML = "";
    pilotPeriods.forEach(p => {
        const o = document.createElement("option"); o.value = p;
        const r = pr.find(x => x.period === p);
        const hasBoth = r && r.block_h_actual > 0 && r.block_h_programmed > 0;
        const hasAct  = r && r.block_h_actual > 0;
        o.textContent = (PERIOD_LABELS[p] || p) + (hasBoth ? " (prog+ef)" : hasAct ? " (ef)" : " (prog)");
        selMonth.appendChild(o);
    });
    const latest = pr.filter(r => r.block_h_actual > 0).sort((a,b) => b.period.localeCompare(a.period))[0] || pr.sort((a,b) => b.period.localeCompare(a.period))[0];
    selMonth.value = latest ? latest.period : pilotPeriods[0];
    selMonth.disabled = false;
    
    // Configurar header sidebar
    const init = selPilot.value.split(" ").filter((_,i) => i < 2).map(w => w[0]).join("");
    document.getElementById("sideAvatar").textContent = init;
    document.getElementById("sideName").textContent = selPilot.value.split(" ").slice(0,2).join(" ");
    document.getElementById("sidePos").textContent = (latest ? latest.pos : selGroup.value) + " \u00b7 " + (latest ? latest.base : "");
    document.getElementById("pageTitle").innerHTML = "<span>" + selPilot.value.split(" ").slice(0,2).join(" ") + "</span> \u00b7 Productividad";
    document.getElementById("pageSub").textContent = (latest ? latest.pos_group : selGroup.value) + " \u00b7 " + (latest ? latest.base : "") + " \u00b7 " + Object.values(PERIOD_LABELS).join(" \u00b7 ");

    updateDashboard();
});

selMonth.addEventListener("change", () => { if (selPilot.value) updateDashboard(); });

document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        e.currentTarget.classList.add('active');
        currentTab = e.currentTarget.getAttribute('data-tab');
        if (selPilot.value) updateDashboard();
    });
});

// Hamburger mobile toggle
const menuBtn = document.getElementById("menuBtn"), sidebar = document.getElementById("sidebar"), overlay = document.getElementById("overlay");
function closeMenu() { sidebar.classList.remove("open"); overlay.classList.remove("open"); menuBtn.classList.remove("open"); document.body.style.overflow=""; }
menuBtn.addEventListener("click", () => sidebar.classList.contains("open") ? closeMenu() : (sidebar.classList.add("open"), overlay.classList.add("open"), menuBtn.classList.add("open"), document.body.style.overflow="hidden"));
overlay.addEventListener("click", closeMenu);
[selPilot, selMonth].forEach(el => el.addEventListener("change", () => { if(window.innerWidth <= 768) closeMenu(); }));
document.querySelectorAll('.nav-item').forEach(el => el.addEventListener("click", () => { if(window.innerWidth <= 768) closeMenu(); }));


// ── MAIN RENDER ──
function updateDashboard() {
    const p = selPilot.value; const g = selGroup.value; const m = selMonth.value;
    if(!p || !m) return;

    document.getElementById("placeholder").style.display = "none";
    document.getElementById("dashboard").style.display = "flex";

    const pr = RAW.filter(r => r.name === p);
    const gr = RAW.filter(r => r.pos_group === g);

    renderKPIs(pr, gr, m, p);

    const tc = document.getElementById('tabContent');
    tc.innerHTML = '';
    if(blockChartInst) blockChartInst.destroy();
    if(compareChartInst) compareChartInst.destroy();
    if(dutyChartInst) dutyChartInst.destroy();

    // RENDERIZAR TABS
    if (currentTab === 'resumen') {
        tc.innerHTML = `
            <div class="card">
              <div class="card-head">
                <div><div class="card-title">Block hours \u00b7 Evoluci\u00f3n mensual</div><div class="card-sub">Piloto vs. promedio del cargo (meses activos)</div></div>
                <div class="legend">
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--clay)" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="var(--clay)"/></svg><span>Efectuado</span></div>
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--dusk)" stroke-width="1.5" stroke-dasharray="2 2"/><rect x="5.5" y="1.5" width="5" height="5" transform="rotate(45 9 4)" fill="var(--dusk)"/></svg><span style="color:var(--dusk)">Solo programado</span></div>
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--sand-400)" stroke-width="1.5" stroke-dasharray="4 3"/><circle cx="9" cy="4" r="2.5" fill="var(--sand-400)"/></svg><span>Prom. cargo</span></div>
                  <div class="leg"><svg width="14" height="12"><polygon points="7,1 13,11 1,11" fill="none" stroke="var(--rust)" stroke-width="1.5"/></svg><span style="color:var(--rust)">Excluido prom.</span></div>
                </div>
              </div>
              <div class="chart-wrap"><canvas id="blockChart"></canvas></div>
              <div class="excl-note" id="exclNote" style="display:none"><svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg><span id="exclText"></span></div>
            </div>
            <div class="charts-row">
              <div class="card">
                <div class="card-head">
                  <div><div class="card-title">Rol Programado vs. Efectuado</div><div class="card-sub">Block hours por per\u00edodo</div></div>
                  <div class="legend"><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--dusk);display:inline-block"></span><span>Programado</span></div><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--clay);display:inline-block"></span><span>Efectuado</span></div></div>
                </div>
                <div class="chart-wrap"><canvas id="compareChart"></canvas></div>
              </div>
              <div class="card">
                <div class="card-head"><div class="card-title">Comparativo por Per\u00edodo</div><div class="card-sub">Programado vs. efectuado \u00b7 \u0394 horas</div></div>
                <div id="compTableWrap"></div>
              </div>
            </div>
            <div class="bottom-row">
              <div class="card"><div class="card-head"><div class="card-title">Acumulado &amp; Proyecci\u00f3n</div><div class="card-sub">Basado en meses activos</div></div><div class="prog-list" id="progList"></div></div>
              <div class="card"><div class="card-head"><div class="card-title">Cumplimiento DAN 121</div><div class="card-sub">Mes seleccionado</div></div><div class="alert-list" id="alertList"></div></div>
            </div>`;
        renderBlockChart(pr, gr, 'blockChart', p); renderCompareChart(pr, 'compareChart');
        renderTable(pr, 'compTableWrap'); renderProgress(pr, 'progList'); renderAlerts(pr, m, 'alertList');
    
    } else if (currentTab === 'block') {
        tc.innerHTML = `
            <div class="card">
              <div class="card-head">
                <div><div class="card-title">Block hours \u00b7 Historial Expandido</div><div class="card-sub">Evoluci\u00f3n mensual y comparativa con el cargo</div></div>
                <div class="legend">
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--clay)" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="var(--clay)"/></svg><span>Efectuado</span></div>
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--dusk)" stroke-width="1.5" stroke-dasharray="2 2"/><rect x="5.5" y="1.5" width="5" height="5" transform="rotate(45 9 4)" fill="var(--dusk)"/></svg><span style="color:var(--dusk)">Solo programado</span></div>
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--sand-400)" stroke-width="1.5" stroke-dasharray="4 3"/><circle cx="9" cy="4" r="2.5" fill="var(--sand-400)"/></svg><span>Prom. cargo</span></div>
                  <div class="leg"><svg width="14" height="12"><polygon points="7,1 13,11 1,11" fill="none" stroke="var(--rust)" stroke-width="1.5"/></svg><span style="color:var(--rust)">Excluido prom.</span></div>
                </div>
              </div>
              <div class="chart-wrap" style="height:350px"><canvas id="blockChart"></canvas></div>
              <div class="excl-note" id="exclNote" style="display:none"><svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg><span id="exclText"></span></div>
            </div>
            <div class="charts-row">
              <div class="card">
                <div class="card-head"><div><div class="card-title">Distribuci\u00f3n Prog vs Ef</div></div><div class="legend"><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--dusk);display:inline-block"></span><span>Programado</span></div><div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--clay);display:inline-block"></span><span>Efectuado</span></div></div></div>
                <div class="chart-wrap"><canvas id="compareChart"></canvas></div>
              </div>
              <div class="card" style="overflow-y:auto; max-height:300px;">
                <div class="card-head" style="position:sticky;top:0;background:var(--surface);z-index:2;margin-bottom:0;padding-bottom:10px;"><div class="card-title">Historial Completo de Block hours</div></div>
                <div id="compTableWrap"></div>
              </div>
            </div>`;
        renderBlockChart(pr, gr, 'blockChart', p); renderCompareChart(pr, 'compareChart'); renderTable(pr, 'compTableWrap');

    } else if (currentTab === 'duty') {
        tc.innerHTML = `
            <div class="card">
              <div class="card-head">
                <div><div class="card-title">Duty hours \u00b7 Evoluci\u00f3n Mensual</div><div class="card-sub">Piloto vs. promedio del cargo</div></div>
                <div class="legend">
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--teal)" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="var(--teal)"/></svg><span>Duty hours</span></div>
                  <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--sand-400)" stroke-width="1.5" stroke-dasharray="4 3"/><circle cx="9" cy="4" r="2.5" fill="var(--sand-400)"/></svg><span>Prom. cargo</span></div>
                </div>
              </div>
              <div class="chart-wrap" style="height:350px"><canvas id="dutyChart"></canvas></div>
            </div>
            <div class="card">
              <div class="card-head"><div class="card-title">Historial Completo de Duty hours</div></div>
              <div id="dutyTableWrap"></div>
            </div>`;
        renderDutyChart(pr, gr, 'dutyChart', p); renderDutyTable(pr, 'dutyTableWrap');

    } else if (currentTab === 'dan121') {
        tc.innerHTML = `
            <div class="charts-row">
                <div class="card"><div class="card-head"><div class="card-title">Acumulado &amp; Proyecci\u00f3n Anual</div><div class="card-sub">Basado en meses activos</div></div><div class="prog-list" id="progList" style="gap:24px; padding-top:10px;"></div></div>
                <div class="card"><div class="card-head"><div class="card-title">Alertas Regulatorias</div><div class="card-sub">Mes seleccionado: ${PERIOD_LABELS[m]||m}</div></div><div class="alert-list" id="alertList" style="gap:14px; padding-top:10px;"></div></div>
            </div>`;
        renderProgress(pr, 'progList'); renderAlerts(pr, m, 'alertList');
    }
}

// ── COMPONENT RENDERERS ──

function renderKPIs(pr, gr, lp, p) {
    const sel = pr.find(r => r.period === lp) || pr[0];
    const ga = gr.filter(r => r.period === lp && r.name !== p && !r.exclude_from_avg && bestBlock(r) > 0);
    const ab = avg(ga.map(r => r.block_h_actual || 0).filter(v => v > 0));
    const ad = avg(ga.map(r => r.duty_h_actual  || 0).filter(v => v > 0));
    const al = avg(ga.map(r => r.libre_days     || 0).filter(v => v > 0));

    const mb = sel ? (sel.block_h_actual || sel.block_h_programmed || 0) : 0;
    const md = sel ? (sel.duty_h_actual  || sel.duty_h_programmed  || 0) : 0;
    const ml = sel ? (sel.libre_days     || 0) : 0;
    
    const isProg = sel && !(sel.block_h_actual > 0);
    const bd = ab > 0 ? (mb-ab)/ab*100 : 0;
    const dd = ad > 0 ? (md-ad)/ad*100 : 0;

    const actP = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);
    const accB = actP.reduce((s,r) => s + (r.block_h_actual || 0), 0);

    const turnos  = sel ? (sel.turnos_programados || null) : null;
    const vuelos  = sel ? (sel.vuelos_efectuados  || null) : null;
    const vProg   = sel ? (sel.vuelos_programados || null) : null;
    const blancos = sel ? (sel.dias_blancos        || null) : null;
    const progTag = isProg ? ' <span style="font-size:9px;color:var(--dusk);font-family:var(--mono)">(prog.)</span>' : "";

    document.getElementById("kpiRow").innerHTML =
        `<div class="kpi k-p1"><div class="kpi-label">Block hours \u00b7 ${PERIOD_LABELS[lp]||lp}</div><div class="kpi-val">${fmt(mb)}<span class="kpi-unit">h</span>${progTag}</div><div class="kpi-footer"><span class="kpi-vs">Prom. cargo: <b>${fmt(ab)}h</b></span><span class="delta ${dc(bd)}">${ds(bd)}</span></div></div>` +
        `<div class="kpi k-p2"><div class="kpi-label">Duty hours \u00b7 ${PERIOD_LABELS[lp]||lp}</div><div class="kpi-val">${fmt(md)}<span class="kpi-unit">h</span>${progTag}</div><div class="kpi-footer"><span class="kpi-vs">Prom. cargo: <b>${fmt(ad)}h</b></span><span class="delta ${dc(dd)}">${ds(dd)}</span></div></div>` +
        `<div class="kpi k-g1"><div class="kpi-label">D\u00edas libres \u00b7 ${PERIOD_LABELS[lp]||lp}</div><div class="kpi-val">${ml}<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Prom. cargo: <b>${fmt(al,0)}d</b></span><span class="delta ${dc(ml-al)}">${(ml-al>=0?"+":"")}${(ml-al).toFixed(0)}d</span></div></div>` +
        `<div class="kpi k-g2"><div class="kpi-label">Turnos prog. \u00b7 ${PERIOD_LABELS[lp]||lp}</div><div class="kpi-val">${turnos !== null ? turnos : "\u2014"}<span class="kpi-unit">${vuelos !== null ? " / "+vuelos+" ef." : ""}</span></div><div class="kpi-footer"><span class="kpi-vs">${vProg !== null ? vProg+" vuelos prog." : "Sin datos prog."}</span></div></div>` +
        `<div class="kpi k-g3"><div class="kpi-label">Block hours YTD</div><div class="kpi-val">${fmt(accB,0)}<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">${actP.length} meses activos</span><span class="delta d-neu">/${PERIODS.length}m</span></div></div>` +
        `<div class="kpi k-p2"><div class="kpi-label">D\u00edas blancos \u00b7 ${PERIOD_LABELS[lp]||lp}</div><div class="kpi-val">${blancos !== null ? blancos : "\u2014"}<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Sin asignaci\u00f3n (rol ef.)</span><span class="delta ${blancos > 5 ? "d-warn" : "d-up"}">${blancos !== null ? (blancos > 5 ? "\u26a0 revisar" : "\u2713 ok") : "\u2014"}</span></div></div>`;
}

function renderBlockChart(pr, gr, canvasId, p) {
    const excl = pr.filter(r => r.exclude_from_avg).map(r => r.period);
    const pData = PERIODS.map(pd => { const r = pr.find(x => x.period===pd); return r ? bestBlock(r) : null; });
    const progOnlyIdx = PERIODS.map((pd,i) => { const r = pr.find(x => x.period===pd); return (r && isProgOnly(r)) ? i : -1; }).filter(i => i>=0);
    const gData = PERIODS.map(pd => {
        const peers = gr.filter(r => r.period === pd && r.name !== p && !r.exclude_from_avg && bestBlock(r) > 0);
        return peers.length ? avg(peers.map(r => bestBlock(r))) : null;
    });

    const ctx = document.getElementById(canvasId).getContext('2d');
    blockChartInst = new Chart(ctx, {
        type: 'line',
        data: { labels: PERIODS.map(pd => PERIOD_LABELS[pd]||pd), datasets: [
            { label:'Piloto', data:pData, borderColor:'#26D800',
              backgroundColor(c) { return makeGrad(ctx, c.chart.chartArea, 'rgba(38,216,0,.12)', 'rgba(38,216,0,.01)'); },
              borderWidth:2.5,
              pointRadius(c) { return excl.includes(PERIODS[c.dataIndex]) ? 6 : 4; },
              pointStyle(c)  { return excl.includes(PERIODS[c.dataIndex]) ? 'triangle' : progOnlyIdx.includes(c.dataIndex) ? 'rectRot' : 'circle'; },
              pointBackgroundColor(c) { return excl.includes(PERIODS[c.dataIndex]) ? '#5CF200' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#26D800'; },
              pointBorderColor(c) { return excl.includes(PERIODS[c.dataIndex]) ? '#5CF200' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#26D800'; },
              pointHoverRadius:7, tension:.35, fill:true, spanGaps:true, order:1 },
            { label:'Prom. cargo', data:gData, borderColor:'#8B6FA8', borderWidth:1.5, borderDash:[5,4], pointBackgroundColor:'#8B6FA8', pointRadius:3, pointHoverRadius:5, tension:.35, fill:false, spanGaps:false, order:2 }
        ]},
        options: { responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},
            plugins: { legend:{display:false}, tooltip: { ...tooltipStyle,
                callbacks:{
                    title(i) { const pd=PERIODS[i[0].dataIndex]; const ex=excl.includes(pd); const po=progOnlyIdx.includes(i[0].dataIndex); return (PERIOD_LABELS[pd]||pd)+(ex?" \u00b7 \u26a0 excluido del prom.":po?" \u00b7 solo programado":""); },
                    label(i) { if(i.raw==null||i.raw===0)return null; const po=progOnlyIdx.includes(i.dataIndex)&&i.datasetIndex===0; return "  "+i.dataset.label+(po?" (prog.)":"")+": "+i.raw.toFixed(1)+"h"; },
                    afterBody(i) { const pd=PERIODS[i[0].dataIndex]; const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex]; if(av==null||my==null||my===0)return[]; const d=my-av; return["  vs prom. cargo: "+(d>=0?"+":"")+d.toFixed(1)+"h"]; }
                }
            }},
            scales:{
                x:{grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:"'DM Mono',monospace"}},border:{display:false}},
                y:{min:0,grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:"'DM Mono',monospace"},callback:(v)=>v+'h'},border:{display:false}}
            }
        }
    });

    const en = document.getElementById('exclNote');
    if(en) {
        const ep = excl.map(pd => PERIOD_LABELS[pd]||pd).filter(Boolean);
        if (ep.length) { en.style.display='flex'; document.getElementById('exclText').textContent='Meses excluidos del promedio comparativo: '+ep.join(', ')+'. (Tri\u00e1ngulo)'; }
        else en.style.display='none';
    }
}

function renderCompareChart(pr, canvasId) {
    const prog = PERIODS.map(pd => { const r=pr.find(x=>x.period===pd); return r?(r.block_h_programmed||0):0; });
    const act  = PERIODS.map(pd => { const r=pr.find(x=>x.period===pd); return r?(r.block_h_actual||0):0; });
    const ctx = document.getElementById(canvasId).getContext('2d');
    compareChartInst = new Chart(ctx, {
        type:'bar',
        data:{labels:PERIODS.map(pd=>PERIOD_LABELS[pd]||pd),datasets:[
            {label:'Programado',data:prog,backgroundColor:'rgba(103,30,119,.5)',borderColor:'#9B44B8',borderWidth:1,borderRadius:5,borderSkipped:false},
            {label:'Efectuado', data:act, backgroundColor:'rgba(38,216,0,.4)',borderColor:'#26D800',borderWidth:1,borderRadius:5,borderSkipped:false}
        ]},
        options:{responsive:true,maintainAspectRatio:false,
            plugins:{legend:{display:false}, tooltip: { ...tooltipStyle,
                callbacks:{
                    label(i){return "  "+i.dataset.label+": "+i.raw.toFixed(1)+"h";},
                    afterBody(i){const idx=i[0].dataIndex;const d=act[idx]-prog[idx];if(prog[idx]===0&&act[idx]===0)return["  Sin datos"];const w=d>.5?"\u25b2 Ef. mayor":d<-.5?"\u25bc Prog. mayor":"\u2248 Similares";return["  \u0394: "+(d>=0?"+":"")+d.toFixed(1)+"h  "+w];}
                }
            }},
            scales:{
                x:{grid:{display:false},ticks:{color:'#8B6FA8',font:{size:11,family:"'DM Mono',monospace"}},border:{display:false}},
                y:{min:0,grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:"'DM Mono',monospace"},callback:(v)=>v+'h'},border:{display:false}}
            }
        }
    });
}

function renderDutyChart(pr, gr, canvasId, p) {
    const pData = PERIODS.map(pd => { const r = pr.find(x => x.period===pd); return r ? bestDuty(r) : null; });
    const gData = PERIODS.map(pd => {
        const peers = gr.filter(r => r.period === pd && r.name !== p && bestDuty(r) > 0);
        return peers.length ? avg(peers.map(r => bestDuty(r))) : null;
    });

    const ctx = document.getElementById(canvasId).getContext('2d');
    dutyChartInst = new Chart(ctx, {
        type: 'line',
        data: { labels: PERIODS.map(pd => PERIOD_LABELS[pd]||pd), datasets: [
            { label:'Piloto', data:pData, borderColor:'#00C89B',
              backgroundColor(c) { return makeGrad(ctx, c.chart.chartArea, 'rgba(0,200,155,.12)', 'rgba(0,200,155,.01)'); },
              borderWidth:2.5, pointRadius:4, pointBackgroundColor:'#00C89B', pointBorderColor:'#00C89B',
              pointHoverRadius:7, tension:.35, fill:true, spanGaps:true, order:1 },
            { label:'Prom. cargo', data:gData, borderColor:'#8B6FA8', borderWidth:1.5, borderDash:[5,4], pointBackgroundColor:'#8B6FA8', pointRadius:3, pointHoverRadius:5, tension:.35, fill:false, spanGaps:false, order:2 }
        ]},
        options: { responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},
            plugins: { legend:{display:false}, tooltip: { ...tooltipStyle,
                callbacks:{
                    title(i) { return PERIOD_LABELS[PERIODS[i[0].dataIndex]] || PERIODS[i[0].dataIndex]; },
                    label(i) { if(i.raw==null||i.raw===0)return null; return "  "+i.dataset.label+": "+i.raw.toFixed(1)+"h"; },
                    afterBody(i) { const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex]; if(av==null||my==null||my===0)return[]; const d=my-av; return["  vs prom. cargo: "+(d>=0?"+":"")+d.toFixed(1)+"h"]; }
                }
            }},
            scales:{
                x:{grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:"'DM Mono',monospace"}},border:{display:false}},
                y:{min:0,grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:"'DM Mono',monospace"},callback:(v)=>v+'h'},border:{display:false}}
            }
        }
    });
}

function renderTable(pr, containerId) {
    let tbl = '<table class="comp-table"><thead><tr><th>Per\u00edodo</th><th>Block prog.</th><th>Block ef.</th><th>\u0394 Block</th><th>Turnos</th><th>Vuelos prog.</th><th>Vuelos ef.</th><th>Blancos</th></tr></thead><tbody>';
    PERIODS.forEach(pd => {
        const r = pr.find(x => x.period===pd); if(!r) return;
        const pg=r.block_h_programmed||0, ac=r.block_h_actual||0, d=ac-pg;
        const tp=r.turnos_programados, vp=r.vuelos_programados;
        const ve=r.vuelos_efectuados,  bl=r.dias_blancos;
        const ex = r.exclude_from_avg ? '<span style="color:var(--warn);font-size:9px"> \u2731</span>' : "";
        const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\u2014";
        const blCell = bl !== null ? (bl > 5 ? `<span style="color:var(--danger)">${bl}</span>` : bl) : "\u2014";
        tbl += `<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">${(PERIOD_LABELS[pd]||pd)}${ex}</td>`
             + `<td style="font-family:var(--mono)">${(pg>0?pg.toFixed(1)+"h":"\u2014")}</td>`
             + `<td style="font-family:var(--mono)">${(ac>0?ac.toFixed(1)+"h":"\u2014")}</td>`
             + `<td style="font-family:var(--mono);color:${(d>=0?"var(--teal)":"var(--danger)")}">${dstr}</td>`
             + `<td style="font-family:var(--mono)">${(tp !== null ? tp : "\u2014")}</td>`
             + `<td style="font-family:var(--mono)">${(vp !== null ? vp : "\u2014")}</td>`
             + `<td style="font-family:var(--mono)">${(ve !== null ? ve : "\u2014")}</td>`
             + `<td style="font-family:var(--mono)">${blCell}</td></tr>`;
    });
    if (pr.some(r => r.exclude_from_avg)) tbl += '<tr><td colspan="8" style="font-size:9px;color:var(--muted);font-family:var(--mono);padding:6px 10px">\u2731 Excluido del promedio comparativo</td></tr>';
    tbl += "</tbody></table>";
    document.getElementById(containerId).innerHTML = tbl;
}

function renderDutyTable(pr, containerId) {
    let tbl = '<table class="comp-table"><thead><tr><th>Per\u00edodo</th><th>Duty prog.</th><th>Duty ef.</th><th>\u0394 Duty</th></tr></thead><tbody>';
    PERIODS.forEach(pd => {
        const r = pr.find(x => x.period===pd); if(!r) return;
        const pg=r.duty_h_programmed||0, ac=r.duty_h_actual||0, d=ac-pg;
        const ex = r.exclude_from_avg ? '<span style="color:var(--warn);font-size:9px"> \u2731</span>' : "";
        const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\u2014";
        tbl += `<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">${(PERIOD_LABELS[pd]||pd)}${ex}</td>`
             + `<td style="font-family:var(--mono)">${(pg>0?pg.toFixed(1)+"h":"\u2014")}</td>`
             + `<td style="font-family:var(--mono)">${(ac>0?ac.toFixed(1)+"h":"\u2014")}</td>`
             + `<td style="font-family:var(--mono);color:${(d>=0?"var(--teal)":"var(--danger)")}">${dstr}</td></tr>`;
    });
    tbl += "</tbody></table>";
    document.getElementById(containerId).innerHTML = tbl;
}

function renderProgress(pr, containerId) {
    const actP = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);
    const accB = actP.reduce((s,r) => s + (r.block_h_actual || 0), 0);
    const excl = pr.filter(r => r.exclude_from_avg).map(r => r.period);
    const pct1 = Math.min(accB/1000*100, 100);
    const avgM = actP.length ? accB/actP.length : 0;
    const proj = avgM * 12;
    const pctP = Math.min(proj/1000*100, 100);
    const totL = pr.reduce((s,r) => s+(r.libre_days||0), 0);
    const avgL = pr.length ? totL/pr.length : 0;

    document.getElementById(containerId).innerHTML =
        `<div><div class="prog-head"><span class="prog-lbl">Block hours acumuladas YTD</span><span class="prog-num" style="color:var(--purple)">${accB.toFixed(0)}h</span></div><div class="prog-track"><div class="prog-fill" style="width:${pct1}%;background:var(--purple)"></div></div><div class="prog-note">L\u00edmite DAN 121: 1.000h/a\u00f1o \u00b7 ${(100-pct1).toFixed(1)}% disponible</div></div>` +
        `<div><div class="prog-head"><span class="prog-lbl">Proyecci\u00f3n a 12 meses</span><span class="prog-num" style="color:var(--violet)">~${proj.toFixed(0)}h est.</span></div><div class="prog-track"><div class="prog-fill" style="width:${pctP}%;background:linear-gradient(90deg,var(--green),var(--teal))"></div></div><div class="prog-note">Prom. ${avgM.toFixed(1)}h/mes en meses activos</div></div>` +
        `<div><div class="prog-head"><span class="prog-lbl">Descanso promedio</span><span class="prog-num" style="color:var(--teal)">${avgL.toFixed(1)} d/mes</span></div><div class="prog-track"><div class="prog-fill" style="width:${Math.min(avgL/20*100,100)}%;background:var(--teal)"></div></div><div class="prog-note">M\u00ednimo reglamentario DAN 121: 8 d\u00edas/mes</div></div>` +
        `<div><div class="prog-head"><span class="prog-lbl">Meses activos</span><span class="prog-num">${actP.length} / ${PERIODS.length}</span></div><div style="display:flex;gap:3px;margin-top:4px"><div style="height:5px;border-radius:2px 0 0 2px;background:var(--teal);flex:${actP.length}"></div><div style="height:5px;border-radius:0 2px 2px 0;background:rgba(103,30,119,.3);opacity:1;flex:${Math.max(PERIODS.length-actP.length,0)}"></div></div><div class="prog-note">${(excl.length?excl.map(p=>PERIOD_LABELS[p]||p).join(", ")+" excluidos":"Sin ausencias prolongadas")}</div></div>`;
}

function renderAlerts(pr, lp, containerId) {
    const sel = pr.find(r => r.period === lp) || pr[0];
    const mb = sel ? (sel.block_h_actual || sel.block_h_programmed || 0) : 0;
    const md = sel ? (sel.duty_h_actual  || sel.duty_h_programmed  || 0) : 0;
    const ml = sel ? (sel.libre_days     || 0) : 0;
    const actP = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);
    const accB = actP.reduce((s,r) => s + (r.block_h_actual || 0), 0);

    function alrt(t,title,desc){return `<div class="alert ${t}"><div class="alert-dot"></div><div><div class="alert-title">${title}</div><div class="alert-desc">${desc}</div></div></div>`;}
    
    let alerts = "";
    alerts += alrt(mb>100?"danger":mb>85?"warn":"ok", "Block hours mensual \u00b7 "+fmt(mb)+"h", mb>100?"Supera l\u00edmite DAN 121 de 100h/mes":mb>85?"Cercano al l\u00edmite de 100h/mes":"Dentro del l\u00edmite (100h/mes)");
    alerts += alrt(accB>900?"danger":accB>750?"warn":"ok", "Block hours acumuladas \u00b7 "+accB.toFixed(0)+"h", accB>900?"Muy cerca del l\u00edmite anual de 1.000h":accB>750?"Supera el 75% del l\u00edmite anual":"Sin riesgo l\u00edmite anual ("+(1000-accB).toFixed(0)+"h disp.)");
    alerts += alrt(ml<8?"danger":ml<10?"warn":"ok", "D\u00edas libres \u00b7 "+ml+"d", ml<8?"Bajo el m\u00ednimo reglamentario (8d/mes)":ml<10?"Dentro del m\u00ednimo, bajo el promedio del cargo":"Descanso adecuado seg\u00fan DAN 121");
    alerts += alrt(md>130?"danger":md>105?"warn":"ok", "Duty hours \u00b7 "+fmt(md)+"h", md>130?"Duty hours muy elevadas, revisar FDPs":md>105?"Sobre promedio del cargo":"Dentro de rango normal");
    alerts += '<div style="margin-top:6px;padding:9px 11px;background:var(--s2);border-radius:7px;font-size:10px;color:var(--muted);line-height:1.5;font-family:var(--mono)">Alertas indicativas. El c\u00e1lculo oficial de FDP y l\u00edmites es responsabilidad de Operaciones.</div>';
    
    document.getElementById(containerId).innerHTML = alerts;
}

"""

def generate_html(records, periods):
    period_labels = {p: PERIOD_LABELS_MAP.get(p, p) for p in periods}
    DATA_JS    = json.dumps(records,       ensure_ascii=False, default=str)
    PERIODS_JS = json.dumps(periods,       ensure_ascii=False)
    LABELS_JS  = json.dumps(period_labels, ensure_ascii=False)

    js_content = JS_TEMPLATE.replace('__RAW__', DATA_JS).replace('__PERIODS__', PERIODS_JS).replace('__LABELS__', LABELS_JS)
    html_content = HTML_TEMPLATE.replace('__JS_CONTENT__', js_content)
    
    return html_content

# ── MAIN ───────────────────────────────────────────────────
if __name__ == '__main__':
    records, periods = build_dataset()
    html = generate_html(records, periods)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print('\nDashboard generado: ' + str(OUTPUT_HTML))
    print('Tamano: ' + str(len(html)//1024) + ' KB')
