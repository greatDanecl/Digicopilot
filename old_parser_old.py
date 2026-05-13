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
    """Extrae YYYY-MM del nombre del archivo.

    Soporta:
      202602_efec_SCL.xlsx     → 2026-02
      efec_feb_2026_SCL.xlsx   → 2026-02
      prog_mar_2026_PMC.xlsx   → 2026-03
      2026-02_prog_SCL.xlsx    → 2026-02
    """
    fl = fname.lower()

    # Formato AAAAMM pegado: 202602, 202510, etc.
    m = re.search(r'(20\d{2})(0[1-9]|1[0-2])', fl)
    if m:
        return m.group(1) + '-' + m.group(2)

    # Formato AAAA-MM o AAAA_MM
    m2 = re.search(r'(20\d{2})[-_](0[1-9]|1[0-2])', fl)
    if m2:
        return m2.group(1) + '-' + m2.group(2)

    # Mes en texto + año: feb_2026, mar2026, etc.
    month_re = r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|ene|abr|ago|dic)'
    year_re  = r'(20\d{2})'
    m3 = re.search(month_re + r'[_\-\s]*' + year_re, fl)
    if m3:
        code = MONTH_MAP.get(m3.group(1), '00')
        if code != '00': return m3.group(2) + '-' + code

    # Año + mes en texto: 2026_feb
    m4 = re.search(year_re + r'[_\-\s]*' + month_re, fl)
    if m4:
        code = MONTH_MAP.get(m4.group(2), '00')
        if code != '00': return m4.group(1) + '-' + code

    return None

def detect_period_from_df(df):
    """Extrae el período del contenido del archivo cuando el nombre no alcanza."""
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
    """Detecta si el archivo es rol programado o efectuado.

    Soporta:
      202602_efec_SCL.xlsx   → actual
      202602_prog_SCL.xlsx   → programmed
      efec_feb_2026_SCL.xlsx → actual
      prog_mar_2026_PMC.xlsx → programmed
    """
    fl, sl = fname.lower(), sheet_name.lower()
    # Busca 'efec' o 'prog' en cualquier posición del nombre
    if 'efec' in fl: return 'actual'
    if 'prog' in fl: return 'programmed'
    # Variantes largas
    if 'efectuado' in fl or 'actual' in fl or 'flown' in fl: return 'actual'
    if 'programado' in fl or 'plan' in fl or 'sched' in fl:  return 'programmed'
    # Fallback por nombre de hoja
    if any(w in sl for w in ['hora', 'actual', 'efect']): return 'actual'
    return 'programmed'

def classify_day(col_vals):
    """Clasifica un dia dado los valores de todas las filas de esa columna.
    
    Retorna uno de:
      TURNO    - dia con turno de reserva (TURNOx)
      VUELO    - dia con vuelo(s) operados
      HOTEL    - pernocte fuera de base
      SIM      - simulador
      ELEAR    - e-learning
      DH       - deadhead (pasajero en avion de la compania)
      ACT      - activacion desde reserva sin turno asignado
      OFNA2    - trabajo en oficina
      CEMAE    - examen medico CMAE
      LQUIN    - libre quincena
      SINDI    - dia sindical
      FVUEL    - fuera de vuelo (pre/post natal u otro)
      VACAC    - vacaciones
      LIBRE    - dia libre en base
      FINDE    - dias libres ley (art 152 ter K)
      BDAY     - cumpleanos
      BLANCO   - sin asignacion
    """
    clean = [v for v in col_vals if v not in ['nan', 'NaT', '']]
    if not clean:
        return 'BLANCO'
    joined = ' '.join(clean).upper()

    # TURNO: debe ser explicitamente TURNOx (numero)
    if re.search(r'\bTURNO\d+', joined):
        return 'TURNO'
    # Vuelos: numeros puros de 2-4 digitos en cualquier fila
    flights = [v for v in clean if re.match(r'^\d{2,4}$', v.strip())]
    if flights:
        return 'VUELO'
    # Resto de codigos por prioridad
    for code, label in [
        ('HOTEL',  'HOTEL'),
        ('SIM',    'SIM'),
        ('ELEAR',  'ELEAR'),
        ('DH',     'DH'),
        ('ACT',    'ACT'),
        ('OFNA2',  'OFNA2'),
        ('CEMAE',  'CEMAE'),
        ('LQUIN',  'LQUIN'),
        ('SINDI',  'SINDI'),
        ('FVUEL',  'FVUEL'),
        ('VACAC',  'VACAC'),
        ('BDAY',   'BDAY'),
        ('LIBRE',  'LIBRE'),
        ('FINDE',  'FINDE'),
    ]:
        if code in joined:
            return label
    # Solo horas sin codigo reconocido = continuacion de vuelo del dia anterior
    return 'CONT'

def get_day_cols(df, pilot_row, col, n_rows=4):
    """Devuelve los valores de n_rows filas para una columna de un dia."""
    col_start = 2
    return [
        str(df.iloc[pilot_row + k, col + col_start]).strip()
        if pilot_row + k < len(df) else ''
        for k in range(n_rows)
    ]

def count_schedule(df, pilot_row, role):
    """Cuenta turnos, vuelos, dias blancos y dias por tipo para un piloto.
    
    Para rol 'programmed': cuenta TURNOs y VUELOs programados.
    Para rol 'actual':     cuenta VUELOs efectuados y dias BLANCO.
    """
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
    """Busca dinámicamente Credits, Block hours, Duty hours desde pilot_row."""
    cred_h = duty_h = blk_h = 0.0
    for k in range(1, max_look):
        row = pilot_row + k
        if row >= len(df): break
        lbl = str(df.iloc[row, 0]).strip()
        # Stop if we hit the next pilot
        if re.match(r'^[A-Z]{4,5}$', lbl) and k > 5:
            break
        if lbl == 'Credits':       cred_h = parse_td(df.iloc[row, 1])
        elif lbl == 'Block hours': blk_h  = parse_td(df.iloc[row, 1])
        elif lbl == 'Duty hours':  duty_h = parse_td(df.iloc[row, 1])
    return cred_h, blk_h, duty_h

def block_size(df, pilot_row, max_look=18):
    """Determina cuántas filas ocupa el bloque buscando el siguiente código de piloto."""
    for k in range(5, max_look):
        row = pilot_row + k
        if row >= len(df): return k
        c0 = str(df.iloc[row, 0]).strip()
        # Next pilot starts a new block
        if re.match(r'^[A-Z]{4,5}$', c0):
            return k
    return 13  # fallback generoso

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
            # In ABCD format: col0=label(F/G/H), col1=field name, col2=value
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
            # Dynamic search for totals — handles variable row offsets
            cred_h, blk_h, duty_h = find_totals(df, i)
            sched   = [str(v).strip() for v in df.iloc[i, 2:].tolist()]
            pilot_row = i
            i += block_size(df, i)

        if not re.match(r'^[A-Z]{4,5}$', code): continue
        # Handle multi-position strings like "19.243.849-7 - CP, C15M"
        pos_raw = rut_pos.split(' - ')[-1].strip() if ' - ' in rut_pos else ''
        # Take first position if multiple separated by comma
        pos = pos_raw.split(',')[0].strip()
        if not pos or pos in ['nan', 'NaT', '']: continue
        name = (fname_p + ' ' + lname).strip()
        if not name or re.search(r'\b(TEST|PRUEBA)\b', name.upper()): continue

        # Position grouping
        # C15M = Capitán (habilitación especial A320 family)
        # FON  = Primer Oficial (igual que FO)
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

        # Count turnos, vuelos, blancos using correct classification
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

    # Load config.json if present
    config_path = DATA_DIR / 'config.json'
    file_map = {}
    if config_path.exists():
        cfg = json.loads(config_path.read_text(encoding='utf-8'))
        for entry in cfg.get('files', []):
            file_map[entry['filename']] = {
                'period': entry['period'],
                'role':   entry['role'],
            }
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

                # Period: config > filename > content
                if cfg_entry:
                    period = cfg_entry['period']
                else:
                    period = detect_period_from_filename(fname) or detect_period_from_df(df)

                if not period:
                    print('  ? ' + fname + '/' + sheet_name + ': periodo no detectado, omitiendo')
                    continue

                # Role: config entry > sheet name > filename > default
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

    # Merge programmed + actual per pilot+period
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
def generate_html(records, periods):
    period_labels = {p: PERIOD_LABELS_MAP.get(p, p) for p in periods}

    DATA_JS    = json.dumps(records,       ensure_ascii=False, default=str)
    PERIODS_JS = json.dumps(periods,       ensure_ascii=False)
    LABELS_JS  = json.dumps(period_labels, ensure_ascii=False)

    # CSS — plain string, no f-string needed
    CSS = (
        ':root{'
        '--purple:#671E77;--purple-l:#9B44B8;--purple-xl:#C480E0;'
        '--purple-dim:rgba(103,30,119,0.18);--purple-dim2:rgba(103,30,119,0.08);'
        '--green:#26D800;--green-l:#5CF200;--green-dim:rgba(38,216,0,0.15);--green-dim2:rgba(38,216,0,0.07);'
        '--violet:#8B35A8;--teal:#00C89B;--teal-dim:rgba(0,200,155,0.12);'
        '--danger:#FF4466;--danger-dim:rgba(255,68,102,0.12);'
        '--warn:#C46AE0;--warn-dim:rgba(196,106,224,0.12);'
        '--bg:#F8F7FC;--surface:#FFFFFF;--s2:#F0EBF7;--s3:#E8DFF5;'
        '--border:rgba(103,30,119,0.18);--border2:rgba(103,30,119,0.35);'
        '--text:#2A1240;--text2:#5A3878;--muted:#8B6FA8;--dim:#B09CC8;'
        '--r:10px;--r2:14px;'
        '--shadow:0 1px 4px rgba(0,0,0,.4),0 4px 20px rgba(103,30,119,.15);'
        '--shadow2:0 2px 12px rgba(0,0,0,.5),0 8px 32px rgba(103,30,119,.25);'
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
        '.f-select{appearance:none;background:rgba(103,30,119,.05);border:1px solid rgba(103,30,119,.25);border-radius:8px;color:var(--text);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;background-image:url(\"data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'12\' height=\'12\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%238B6FA8\' stroke-width=\'2\'%3E%3Cpolyline points=\'6 9 12 15 18 9\'/%3E%3C/svg%3E\");background-repeat:no-repeat;background-position:right 9px center}'
        '.f-select:focus,.f-select:hover{border-color:var(--green);box-shadow:0 0 0 2px rgba(38,216,0,.15)}'
        '.f-select option{background:#FFFFFF;color:var(--text)}'
        '.sidebar-nav{padding:10px 8px;flex:1}'
        '.nav-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:var(--text2);cursor:pointer;transition:all .15s;margin-bottom:2px;border-left:2px solid transparent}'
        '.nav-item:hover{color:var(--purple);background:rgba(103,30,119,.08);border-left-color:rgba(103,30,119,.4)}'
        '.nav-item.active{color:var(--purple);background:rgba(103,30,119,.12);border-left-color:var(--purple)}'
        '.nav-item svg{width:14px;height:14px;flex-shrink:0}'
        '.sidebar-footer{padding:12px 16px;border-top:1px solid rgba(103,30,119,.15)}'
        '.pilot-badge{display:flex;align-items:center;gap:10px}'
        '.pilot-avatar{width:34px;height:34px;border-radius:50%;background:var(--purple);border:1.5px solid var(--purple-l);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}'
        '.pilot-name-s{font-size:11px;font-weight:500;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
        '.pilot-pos-s{font-size:10px;color:var(--muted);font-family:var(--mono)}'
        '.main{display:flex;flex-direction:column;min-height:100vh}'
        '.topbar{background:#FFFFFF;border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;box-shadow:0 1px 8px rgba(103,30,119,.08)}'
        '.page-title{font-family:var(--display);font-size:17px;color:var(--text)}'
        '.page-title span{color:var(--purple)}'
        '.page-sub{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.topbar-right{display:flex;align-items:center;gap:8px}'
        '.pill{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--s2);color:var(--text2)}'
        '.dot{width:6px;height:6px;border-radius:50%;background:var(--green);box-shadow:0 0 7px var(--green)}'
        '.content{padding:18px 26px;display:flex;flex-direction:column;gap:13px;flex:1}'
        '.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}'
        '.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 15px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:box-shadow .2s,transform .15s,border-color .2s}'
        '.kpi:hover{box-shadow:var(--shadow2);transform:translateY(-1px);border-color:var(--border2)}'
        ".kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;border-radius:var(--r2) var(--r2) 0 0}"
        '.kpi.k-p1::before{background:var(--purple-l)}'
        '.kpi.k-p2::before{background:var(--violet)}'
        '.kpi.k-g1::before{background:var(--green)}'
        '.kpi.k-g2::before{background:var(--teal)}'
        '.kpi.k-g3::before{background:var(--green-l)}'
        '.kpi-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;font-family:var(--mono)}'
        '.kpi-val{font-size:24px;font-weight:600;color:var(--text);font-family:var(--mono);letter-spacing:-.03em;line-height:1}'
        '.kpi-unit{font-size:12px;font-weight:400;color:var(--muted);margin-left:2px}'
        '.kpi-footer{display:flex;align-items:center;justify-content:space-between;margin-top:7px}'
        '.kpi-vs{font-size:10px;color:var(--muted)}.kpi-vs b{color:var(--text2);font-weight:500}'
        '.delta{font-size:10px;font-family:var(--mono);padding:2px 6px;border-radius:4px}'
        '.d-up{background:var(--green-dim);color:var(--green-l)}'
        '.d-down{background:var(--danger-dim);color:var(--danger)}'
        '.d-neu{background:var(--s3);color:var(--muted)}'
        '.d-warn{background:var(--warn-dim);color:var(--warn)}'
        '.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow)}'
        '.card-head{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}'
        '.card-title{font-size:13px;font-weight:500;color:var(--text)}'
        '.card-sub{font-size:10px;color:var(--muted);margin-top:2px;font-family:var(--mono)}'
        '.legend{display:flex;gap:12px;align-items:center;font-size:10px;color:var(--muted);font-family:var(--mono);flex-wrap:wrap}'
        '.leg{display:flex;align-items:center;gap:5px}'
        '.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}'
        '.chart-wrap{position:relative;height:220px}'
        '.comp-table{width:100%;border-collapse:collapse;font-size:12px}'
        '.comp-table th{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--s2);color:var(--text2)}'
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
        '.alert.ok{background:var(--green-dim2);border-color:rgba(38,216,0,.25)}'
        '.alert.warn{background:var(--warn-dim);border-color:rgba(196,106,224,.3)}'
        '.alert.danger{background:var(--danger-dim);border-color:rgba(255,68,102,.3)}'
        '.alert-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px}'
        '.alert.ok .alert-dot{background:var(--green)}'
        '.alert.warn .alert-dot{background:var(--warn)}'
        '.alert.danger .alert-dot{background:var(--danger)}'
        '.alert-title{font-size:11px;font-weight:500;color:var(--text)}'
        '.alert-desc{font-size:10px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.excl-note{display:flex;align-items:center;gap:6px;padding:8px 11px;border-radius:7px;background:var(--s2);border:1px solid var(--border);font-size:10px;color:var(--muted);margin-top:10px}'
        '.excl-note svg{width:12px;height:12px;flex-shrink:0}'
        '::-webkit-scrollbar{width:4px}'
        '::-webkit-scrollbar-thumb{background:rgba(103,30,119,.3);border-radius:2px}'
        '@keyframes fadeUp{from{opacity:0;transform:translateY(7px)}to{opacity:1;transform:none}}'
        '.kpi,.card{animation:fadeUp .28s ease both}'
        '.kpi:nth-child(1){animation-delay:.04s}.kpi:nth-child(2){animation-delay:.08s}'
        '.kpi:nth-child(3){animation-delay:.12s}.kpi:nth-child(4){animation-delay:.16s}'
        '.kpi:nth-child(5){animation-delay:.20s}.kpi:nth-child(6){animation-delay:.24s}'
        '.turnos-bar{display:flex;align-items:center;gap:8px;margin-top:6px}'
        '.tbar-track{flex:1;height:7px;background:var(--s3);border-radius:4px;overflow:hidden;position:relative}'
        '.tbar-prog{height:100%;background:var(--purple-l);border-radius:4px;position:absolute;left:0;top:0}'
        '.tbar-act{height:100%;background:var(--green);border-radius:4px;position:absolute;left:0;top:0;opacity:.85}'
        '.tbar-label{font-size:10px;font-family:var(--mono);color:var(--muted);white-space:nowrap}'
        '.hamburger{display:none;position:fixed;top:12px;left:12px;z-index:200;width:38px;height:38px;border-radius:8px;border:1px solid var(--border2);background:var(--s2);cursor:pointer;align-items:center;justify-content:center;flex-direction:column;gap:5px;padding:9px}'
        '.hamburger span{display:block;width:16px;height:1.5px;background:var(--text2);border-radius:2px;transition:all .2s}'
        '.hamburger.open span:nth-child(1){transform:translateY(6.5px) rotate(45deg)}'
        '.hamburger.open span:nth-child(2){opacity:0}'
        '.hamburger.open span:nth-child(3){transform:translateY(-6.5px) rotate(-45deg)}'
        '.sidebar-overlay{display:none;position:fixed;inset:0;background:rgba(40,20,60,.5);z-index:155;backdrop-filter:blur(2px)}'
        '.sidebar-overlay.open{display:block}'
        '@media(max-width:1024px){'
        '.kpi-grid{grid-template-columns:repeat(3,1fr)}'
        '.charts-row{grid-template-columns:1fr}'
        '.bottom-row{grid-template-columns:1fr}'
        '}'
        '@media(max-width:768px){'
        '.shell{grid-template-columns:1fr}'
        '.sidebar{position:fixed;left:-240px;top:0;height:100vh;width:230px;z-index:160;transition:left .25s cubic-bezier(.4,0,.2,1)}'
        '.sidebar.open{left:0;box-shadow:6px 0 32px rgba(46,36,22,.3)}'
        '.hamburger{display:flex}'
        '.topbar{padding:12px 16px 12px 58px}'
        '.content{padding:14px 16px}'
        '.kpi-grid{grid-template-columns:repeat(2,1fr);gap:8px}'
        '.charts-row{grid-template-columns:1fr;gap:10px}'
        '.bottom-row{grid-template-columns:1fr;gap:10px}'
        '.chart-wrap{height:190px}'
        '.page-title{font-size:14px}'
        '.page-sub{font-size:10px;margin-top:0}'
        '.card{padding:14px}'
        '.card-head{flex-direction:column;gap:8px;align-items:flex-start}'
        '.legend{gap:8px;font-size:9px}'
        '#compTableWrap{overflow-x:auto;-webkit-overflow-scrolling:touch}'
        '.comp-table th,.comp-table td{padding:6px 8px}'
        '}'
        '@media(max-width:420px){'
        '.kpi-grid{grid-template-columns:1fr 1fr}'
        '.kpi-val{font-size:21px}'
        '.kpi{padding:12px 12px}'
        '.chart-wrap{height:165px}'
        '.content{padding:10px 12px}'
        '}'
        '.kpi:nth-child(5){animation-delay:.20s}'
    )

    # JavaScript — plain string, injecting data via concatenation
    JS = (
        'const RAW = ' + DATA_JS + ';\n'
        'const PERIODS = ' + PERIODS_JS + ';\n'
        'const PERIOD_LABELS = ' + LABELS_JS + ';\n'
        '\n'
        "document.getElementById('periodPill').textContent = Object.values(PERIOD_LABELS).join(' \u00b7 ');\n"
        "document.getElementById('periodsHint').textContent = 'Per\u00edodos: ' + Object.values(PERIOD_LABELS).join(' \u00b7 ');\n"
        '\n'
        'let blockChartInst = null, compareChartInst = null;\n'
        'const selGroup = document.getElementById("selGroup");\n'
        'const selPilot = document.getElementById("selPilot");\n'
        'const selMonth = document.getElementById("selMonth");\n'
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
        '  }\n'
        '});\n'
        '\n'
        'function fmt(v, d) { d = d === undefined ? 1 : d; if (v == null || +v === 0) return "\u2014"; return (+v).toFixed(d); }\n'
        'function avg(arr) { const v = arr.filter(x => x != null && x > 0); return v.length ? v.reduce((a,b) => a+b, 0)/v.length : 0; }\n'
        'function dc(d) { return d > 2 ? "d-up" : d < -2 ? "d-down" : "d-neu"; }\n'
        'function ds(d) { return (d >= 0 ? "+" : "") + d.toFixed(1) + "%"; }\n'
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
        '  document.getElementById("pageTitle").innerHTML = "<span>" + pilotName.split(" ").slice(0,2).join(" ") + "</span> \u00b7 Productividad";\n'
        '  document.getElementById("pageSub").textContent = (latest ? latest.pos_group : group) + " \u00b7 " + (latest ? latest.base : "") + " \u00b7 " + Object.values(PERIOD_LABELS).join(" \u00b7 ");\n'
        '\n'
        '  // Populate month dropdown — only periods where pilot has data\n'
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
        '  // Default: most recent period with actual data, fallback to most recent any\n'
        '  const defaultPeriod = (latest ? latest.period : pilotPeriods[0]);\n'
        '  selMonth.value = defaultPeriod;\n'
        '  selMonth.disabled = false;\n'
        '\n'
        '  // Render charts (static per pilot)\n'
        '  renderCharts(pilotName, group, pr, gr);\n'
        '  // Render KPIs for selected month\n'
        '  renderKPIs(pilotName, group, defaultPeriod);\n'
        '}\n'
        '\n'
        'function renderKPIs(pilotName, group, selectedPeriod) {\n'
        '  const pr = RAW.filter(r => r.name === pilotName);\n'
        '  const gr = RAW.filter(r => r.pos_group === group);\n'
        '  const sel = pr.find(r => r.period === selectedPeriod) || pr[0];\n'
        '  const lp  = selectedPeriod;\n'
        '\n'
        '  // Group averages for selected period (peers, no exclusions, no self)\n'
        '  function bestBlock(r) { return (r.block_h_actual && r.block_h_actual > 0) ? r.block_h_actual : (r.block_h_programmed || 0); }\n'
        '  const ga = gr.filter(r => r.period === lp && r.name !== pilotName && !r.exclude_from_avg && bestBlock(r) > 0);\n'
        '  const ab = avg(ga.map(r => r.block_h_actual || 0).filter(v => v > 0));\n'
        '  const ad = avg(ga.map(r => r.duty_h_actual  || 0).filter(v => v > 0));\n'
        '  const al = avg(ga.map(r => r.libre_days     || 0).filter(v => v > 0));\n'
        '\n'
        '  // Use actual if available, else programmed\n'
        '  const mb = sel ? (sel.block_h_actual || sel.block_h_programmed || 0) : 0;\n'
        '  const md = sel ? (sel.duty_h_actual  || sel.duty_h_programmed  || 0) : 0;\n'
        '  const ml = sel ? (sel.libre_days     || 0) : 0;\n'
        '  const isProg = sel && !(sel.block_h_actual > 0);\n'
        '  const bd = ab > 0 ? (mb-ab)/ab*100 : 0;\n'
        '  const dd = ad > 0 ? (md-ad)/ad*100 : 0;\n'
        '\n'
        '  // YTD block hours (always all-period accumulation)\n'
        '  const actP  = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);\n'
        '  const accB  = actP.reduce((s,r) => s + (r.block_h_actual || 0), 0);\n'
        '  const excl  = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '\n'
        '  const turnos  = sel ? (sel.turnos_programados || null) : null;\n'
        '  const vuelos  = sel ? (sel.vuelos_efectuados  || null) : null;\n'
        '  const vProg   = sel ? (sel.vuelos_programados || null) : null;\n'
        '  const blancos = sel ? (sel.dias_blancos        || null) : null;\n'
        '  const convPct = (turnos != null && vuelos != null && turnos > 0) ? vuelos/turnos*100 : null;\n'
        '  const progTag = isProg ? \' <span style="font-size:9px;color:var(--dusk);font-family:var(--mono)">(prog.)</span>\' : "";\n'
        '\n'
        '  document.getElementById("kpiRow").innerHTML =\n'
        '    \'<div class="kpi k-p1"><div class="kpi-label">Block hours \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + fmt(mb) + \'<span class="kpi-unit">h</span>\' + progTag + \'</div><div class="kpi-footer"><span class="kpi-vs">Prom. cargo: <b>\' + fmt(ab) + \'h</b></span><span class="delta \' + dc(bd) + \'">\' + ds(bd) + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-p2"><div class="kpi-label">Duty hours \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + fmt(md) + \'<span class="kpi-unit">h</span>\' + progTag + \'</div><div class="kpi-footer"><span class="kpi-vs">Prom. cargo: <b>\' + fmt(ad) + \'h</b></span><span class="delta \' + dc(dd) + \'">\' + ds(dd) + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-g1"><div class="kpi-label">D\u00edas libres \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + ml + \'<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Prom. cargo: <b>\' + fmt(al,0) + \'d</b></span><span class="delta \' + dc(ml-al) + \'">\' + (ml-al>=0?"+":"") + (ml-al).toFixed(0) + \'d</span></div></div>\' +\n'
        '    \'<div class="kpi k-g2"><div class="kpi-label">Turnos prog. \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div>\' +\n'
        '    \'<div class="kpi-val">\' + (turnos !== null ? turnos : "\\u2014") + \'<span class="kpi-unit">\' + (vuelos !== null ? " / "+vuelos+" ef." : "") + \'</span></div>\' +\n'
        '    \'<div class="kpi-footer"><span class="kpi-vs">\' + (vProg !== null ? vProg+" vuelos prog." : "Sin datos prog.") + \'</span>\' +\n'
        '    \'</div></div>\' +\n'
        '    \'<div class="kpi k-g3"><div class="kpi-label">Block acum. YTD</div><div class="kpi-val">\' + fmt(accB,0) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">\' + actP.length + \' meses activos</span><span class="delta d-neu">/\' + PERIODS.length + \'m</span></div></div>\' +\n'
        '    \'<div class="kpi k-p2"><div class="kpi-label">D\u00edas blancos \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div>\' +\n'
        '    \'<div class="kpi-val">\' + (blancos !== null ? blancos : "\\u2014") + \'<span class="kpi-unit">d</span></div>\' +\n'
        '    \'<div class="kpi-footer"><span class="kpi-vs">Sin asignaci\u00f3n (rol ef.)</span><span class="delta \' + (blancos > 5 ? "d-warn" : "d-up") + \'">\' + (blancos !== null ? (blancos > 5 ? "\\u26a0 revisar" : "\\u2713 ok") : "\\u2014") + \'</span></div></div>\';\n'
        '\n'
        '  // Progress bars\n'
        '  const pct1 = Math.min(accB/1000*100, 100);\n'
        '  const avgM = actP.length ? accB/actP.length : 0;\n'
        '  const proj = avgM * 12;\n'
        '  const pctP = Math.min(proj/1000*100, 100);\n'
        '  const totL = pr.reduce((s,r) => s+(r.libre_days||0), 0);\n'
        '  const avgL = pr.length ? totL/pr.length : 0;\n'
        "  document.getElementById('progList').innerHTML =\n"
        '    \'<div><div class="prog-head"><span class="prog-lbl">Horas bloque acumuladas</span><span class="prog-num" style="color:var(--purple)">\' + accB.toFixed(0) + \'h</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pct1 + \'%;background:var(--purple)"></div></div><div class="prog-note">L\\u00edmite DAN 121: 1.000h/a\\u00f1o \\u00b7 \' + (100-pct1).toFixed(1) + \'% disponible</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Proyecci\\u00f3n a 12 meses</span><span class="prog-num" style="color:var(--violet)">~\' + proj.toFixed(0) + \'h est.</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pctP + \'%;background:linear-gradient(90deg,var(--green),var(--teal))"></div></div><div class="prog-note">Prom. \' + avgM.toFixed(1) + \'h/mes en meses activos</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Descanso promedio</span><span class="prog-num" style="color:var(--teal)">\' + avgL.toFixed(1) + \' d/mes</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + Math.min(avgL/20*100,100) + \'%;background:var(--teal)"></div></div><div class="prog-note">M\\u00ednimo reglamentario DAN 121: 8 d\\u00edas/mes</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Meses activos</span><span class="prog-num">\' + actP.length + \' / \' + PERIODS.length + \'</span></div><div style="display:flex;gap:3px;margin-top:4px"><div style="height:5px;border-radius:2px 0 0 2px;background:var(--teal);flex:\' + actP.length + \'"></div><div style="height:5px;border-radius:0 2px 2px 0;background:rgba(103,30,119,.3);opacity:1;flex:\' + Math.max(PERIODS.length-actP.length,0) + \'"></div></div><div class="prog-note">\' + (excl.length?excl.map(p=>PERIOD_LABELS[p]||p).join(", ")+" excluidos":"Sin ausencias prolongadas") + \'</div></div>\';\n'
        '\n'
        '  // Alerts\n'
        '  function alrt(t,title,desc){return \'<div class="alert \'+t+\'"><div class="alert-dot"></div><div><div class="alert-title">\'+title+\'</div><div class="alert-desc">\'+desc+\'</div></div></div>\';}\n'
        '  let alerts = "";\n'
        '  alerts += alrt(mb>100?"danger":mb>85?"warn":"ok", "Bloque mensual \\u00b7 "+fmt(mb)+"h", mb>100?"Supera l\\u00edmite DAN 121 de 100h/mes":mb>85?"Cercano al l\\u00edmite de 100h/mes":"Dentro del l\\u00edmite (100h/mes)");\n'
        '  alerts += alrt(accB>900?"danger":accB>750?"warn":"ok", "Bloque acumulado \\u00b7 "+accB.toFixed(0)+"h", accB>900?"Muy cerca del l\\u00edmite anual de 1.000h":accB>750?"Supera el 75% del l\\u00edmite anual":"Sin riesgo l\\u00edmite anual ("+(1000-accB).toFixed(0)+"h disp.)");\n'
        '  alerts += alrt(ml<8?"danger":ml<10?"warn":"ok", "D\\u00edas libres \\u00b7 "+ml+"d", ml<8?"Bajo el m\\u00ednimo reglamentario (8d/mes)":ml<10?"Dentro del m\\u00ednimo, bajo el promedio del cargo":"Descanso adecuado seg\\u00fan DAN 121");\n'
        '  alerts += alrt(md>130?"danger":md>105?"warn":"ok", "Horas deber \\u00b7 "+fmt(md)+"h", md>130?"Horas deber muy elevadas, revisar FDPs":md>105?"Sobre promedio del cargo":"Dentro de rango normal");\n'
        '  alerts += \'<div style="margin-top:6px;padding:9px 11px;background:var(--s2);border-radius:7px;font-size:10px;color:var(--muted);line-height:1.5;font-family:var(--mono)">Alertas indicativas. El c\\u00e1lculo oficial de FDP y l\\u00edmites es responsabilidad de Operaciones.</div>\';\n'
        "  document.getElementById('alertList').innerHTML = alerts;\n"

        '}\n'
        '\n'
        'function renderCharts(pilotName, group, pr, gr) {\n'
        '  // pData: muestra efectuado si existe, si no programado (meses futuros/sin efectuado aún)\n'
        '  // gData: promedio del cargo completo, pilotos activos sin ausencias prolongadas\n'
        '  //        usa efectuado si existe, si no programado — mismo criterio que pData\n'
        '  const excl  = pr.filter(r => r.exclude_from_avg).map(r => r.period);\n'
        '  function bestBlock(r) { return (r.block_h_actual && r.block_h_actual > 0) ? r.block_h_actual : (r.block_h_programmed || 0); }\n'
        '  function isProgrammedOnly(r) { return !(r.block_h_actual && r.block_h_actual > 0) && (r.block_h_programmed && r.block_h_programmed > 0); }\n'
        '  const pData = PERIODS.map(p => { const r = pr.find(x => x.period===p); return r ? bestBlock(r) : null; });\n'
        '  // progOnlyPeriods: períodos donde el piloto solo tiene programado (sin efectuado)\n'
        '  const progOnlyIdx = PERIODS.map((p,i) => { const r = pr.find(x => x.period===p); return (r && isProgrammedOnly(r)) ? i : -1; }).filter(i => i>=0);\n'
        '  // gData: promedio del segmento comparable (mismo cargo, sin ausencias)\n'
        '  // Excluye al piloto seleccionado del cálculo del promedio\n'
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
        "        backgroundColor(c) { return makeGrad(bc, c.chart.chartArea, 'rgba(38,216,0,.12)', 'rgba(38,216,0,.01)'); },\n"
        '        borderWidth:2.5,\n'
        '        pointRadius(c)          { return excl.includes(PERIODS[c.dataIndex]) ? 6 : 4; },\n'
        "        pointStyle(c)           { return excl.includes(PERIODS[c.dataIndex]) ? 'triangle' : progOnlyIdx.includes(c.dataIndex) ? 'rectRot' : 'circle'; },\n"
        "        pointBackgroundColor(c) { return excl.includes(PERIODS[c.dataIndex]) ? '#5CF200' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#26D800'; },\n"
        "        pointBorderColor(c)     { return excl.includes(PERIODS[c.dataIndex]) ? '#5CF200' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#26D800'; },\n"
        '        pointHoverRadius:7, tension:.35, fill:true, spanGaps:true, order:1 },\n'
        "      { label:'Prom. cargo', data:gData, borderColor:'#8B6FA8', borderWidth:1.5, borderDash:[5,4],\n"
        "        pointBackgroundColor:'#8B6FA8', pointRadius:3, pointHoverRadius:5,\n"
        '        tension:.35, fill:false, spanGaps:false, order:2 }\n'
        '    ]},\n'
        '    options: { responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},\n'
        '      plugins: { legend:{display:false}, tooltip:{\n'
        "        backgroundColor:'#18102A', borderColor:'rgba(103,30,119,.5)', borderWidth:1,\n"
        "        titleColor:'#F0E8F8', bodyColor:'#8B6FA8', padding:11,\n"
        "        titleFont:{family:\"'DM Sans',sans-serif\",size:12,weight:500},\n"
        "        bodyFont:{family:\"'DM Mono',monospace\",size:11},\n"
        '        callbacks:{\n'
        '          title(i)     { const p=PERIODS[i[0].dataIndex]; const ex=excl.includes(p); const po=progOnlyIdx.includes(i[0].dataIndex); return (PERIOD_LABELS[p]||p)+(ex?" \u00b7 \u26a0 excluido del prom.":po?" \u00b7 solo programado":""); },\n'
        '          label(i)     { if(i.raw==null||i.raw===0)return null; const po=progOnlyIdx.includes(i.dataIndex)&&i.datasetIndex===0; return "  "+i.dataset.label+(po?" (prog.)":"")+": "+i.raw.toFixed(1)+"h"; },\n'
        '          afterBody(i) { const p=PERIODS[i[0].dataIndex]; const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex]; if(av==null||my==null||my===0)return[]; const d=my-av; return["  vs prom. cargo: "+(d>=0?"+":"")+d.toFixed(1)+"h"]; }\n'
        '        }\n'
        '      }},\n'
        "      scales:{\n"
        "        x:{grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},\n"
        "        y:{min:0,grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"},callback:function(v){return v+'h';}},border:{display:false}}\n"
        '      }\n'
        '    }\n'
        '  });\n'
        '\n'
        "  const en = document.getElementById('exclNote');\n"
        '  const ep = excl.map(p => PERIOD_LABELS[p]||p).filter(Boolean);\n'
        "  if (ep.length) { en.style.display='flex'; document.getElementById('exclText').textContent='Meses excluidos del promedio comparativo: '+ep.join(', ')+'. Los datos se muestran en el gr\u00e1fico (tri\u00e1ngulo).'; }\n"
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
        "      {label:'Programado',data:prog,backgroundColor:'rgba(103,30,119,.5)',borderColor:'#9B44B8',borderWidth:1,borderRadius:5,borderSkipped:false},\n"
        "      {label:'Efectuado', data:act, backgroundColor:'rgba(38,216,0,.4)',borderColor:'#26D800',borderWidth:1,borderRadius:5,borderSkipped:false}\n"
        '    ]},\n'
        "    options:{responsive:true,maintainAspectRatio:false,\n"
        '      plugins:{legend:{display:false},tooltip:{\n'
        "        backgroundColor:'#18102A',borderColor:'rgba(103,30,119,.5)',borderWidth:1,titleColor:'#F0E8F8',bodyColor:'#8B6FA8',padding:11,\n"
        "        bodyFont:{family:\"'DM Mono',monospace\",size:11},\n"
        '        callbacks:{\n'
        '          label(i){return "  "+i.dataset.label+": "+i.raw.toFixed(1)+"h";},\n'
        '          afterBody(i){const idx=i[0].dataIndex;const d=act[idx]-prog[idx];if(prog[idx]===0&&act[idx]===0)return["  Sin datos"];const w=d>.5?"\u25b2 Efectuado mayor":d<-.5?"\u25b2 Programado mayor":"\u2248 Similares";return["  \u0394: "+(d>=0?"+":"")+d.toFixed(1)+"h  "+w];}\n'
        '        }\n'
        '      }},\n'
        "      scales:{\n"
        "        x:{grid:{display:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},\n"
        "        y:{min:0,grid:{color:'rgba(103,30,119,.25)',drawBorder:false},ticks:{color:'#8B6FA8',font:{size:11,family:\"'DM Mono',monospace\"},callback:function(v){return v+'h';}},border:{display:false}}\n"
        '      }\n'
        '    }\n'
        '  });\n'
        '\n'
        '  // Comparison table\n'
        '  let tbl = \'<table class="comp-table"><thead><tr><th>Per\\u00edodo</th><th>Block prog.</th><th>Block ef.</th><th>\\u0394 Block</th><th>Turnos</th><th>Vuelos prog.</th><th>Vuelos ef.</th><th>Blancos</th></tr></thead><tbody>\';\n'
        '  PERIODS.forEach(p => {\n'
        '    const r = pr.find(x => x.period===p); if(!r) return;\n'
        '    const pg=r.block_h_programmed||0, ac=r.block_h_actual||0, d=ac-pg;\n'
        '    const tp=r.turnos_programados, vp=r.vuelos_programados;\n'
        '    const ve=r.vuelos_efectuados,  bl=r.dias_blancos;\n'
        '    const ex = r.exclude_from_avg ? \'<span style="color:var(--warn);font-size:9px"> \\u2731</span>\' : "";\n'
        '    const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\\u2014";\n'
        '    const blCell = bl !== null ? (bl > 5 ? \'<span style="color:var(--warm-red)">\'+bl+\'</span>\' : bl) : "\\u2014";\n'
        '    tbl += \'<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">\' + (PERIOD_LABELS[p]||p) + ex + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (pg>0?pg.toFixed(1)+"h":"\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ac>0?ac.toFixed(1)+"h":"\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + (d>=0?"var(--teal)":"var(--warm-red)") + \'">\' + dstr + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (tp !== null ? tp : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (vp !== null ? vp : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ve !== null ? ve : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + blCell + \'</td>\'\n'
        '         + \'</tr>\';\n'
        '  });\n'
        '  if (excl.length) tbl += \'<tr><td colspan="8" style="font-size:9px;color:var(--muted);font-family:var(--mono);padding:6px 10px">\\u2731 Excluido del promedio comparativo</td></tr>\';\n'
        '  tbl += "</tbody></table>";\n'
        "  document.getElementById('compTableWrap').innerHTML = tbl;\n"
        '}\n'

        '// Hamburger menu toggle for mobile\n'
        'const menuBtn = document.getElementById("menuBtn");\n'
        'const sidebar  = document.getElementById("sidebar");\n'
        'const overlay  = document.getElementById("overlay");\n'
        'function openMenu()  { sidebar.classList.add("open"); overlay.classList.add("open"); menuBtn.classList.add("open"); document.body.style.overflow="hidden"; }\n'
        'function closeMenu() { sidebar.classList.remove("open"); overlay.classList.remove("open"); menuBtn.classList.remove("open"); document.body.style.overflow=""; }\n'
        'menuBtn.addEventListener("click", () => sidebar.classList.contains("open") ? closeMenu() : openMenu());\n'
        'overlay.addEventListener("click", closeMenu);\n'
        '// Auto-close sidebar when a pilot is selected on mobile\n'
        'document.getElementById("selPilot").addEventListener("change", () => { if(window.innerWidth <= 768) closeMenu(); });\n'
        'document.getElementById("selMonth").addEventListener("change", () => { if(window.innerWidth <= 768) closeMenu(); });\n'
        '\n'
    )

    # HTML body — plain string concatenation throughout
    html = (
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
        '<!-- Hamburger mobile button -->\n'
        '<button class="hamburger" id="menuBtn" aria-label="Abrir menú"><span></span><span></span><span></span></button>\n'
        '<div class="sidebar-overlay" id="overlay"></div>\n'
        '<div class="sidebar" id="sidebar">\n'
        '  <div class="sidebar-top">\n'
        '    <div class="logo-wrap"><img src="data:image/png;base64,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08CmV7fbNZJdegtmTp6PB+ej+MFVssTu1kYfE4LX23lIzy9xL804u08G18wH013sL5W890Emk3HqqccJ6udw3/r+FbntVmO5ue12Y7k+bg70UMcG5w2c6aYOn6SqGj3J6M4nUdxlWWclc4vAMTIhxmGWhl7t7WRs90drpTwHnUazRW190RrNHbD8GdUXl4LmHBsYhEmH4lSVAP5udjuPAngV6iitNdyM013CIiwYCIiAIiIAiIgCKjyQ27bX8Kj9yscteY8mZiOHxYbRyRnY6SCQ+4Yfzg7pbK65WPETZXW7HhEgkUSXc5/NGkAzCcKIO808lvXKvXPZr71YN6PL7Zb3obl+DfyV3glqiiV1z2a+9WDejy+2Trns196cG9Hl9snI3eByV3glqiiX1z2a+9ODejy+2Trns196cG9Hl9snJXeDHJW+CWiKJfXPZr704N6PL7ZOuezX3pwb0eX2ycld4HJW+CWiKJfXPZr704N6PL7ZOuezX3pwb0eX2ycld4HJW+CWiKJfXPZq704P6PL7ZOuezV3pwf0eX2ycld4HJW+CWiKJfXPZq704P6PL7ZV657NPenB/R5fbJyV3gclb4JZoomdc9mnvTg/o8vtk657NPenB/R5fbJyV3gclb4JZoomdc9mnvTg/o8vtlXrns0d6cI9Hk9snJXeBydvglkiib1z2aO9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2aO9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2aO9OEejye2Trnsz96cI9Hk9snJXeByVvglkiib1z2Z+9OEejye2Trnsz96cI9Hk9snJXeByVvglkiib1z2Z+9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2aO9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2aO9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2aO9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2aO9OEejye2Trns0d6cI9Hk9snJXeByVvglkiib1z2Z+9OEejye2Trnsz96cI9Hk9snJXeByVvglkiib1z2Z+9OEejye2VeuezN3pwn0eT2ycld4HJW+CWKKJ3XPZm704T6PJ7ZOuezL3pwr0eT2ycld4HJ2+CWKKJ/XPZj34Vhfo8ntlyHOe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DwKPrmaM4Ok4ae4HZYb1KjHqUVOHSREXvbVbwhRix+nNFik0Tmluho21W2tH818u/mOmUq42/lHzP+XadOpXLujotabOOm63hKoL907zqrtTG69t7oO2Xz1e3JwEHmHEwQ7unedU7LunedciVRecsLqU7LunedOy7p3nVUWOJgp2XdO86dl3TvOqonEwU7LunedOy7p3nVUTiYKdl3TvOnZd07zqqJxMFOy7p3nTsu6d51VE4mCnZd07zp2XdO86qicTBTsu6d507LunedVROJgp2XdO86dl3TvOqonEwU7LunedOy7p3nVUTiYKdl3TvOnZd07zqpVAFni6DBXX3TvOqHS7p3nVURNmehTWfdkeVVFxte4eMoWjYTYLiTpN0n6mhZfVZMS9pycSHDsreF57HyrI3I3Q1rsRdURsmjhOzTBG544W2q2cn5bqsw1ccZhe2nN7ktI3O8BG1qkPgGC0mE0jYYIY2kX1hoG8ncBxXb/xba7HL48uiOx/jG1zb+M+iPTYCBrVURfQz6CEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBeBnHL1JjdA6OWGMv1WJaL7W+A8F76HXtWu2qNsHCS6M121xti4yXRkWMxYJU4NiJglY/Qd2rrHRGoHbYcV5jOyYXXsPCpH51ynSYzSOtEwSatYaL7W/qngo+41hlThNY+nq43MaLaJ0SL6gTtA4hfLd+2SejnxwX0s+Z73ss9HNzrX0s6IN23sR49qDbdcrhwBdqI7UcVxF9S5yDz0ZztUnNYkN6INiLwzzFYCIiwegiIgCIiAIiIAiIgCIiAIiIAiIgCIiAFU2KoVDtXqKy8Du8HK2lYatapJbobiBsVCSG6toXKTU8t3FZjHMkjxGLdvCXzyR4f0zjkb9EWjvc24tf4FIBoDRYLEfIhREMfVOaRpWtcf3gWXV9c/jun+Do1/Z9Z/j+n+DpF/ZaHK9USUuR62aMkOHQ7W/vWBa/JhpyP6I4nTta51alsR5Q6D7JZXqaW19PQ/g9p4HgteuN0r6So6Uka5kke24sdYB/1Xa7Qk7GdjtSzYzql5c50j2XczYLajdexk2ChfmCjixBzdBunpB5bo9o623yLyHEOLWiw26RVGvLXh93tkGxwNlcWfVFpFtPqmjY9lqmwiLD2HDI6YR67GMM4nufKvXUJuT/lyzBlynENaXVMTe66I/e475B3QWVsH5zWAPYPsjhmINJ3xQMtv4zeJc5bo7Yvtk5+zSWxfbJINfGppaapjMdRBFK07nsDh/FYxwvl4yJXWvWOp77ppYW8f7TwK8sGzvlXFwOkcdw2Vx9y2ricd+4OPAqO65x7o0Ouce6LfxvkfyZiuItrpaBsMjb6oYYWtOoDuDwV9UFHBQ07aemYGRtvYAADaTu8a+0ckcgux7XDwG65LzKcpdGzy5N9wuMnaFclxk7QryzyyOvKj+MEvk+gxWgrx5U2kY/LfwfQYrOXxnff/tn/AKfIt5/+yf8AoHbbvKs28mVczDcpsqZ5A0673db+seN5HFYSDQ46JNr716Zx+vOEMoIXyRtF72LgO2vuKkbFro6OyVj74Peya2OjulN98Hp5+zNU5gxJ7IpXinjtdukbaw3wkbWq2pO3A1dD4DaqgNj0gw6Rda5vfYuLO6edSganVz1l0rLGQbdXPW3TnYyus30S0Ddde9k3LtTjWIN0IniEXvdptsd4CNoXQwDCp8Yr2QQMcQ6+sA8CdwPBSJybl2mwajDWxM0+OiL7XeAcVd/x/ZpaufFYvpRc7Ds8tXPisX0o7uXcIgwmibDExrfEAN5PAcV6qIvp9dca4qMV0R9MrrjXFRiuiCIi9nswPzmOTluNYZ9l8Pp2mri3tZrN3RN3MJ2A71EOVro3PaY3hzbdiRr1+BbLMQpIqymdBMxr2utcOAO8Hf4lCbl/yDLlHMZrYWAUNRsLQdFmiyMbmgC5crTQanhfAy00Gow+FmK3EtGk3WPcuGx3G3FdvDK+qwuvjqKRzmzNv90jJGjcEaiCCNRK6jgGDR/qo/vQO++3/sLkXGGSzRpNdtO21leOKfVlzYlOJOXkl5Q6PHcrMqJ52mVl9Lsxve8Da47go884vlDnzHmI0VK89KUvascTou0mRE3AcQdbVj/KmbsTyzTPoqaqndFJbW2R2qxJ3OHdLwqmeapmkqKmRz3vtcuJJ1at6gUaJV28b7FfVolCziPm1ui0vvcv28dSEHSAsSd43jxqj3BrGuGvR3LvYPh9XX4kylpY3TVD76g0uDrAncL7Ap0nwRbLGUlGJd3I1kepzlmOKLoRNGy/RHFpsbsfbXokbWqdmB4ZS4Th8dHSRMjjZewa0DaSdwHEqxuRLIdHlLAI29Bb0y6+k4tGl28lteiDscsjrmtXe7Z/0c3qr3bL+giIopFCIiAIiIAdSjlzrs6dLYf9gaaZwfNt0HbLGF+53+iz9j1fHh2GyVUrmtay2sm21wHHwqA3KhmabM2bZ8Qkkc+mGj0O5JH3tgNtZG1qnaGn4k8+CdoafiTz4LSLiS0vaN9w0K9uRzK8mac7UsEkemxun0U6N2643kbj3O9WTDtOlbstl1KvmmZPkpKKTG62ANklto3Za1jM07W+Eb1a6y34VeC11dnw6j7c5XJcUuVenKakZ0WDYWRi/ZSRDc1ROa0XLTrLdjt3nWx7NmDwY1g0tFMxrmvttAOxwO8HgoBcoWX6jLuPVWEOYWiHQ0XWIvdjXHXYd1wWnbNRlODI+235TgzscmmPyYBmWnrmyPjY3S6IQ617scBvHFT8y9iEWJ4XFVxO0mvvruDscRxPBa3HNIkLRIQ0bdFymLzVs5jHstyYXUTh1TR2uHP1nTfM7VdxOwcFr3Kn/wBo17jT040ZuREVMVAUZOeh97oP3n/QUm1GXnofesP/AHn/AEFJ0n3USNL91EZBuWUubR+UGH4XqpVi0bllHm0flCh+F6qVdDrPsMvtV9lk42dqFjzl8F8i1XwPWxrIbO1Cx9y9fiJV/A9bGuZp96Odp+4iBzHhha7RDrX1EXC7eW2n7JU7ZnEP7K+gdWw8V0Xdq1elgLQcUidruL/RK6e1NVHTWpqo2JZaIODw2JPbaz+0V6S8rKn9CQfC+kV6q5V9zln3CIiwYPjW08VVTuhmY1zHWuCAd996gry/5fpcAz1N0MDoDtGzGWuLRR7rAbSp2zPayMucQAOKhDzlcSpcRzzI2lc2QstpWIO2KLgfArDbm1aWG3Z+KYquAdGY6Tm7dHXt8ak7zN8SqAa7C3EmNnQ7Xvqv0d3H/RRiu14Y9o1m+lq8yktzN4Hy4nilWGkR/cbXH6s4VruaXwiz3Np1klcYF8Ol8nzha7886s3VQtbtPVtWxDGCBh8l/B84WvDPhvnCrts7D1bVA2n3yK/bfczxI9hU5ebJ+SnDP3vr5VBqPYVOXmx/kowz976+VSN29kSTuvsRlBERUJRhERAEREBYvLp+TnEf3Xro1AQdo7yKffLr+TjEf3Xro1AQdo7yK/2f2sutr9rO5hJ/8xaON/mK2HZJ/oCn+F9Ny14YV/SbPL9ErYdkn8X6f4X03LRuvdHjdO6Pbf2hUGucz+UCf4PqolOV/aFQa5zItygT/B9VEo+2/eI2g+4YxeSC4gaXgshZoMkmcSNC1nOP3N99XY8bb/Cjw4lwZ225Z15O8gYTnPk0c2Gnj6aFtF+g3SH3Z19YaTsar6+5VdWXd1qrWWYJZK6PstA6J7bSGo+JZ95t3Ki/Ba9mA45WONO++jLNJqbYSv2ueALkgbFhHHsOrcDxGWixSBzHxW+56BF7gHUHAcQupE51OQ10zw5nbSxO7I34H+C02VR1NbPEq46itmy2CaKeMSRPa9p2FpBC+iwBzduVQYpTMwXF6lhqWX0Xuf21zK7a55J1Abln5pDhcEELm7anXLhZzttbrlwsxFzpHEcndWP2PXRKFB7ZTW50gvye1fwPWxKFJ7ZXW0vEWXe0+1ntZK/GWk+H9By2K0X4Mzy/OtdWSBfM9J8P6DlsWo/wdvl+dRd0eZoibp70dbHP6Ol8nzha38R+/j/vcFsfx3+jZfJ84WuDEPv4/wC9wWzafcxtnuZwg1Tm5A8fiWwnkycDlamtc9tt/bete0ABqCDa3h8SnzybY7gzcs08bsRoY3DSuDMwe7d4V73VN4N26JtIvhF0BjWDnZitCf8A7DP5qv2YwnvnR/Lt/mqXDKbDO8i6P2YwnvnRfLt/mgxjCjsxKj+Xb/NMMYZ3kXR+zGE986L5dv8ANdimqqWpF6eohmHGN4d8yYYwz6vaHNLTvUeuVvDulswOLWNAk8HBjPApDLEvLjh4EUdcG9pe9hx6GOCov5DR8bRS/oo/5BR8XRS/ow8bl7uAQCxRpGm/w2VV8hfRYPlLXDDBxO1VRyLDWDC7BEReQEREAREQBERAEREAREQBERAEREAREQBERZM/ge6tcBVA4kKj26QsNRTVI27TbRXtLi6I8Ri5Mo/smHd417eVcDqcbrmwsgk6EL3JYbbCeB3hdTBcKq8aqGxU8LwNdyGngeAPBSHyZlmlwWkA6DH0Tjoi+13gHFdPsWyS1VnFNfSjptk2WWqnxT9qPvlTL1Lg1I1kcTA4X16I4nwDiveRF9OqqjVBQiuiPpVVUaoKEV0QREWw2BERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBWRyi5Qp8YozLDBH0YbwwcW8Gk7Ar3VHAOFiLhaNRp4aitwmujNOoohfW4TXRkTK+jqaKpMNVE9j2bLtIvcX3+NfFutt9IDwErOnKVkqPEKR9VRxME4tazde1o3NvsBWD6ylkpap1PM1zHttqItuvvXyjetplt9mf/LPlu8bZLbrG/8Ayz5gdjfZbcd6oqaRcdeojcqgBUWMdylxwrr+StlQqpIVFgRTbGtEOtEePwenj8BERYPIREQBERAEREAREQBERAEREASyoVyHFek8djD6PKKAFwsN/FVF3NYbXOtGFdvBYDVVkNO1pJdpbvASpGmg53xS8kiiDlfHBn3kvw7pXA4nFoBN936z/AryXmZag6Xw1kdrWv8ASK9NfatJWq6Yx/o+x6WCrpjH+jhURNmiMbwCDxUPecTya4jhuIzYphtCZIH6N+hROJFhE33LANpO9TFXQxrCaHF6R1NWwMkY619JjTvB3g8ArDT3yonxIn6fUSonxRNbDg27y54Y3Vck20PHwuuT7uAtGS3uravOpOcpXN3Ek763LXQeyteCXtHamjW1kWv3RWD8ycnmcMEmLKnBMSMfGKll6HsHFo4+dX2n1dc+uS7p1dc+paJIc4A627+CaTtjSCF956WohdoyU00Z/WjIaPGvi4C2z4incUZIm/Fg0VJcG2DADxA1ru0OMYtRuD6HEK6jcN8Ezoz/AAPhPnXngNv204du0jqXJxe0jQcHcbG68KEZdGa8Rn0Mv8n3LnmfBpGxYhUvrYBe/R3ySSHttl5ANpHkClHyc8omA5yohJR1kLZxthdIwP2u9yHE7GkrX8QHSAtdoHjewVwZBzViOWMZNfhVRUNMXbxOe4NN2uA1NcO6JVfrdvjjMSDq9BHGYmxRUf2hVp8mebKbNWAxVsc0bnm+kGuBt2Tx3R7lXXJ2hVBJOLwyjknHoyPXKv8A09J5PoMVlq8+VVwdj8tvB9BisxfGd9/+6f8Ap8h3n/7J/wClDfXYi65Odd/YgBv8VQDWqkC9wqqMuHsU6ajNv8nEjRB0LniuxSU7qxwiYNZ//wB/0XwuRs2r60lTLQyh8VnO4az/AN7Vtqgm1KXbJvrgsKUu2TPXJplKDC6RtRNHG6U3sdEG2tw3tHFX4NWxY/5N83QYjTimmkY2UbAXAb3He7wLIAIIuCvse1Oh6aPwex9d2p0vTR+D2CIisixCIiAKyuVjJFDnLLU1DURNMp0dB+i247NhOstPcq9UIBFisxk4vKMxk4vKNbmYsIq8GxKSixKB8UkdrdgQ03AOrSA4heXpPaOhEXvsJUpec7ycPqmfZzDaZukztmsZtv0Jg7Vnj3qLrtK5a9ha9u24sum0NyvjiR0uhuVy6lAA1hZt4krnG3otT0Nx0W90dTdnFcZLAA31FUnDgOgm7Xd0NXhUzKk+F9kSYtOTQgDXkE6mO2aW6yktzXeTR72fbDjNKLD7wJY+y/rWOvpM8WwrE3I9kSqzjmVsAhf0hHe79E2N2PO3RI7ZqnVgmGU2E4fHR0sTI42XsGtA2kncBxVLuGqa+hFRr9Q19CO5GxrG6LQAOAC5IipinCIiAIiIAiL5VUzYITI9wAG8nwoDC/Olzi/Asutw+lnDZ6m9g19iNF8TtzgdhUNpSTAWxl5YO10u22rJHL7m5+Z86TuExdTwaPQg1129lHHf3RG1u5Y2c7oTWC19t10u30/Dryzo9BSq68s9/JuCy5hzDTUNNC94fpXAaTazHHXYHuVP7JuDw4LgkNHDG1gbpag0Da5x4Dio180jJs8tW/HayAdDFtAvYeEzDa7fFvUrgLCyqtwuc58PgrNfc5z4fAIuLKL3OyyY+EszFTQ9gL9GLW7fvLG3s35ypQq2eU3L8WZcoVmGPiY90mholzQbWkY47j3PBRtNb8KxSI2ntddiZrvAJY0nTDj21/4LInIRmx+V87UlQ2Uso6jT6O0OtbRikDdIaQG12q6sjGKCrwvFKikrY3MfHo6QLSNrQRtA4hdakf0ERTB7mPkvpWNi22ryXXS2qN1eEdHao214NllHUR1MDZonBzTsIN99l9linm55vOZMpxtnkLp476Qc651yS22uJ2NWVly1kHCTizmLIOEnFhRm56OqLDvD0X/oKTKjPz07dBwz97/0Fu0n3UbdL91EYhuWUebR+UKH4XqpVi4bllLm0flBh+F6qVdDrPsMvtV9lk4mdqFj7l6/ESr+B62NZBZ2oWPuXn8Rav4HrY1zNXvRztXvRAt5swEr08vlgxdjHva21+yJsO1K82wIaDs1rlE4tkiAcQ831g6wusw51YZ1PWVeGbEMrYzhAwaFn2UorjSuOmGd0fCvV+zOEd9aH0hn81rxizRmiKNrmY9iEbXXs2OrkGzwXXL7bc078wYt6ZL9ZUj21t9GUr29t9GbDPszg97fZag9IZ/NcJ8dwSnZpz4vh8beL6lgH8Sters2Zmtd2YMZA4trJL/SXGpzNmaoj6HPmDFHxfr1khP8Snpc/IW2TZLLld5Z8CwvD5KPBcQgrqp1rGmmZINrD7mQHYT5iogV9fV4hUzYlUPM9S/RuLlxFho7zfYF15pS+QEztN9sjn9mPKuMYLpToSNjHAOsSrHS6WFPctNNpYUx/srAwlga1ti7aCNYU0+bFlabAMo9HqYmslqNvYkHsZJRruBuKwLyEcl+I5lxWOuxOinjo23++xOAd2Mg90wg62hTRw2jhoaRlNAxrGNvYNAG0k7vGoG5ahSfBErNwvUnwo+eOf0ZL5PpBa8M7m+bKk/seratiGNi+GS/B+kFruzr+NdT8D1bU2n3SG2+5njx7Cpzc2T8k+GfvfXyqDcG1Tk5sv5KcM/e+vlW/dvZEkbr7UZPREVCUYREQBERAWLy6fk4xH9166NQEHaO8in3y5/k5xL9166NQEHaO8iv9n9rLra/azt4V/SbPL9ErYdkj8X6f4X03LXjhX9Js8v0Sth+SPxfp/hfTctG690eN07o9p/aFQb5zYtygz/B9VEpyP7UqDnOd/KDN8H1USj7b94jaD7hjBxIcSNRUt+aJDGcpuFrgW1fvJlEdwu5wUu+aG0tys/yfTmVlub/AObLDcX/AMsHy5xHJPDi0TsdwmnZ0022k0MHZfemDU1hJ1A71EyaCWjmkp52Oa6K338EE3167+NbLKqCOphMUrGuadoIvvUSectyY/YmpOPYfTvNKfvjImbNUTBqDANpO9RNu1jjLgZG2/VOL4GYTwPE6/BK2GpoKowSM0rOdI5pFwR2RBHE2U4eRDP1FnDLkZNQBWMv0SOR40xd8ltWkTsbdQQDmvuJQbDtjbtuFuKu/kxzliOTMwtrop5OhC/Ro2vdov7B4bYaQvYuvrUzX6aNseKJL12mVi4okpOdQ8N5Pqjbr0fWwqFjtRUtOXPMNDmfkpfWUM7JNK1wHtNvu8Y3E9yVEt5u7YAtW2RcYvJjbE4xeT2sj/jNSfD+g5bFaL8GZ5fnWuvIv4zUvw/oOWxSi/BmeX51H3VJTWCJunvR1cf/AKNl8n0gtb9d+En/AL3LZDj39Gy+T5wtb1d+En/vcvW0+6Q2z3M+LgSXgFw2a27V6sOYceiAbFj2JRMHuY6t7f8AVeXcBziXNbs1uNguTI5Wa5KeVzTvYwlXk4KUepeWcPDhnrxZszS0OtmDF9HujWS6v+JcvtvzTs+2XFfTpfrLyG0s99JjKgtPuXg2K5Cmqfew+IVprph/Rqgqkux6v235p/SXFfTpPrKozfmkbMzYr6dL9ZeR0Gf8w34hToM/5hvxCscNSf1I8L4WeqPWObc1PcAMy4ufA2ul1/8AEpRc0/FcUxLBpnYlV1NQ4aNnSyOf7qbuieAUSI4J9IXgfb+yYdLyKWHNEbI3AZhJBURHsfvrLe6mVduPw+D6UQde4OH0oz+rP5UsKdiOXZwy2l2O39tng8CvBdPGIBUYfJERcG3zhc5qalbVKD/KOe1NStqlB/lEUA3sn22stfw3RdzGKc0de+nIIdq0vMD/AKrp718R1NbrtlB/hnxi6DjbOD/DDkQotDeTw10QREWDAREQBERAEREAREQBERAEREAREQBERAUOrWVXWH6NlR4BaQVWQkzAMFxxXuPVYH5AftA2hdvDcNqK2qbBSRvk0r3LGkjYTuHgK+dHSy1lR0KlifJIdmi0nd4PEs78m2S4MKoGVFVEDO69w5ouNbhvaDsKvdl2ezW3JrpFFzs+02625NdI/k7OQspQ4TStklhZ0TXtaOLv1RxV6IAALAWCL6tp9PDT1qEF0PqWn08NPWoQXQIiLebwiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAo9jXtLXAEHiFinlOyOZw6tw+Fmnqvot1+4G5vjWV1xljZKwse0OB3EXULXaGvW1OuxEPXaKvWVOuxESJoZIptCWNzHt2gi17ri0gk69E8Csq8qOSJg412GRDSG1rWn9Qe5b41iuaMCUscCyVu0bF8o3Pa7NJc4tdPwfLNz2yzTXcMu34OOu+sEeNVGtNIym5GjbwKosqmT/BWuSXQKiEovJ5SwERFgyEREAREQBERAEREAREQBERAEJ1puQ7V6ishdXgAEjUrr5M6Lp3Go5mt7Ft7gji14/wBFagNhdZV5FcOvJJOGHQ1WJH94OCvf49p/jatLwXWwU/H1SXgzDSs6HCGr6oBYWRfXUsLB9XSwsBEVt59zTQ5YweSsqaiKNzbWD3tBPZNG8julkyXJtXVrMOoK1ujV0VNO3hJE1w/iPAFgnIXOFwyurHUmNGOnGrRkOi0HU4nW6XwALNeDZiwXF4hJQYlSTg7mTscd/AngV6cZR7rB6cZR7rB5GIcnOSq1pEuWsJF9pbQw33fqeBW5WchuQ6i9sPEN/wA1DA3/AKayeCDsIRelbOPZmVbNdmR6zFzbsLc4zYVVSAjY2WRttw2Ni8ajrnvJWK5Mxd1NiETy09q4NdonsWk2JaO6C2GTSMjbpPc1o8Jsorc6vH8Fr62no6ToMk7NLSezQJFxCRrBvsBU7Sam1zw+pN0t9kp4I6ANdpNc6wO9p2Lm6SXSFmMY89uYwQDwXziAki0+1B3L6lx6IWgX8NlfybcOpeybcOpJPmf4xKZKjDHyve1mja7iRr6M7j/opPS9oVEfmf2GZ60a7/c/oTKXEvaLmNasWs5vV/cZHTlQJ+z8o8X0GK0Vd3Kh+MEvk+gxWiviG+//AHT/ANPjO9f/AGT/ANF7XXc+xtX0m2rDCYnX12OrXbhxXSIBBusyZBwanxjKTIJIQduvRH5xx4HgvW07bz0nFd8Da9uWtnKK74MNtOmwPAIB2E71zbomcPYHlo3EeBe7nLLtVgNUIpIXtp29q7RNtYbfXYDaV4NyNcRB4qDq9NbpZOqaK7W0XaVuqaO5hGI1GG1ZqoJZI3N2DSIGwjcRxWe8h5rixalDJX6Mu9riAdrv1juCjs4aZuXADeLr1sv45V4TVipbI4O/Nhx16iNlxxV1sW8S0U1CT+ll3se7S0U4wk/pZKZFb2TcxU2NUDHtkb0TXdukL7XeE8FcK+p02xugpwfRn0+q2NsFOL6MIiLYbAiIgPOzHQwYhhcsFRG17DbU5oPugd48CgBylYfBheca2jhbotb0OwAAH3tp4DithNf+CP8AJ84UAuWg35Q8R/depYrXam/iYLTa2/iNFmgXAa/XbbZfRoMxAkPZneFwPbFfSn+/t8vzK84Wm2XOGm2TF5qmH0MWTmVMUDBK+93ljb6pJhtss2LDnNX/ABEh+F62ZZjXK6l5tZzOpebWERFoNAREQBERAFi/nB5uOWsqvdBNozvtohrrHVJHwcDscsnSvDGFziABxUL+crnCXF8zOo4ZxJBTWuwPJB0mRHZpEbQpWkqdliRJ0lXxLEYfqHumd0WR7pHv92DcavCqN0C92mNRta+5cW2gD4QdJotoHb4SgbpbSuneIw4UdM8Rhwoknya8uGUMqZcgw12G4iZI9LSdFBFY3e5w19EHdK6hzm8o78NxfyQRe2UQdHROt58V0uO6PnUBbbCbzIr/AE6M5ZZL485vKPe3F/kIvbIecxlFzCDhuL/IRe1UQbjiU09dta8z2uuLPEtthFl58quYcGzJmKTEsJppoGvtptkjY29mMaNTSe5O9WbKNN4IFm7wN3iVQG6OlrugVhXXGtKCLKqCguEy7za84S4HmplDUTaNPUX1BxAGiyU67uA2lTVie2Rge0gg8FrVw2sqKCtZVUshjlZfRIJG0EbvGp6ci2Z48z5Mp6zowfKNLTGlcj7pIBfWdzVR7lRwy4kUu5U8MuJF8KMvPRv0PDv3v/QUmlGXnoG8WH6tnRf+gomk+6iHpfuojINyylzZ/wAoUPwvVSrFo3LKXNnH/qFD8L1Uq6HWfYZfar7LJxM7ULHvL1f7Raq36nrY1kJnahY95ejbItV8D1sa5mn7iOdp96IHW0tEbNqpAdK87Y5H9w1ou7gdSaWiGki69DLrWuxeBmjdo0tRGrtSuu4lGo6riUaz7x5VzGYWD7CYrKWXs6OlkLTc8dFcxlPMpdb7A4v4+lJLfRU98sYFg5waEnDKNxOlrMDD7o+BemcCwYtsMKoR4qdn8lReouMuiKT1Bxl0RrgqaaamldDUxywSNt2EoLSd+wr5kBzrNL328oUpucHyN09U047gVPoTM2xRsABv0Nmxkd9lztUWZYpoZpKbRfFM212vBadYv49itqNbG6PQtadarY9D28n5enzTiTKGkfQ073Xt0cloOpx3A9yVJLk85uWF0b2VuYJ21Uov2Eb2vj90Nj4vC3yhRewbEazDK6Gpw6cxSs0ru0yALgjaCNxKmdyFcqVBm/DW01ROyOsZe7XvaCbukOwvJ2NUHcHbHrHsQtdO1dY9jJ2F4Xh+F04goKOnpoxsbFE1g2k7gOJXcRFRN5KRvJ0sbP8A5ZL5PpBa8c8ADNtSB+p6tq2G47/Rsp8X0gteOeDfNtSf2PVtVttXuZZ7b7meRBtU5ObJ+SjDP3vr5VBqO9jbapyc2T8lGGfvfXyqRu3siSd19qMoIiKhKMIiIAiIgLH5cvycYl+69dGoBDtHeRT+5cNfJziP7r10agAO0d5FfbP7WXW19mdzCR/5kzy/RK2H5J/F+D4X03LXlhH9IN8vzFbD8mC2AU/wvpuWnde6PG6d0ew/tSoO8578oU3wfVRKcT+1Kg9zoPygy+T1USj7b95EbQfcMWyX7K23Upec0K5ypIT4PWTKIcm13kUvuaJ+KL/J6yZWG6e1lhuP2zOy8zMmCUOO4ZJQ19PFNG+1w9jXbHA7weAXpoqBNp5RRJtPKIFcsWQqzJ+NyNEF6M26E5rDojsWX0jogbXarKwna5XRvu1rdhOo+VT/AOU/I+G5uwWSmngZ0U2s8MbftmHaWnuVBnOGW8Ry1jMmHYlDM2SO3bNcOiXa12rSAvbSHiV/odUrFwy7l9otUprhl3PnHmTGBghws1Mop5PvsL3v7CztIWbewudesLxyLlG6Ti+R7h0R1tMA6vAhVi4qHYntKPY9vI34zUvw/oOWxWk/B2+X51rryMP9p6T4f0HLYpSfg7fL86o909yKXc19SOrj39Gy+T5wtcGIffx/3uC2QY3/AEdL5PnC1v4j9/H/AHuC97T7pDbPcz5xBpmIe0ObvBFxsU5OT3IGUKnLkEtVlvCJnHSuX0MTj27uLVBuL78VsM5NT/svT/C+m5bN0nKOMM3bpOSwkcDyeZH1D7VMDAH/AO3Q/VVep5kj9FME/wAPh+qroRU3xJ+Sn45eS1Tyc5GP+6eB/wCHQ/UVRydZGH+6WBf4dD9VXSifEl5HHLyWwOT7I7TduUsCB8GHQ/VXsYTg2FYSwswzD6SjadoghbGN/cgcT5130Xlyb7sw5N92FxkbpMLTvXJFgwRw5TqDpTMM7i3R0tHRsNXaMvu8KtM9ssq8t9BozR1DW2ve5t4IxwWKnbQvkf8AIdOqdc0vyfKN9oVWuaX5FkVRrVFz77lJnq0ERFgBERAEREAREQBERAEREAREQBERACbC6DX4PGqtGkbKjbSXGyyyA7RHbEW8C+lNFNLJ0KKJzpHbG6JJ8y4xxiQ9Da1zyeAustclmRnOH2SxaIiT3LS39tuxzfErbatqs3CxRj2LLbdss3CxQj2PS5LsmRUcYrK2Fpl3BzRq7cb2+ELJoAAsAAFwhijiboxtDR4BZc19b0WjhpKlXA+q6LSQ0lSrgERFLJYREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQHyqqeOoiMcjWuB4i6wfyj5Gnoqg11JFpN3hjSdzRub4Ss6rrYlRQ1tM6GVjXA8QOIP8Aoq7cdur1tfDLv+Cv3Hb69ZXwy7/gia0nWCw3G3Vs8aHUr55QMl1WD1MlTSxOdTutpBrSdzQNjQNpVitdpu0XDRcNx1L5JuGgs0drjNHyfcNDPRWtWIatiqqAEOtuXKygN5I3EmURVslljAyiiKtksmBlFEVbJZMDKKIq2SyYGUURVslkwMooirZLJgZRRFWyWTAyiiWvdLLkxZTwuhrssVayu5QC5AUgeSTDzS4FG57QHG+79Z/gWB8NgdU1jImi5N/mKk3lemFNhrIw2wF93hK7r+HaZOUrTt/4hp027cdT1kRF9BO+OtidbBQUj6md7WMba5JA2kDf41Cvl45RanNeLyUdJNKKOK3aOOibtjOuziNrSpJ8vlFmWtyq6PL7n6erSDDJc9nH3HiKhDiVDX0lVJT1UFXTzG2kZ2OYDqB369hU7QVRnZmX4JmirjOeZfg6TSWxRNi0Xtjv90Gt2vif4L38vZxzLgLg/D8axAEbGuqpNHfuDhxK8G8dmthcAG9sGnUeGxUXR2RhJYayX9kYS6NZMx4RziM70kYbOygnI3vEzuPGXwr2ouc1mdsXZ4dhxk8EMlvXLASa1Eejof8A5I3J0v8ABmLMHODznisYhDKGljPbOhEzHbjt6KeH8VibEMSxDEagVVdPJUVDu3Je5w2AC1zfYAusGhx0XbCqscXXawDVvWyrTVwf0o3Vaeuv2oo5jjotY1wA3AI9w0nOjc0tba5B1KjA9xLRO1rjv07K+eSvIWI5zxKGGmw+ojoW6XRJXQuAPYvI1hpG1tlu1FyhA9X3qMTOPNMyrPR0kmL1EZaZrWu0g6jM3e3w8VImY2jK8rKeBUuA4VHRU0bWNZfYANridwHFerP96K5TUWuybkczdZ8STkRy5TnH7YZBq12+gxWorq5TPxid5PoMVrEa18W3hxe4WcXbJ8c3axeozUuxxcNR2rPfI1MBgLI7C2u3xnrAxbdl1njkbi/8kafH9J6uP4jJc21Es/4lbGWslwFwZ4y3T4/hb4XRs6Lq0XEDumngeCjvjGFVGEV0sM8T2AWsC0i+oHVcDipVKx+UPJ0GMUrpoYwJxsIaOLRuaTsC6jf9mWur44e5HVb7s8dbXxxX1Ij4xgs4uJBO7eFUWd91d99Gzuf+7L7V1FUUNQYqsObKdxBG7w+AhfA6ti+WWwlXL4b7o+ZXVyjJwa7FxZLzHVYFWsd0R7ozfSbcnc628cVIPLuN0uL0jZYntub6rjifCeCi0HWNzqV35CzZU4LWNp55HOiN9ZJIGpx3uHFdZ/Hd8enaptfQ6r+PbxKhqmz2kjEXSwnEafEKZssErHg8HA7zwPgXdX0mMlJZR9DjJSWUERF6PR8K/wDBH+T5woA8s/5Q8R/depYp/Yh+Bv8AJ84UAeWf8oWI/uvUsVptf3Sz2vPxS0D2xXKnP3Znl+ZcHdsVyg+/N/73LoG85L59ckzuaqb5Ei8vrZlmVYY5qn4iReX1syzOuS1KxbI5XU/dYREWg0BERAERCbC5QFl8sOZocsZPqa2SXQf2OhZwBP3RgNtY7pQHxPEJsTqnYlUF507aQdfSNgG67k8Fnbnc5wfiWLwZeo6odCg0ujBkmo6TYXtvZ3gO0LAT3NkkYLBsD7+LV/Daug23TuMOMv8AbaowhxPuGxOe4RRNdK5uzRGkXfzWRsI5Fc84lRtqYKSNjHXsHxzA7SN0Z4Lr8h2VJczZzhp2xvdTRaWm8NJGuOQi50SNrVOzDqSKjpWwxMDWtvsAG8n/AFXnXax1yxE163VyrniJCZvIJn15/BoR+1HP7NfZvN7z8fcUI8Yn9kpuIoT3O5vJFe53ZyQlHN5z+NdsP/z/AGSHm95/7nD/ADT+yU2kXl7jc+5iW5XS7kJW837P97FlCB4BP7JWNnbJ+M5RrhS4tEGl3aua14adTTqLmjugtidgsJc6PJQxrKT8Ro6drqumtbRZcnSkibuaTsBW2jX2OxcRsp19krFxEOG2c1rg4EG9taz1zT84SYdjM2CVFQBDNo6DS/U2zZnHUXWGs8FgWUPBLWRhkUWw6NjrXpZbxOXAscjraed7JG37V5G1pG4jirbV0q2stNVSrazZCwhzQQbqM/PPsY8Pt/a3/wAhZ2yBj1Pj2AQ1cMrXl2lezgfduG4ngsGc81o6WoTv+6fPAuf0sWrkmUOmi1ckRgA2LKnNkF+UKHX3XqpVise5WU+bMbcoMJ/a9VKui1sf+DL7Vr/iycLdix7y9gnItV8D1sayE3tVj7l6Nsi1XwPWxrmKfuI5un7iIFyA20Ra44r0sDB+yrHMNjrt8Urzn9u5engBAxSMX1m9vildXdH/AInTWx/5GxDLOj9hoNG9uy2/tFekvKyp/QkHwvpFequQl3OWfc+dTBFURGKZjXsO0OAI/iosc4jkjNE/7OYDSPcT98ZBHrP3pg1MZ4TvUq116+ipq6Aw1MMcrDue0Hfff4ltoulTLiRtpudUso1olpbePRe1p7cWs9vDxL2sp5lxTLGKx1+HPbA9l9JrS5oddpGuxF+2Kydy/clFZl3FJMUwmmc6iltpNjYTo2bG0dqwAa3HesNSBriQTr3hdJC5amrDOjqsjqKupPPkm5SMNzrhLJWvbHUa9Jji0e6eBYaZOxqyAteGQM44pk/E456OeYxa7xh7rHsXDYHDuipx8m+ccPzZg0dVS1ETnm+kwPaSOycBqDj3KodXpXU8rsUWr0rqeV2Pdx4gYXKT4PpBa8M4nSzPUvGzsPoNWw/HgDhktxfZ9ILXjnNobmSqA/U+g1Stq9zJG2e5njaYjZpEOI/V2qYHN/z3lTBcgUeGYljuHU88WncSVcTdssjvdOB2EblEAXs2waTr7bYq6Whr6YN+5Y/UrTWaZXJJlnq9P8ZYZsCbyo5DLrfbRg4/+/B9dV6p2RP0pwb/ABCH661+mSQjSBcB/FU6JL3b/OVX+lf2QVtWfybA+qdkX9KMH9Ph+unVOyL+lGD+nw/XWvzosvdv85Tosvdv85T0r+zPpP8AZsFZymZEfq+2vBGng7EYAfpq5cLxGjxOkbVUNRFPC69nRvDgbEjaCRtBWtUPe49k+QeFp1qevIOScgUVw/3fbbfvsiiavR/Ainkh6vR/AWcn35cL9TrErD8166NQCHaO8i2Actv5OcS12+9eujWv4do7yKw2ZZUibtSymd7B/wAPH/e4rYhk7+gYPhfScteOCi+IN/73FbD8o/0HB8L6TlH3T3I1bp3R6ztig/zoR/6gyeT1USnA/U0qD/OgdflAk8FvVRKPt2fjdCPt/wB0xZJtcPEpfc0X8U3+T1kyiDL27vIpf80Uf7JOPi9ZMrLc/t9Sy3Jr4fQzqiIueOfCw/zguTKPNWGjEsMhgZiUGxxbbS0nRt16LC49i0rMC4yMbIwteAQdxC912OuSkj3XY65KSNaddQ1FBO+Cohmilitph7SCbi42+NfB2ogE2B37lJrnI8lE0jXY5gVM0ObbojI2HsvvTBqYzXv3qMznMbdrmOc0dtquW8F02n1Mbops6TT3wuimz2siG+Z6XSBZ29tLVfsHbFsVpPwdvl+da68kRkZmonSu0g7onQ9E32MddbFKT8Hb5fnVXuvuiVu6+9HXxsf+XS+T5wtb2I/fx/3uC2QY4bYfL5PnC1wYl+Ef98Atm0LMpDbFmTPjGWtlJc4NHElTHyLyy5HocBhp6nEBE9ulcPmhb7px3yeFQ3aCZHANY7wPFwknZH7oyOw7kKw1OlV/uLLVaZXYUidA5cMgkXGKR249MQe0Tq4ZA76xekQe0UFt3Y6AbvCpq4MUT0uvyQ/TIeSdY5cMgHZisfpEHtFyHLbkLvpF6RB7RQTBG7QVdI8Wry9rh5C2yDfcnX1bMhXt9lYvSIPaL1Ms8p2VcxYp9jsMrWyzHZaWIg9iXe5eTsaVADSHutY/V2rL/Nes7lIj7GQNF7Bw/sZV4v26FVbkmatToIVRymTVRBsRUxUlj8rOGdPYI9zWguFt36zPB4FH2TVM6Ox7HeVKbM8Imwt7S2+zd+sFGPGYugTygtIdq3eAL5//ADDTpTjcl1OC/ltChONyXU6l7AHii5PA03jcLWVLLg289TiITjLLKIq2SyYPeUURVslkwMooirZLJgZRRFWyWTAyiiKtksmBlFEVbJZMDKKIq2SyYGUUOobCfEm65VdmtUsSbnUFlY/J5U4qWGGkjWBchco2E6mg3O/cuIcC4NZt4q/uTnJdXilY2qrIXMpRezXNIB1OGwtI2gKft+326y1QguhN0W33au9Rguh6HJlkmWqLKutgDYxe2mw3Pbje3xLNMETIYxHG1rQNwFlwoaWKkgbDCxrWjgAN9/8AVfdfW9t26vQ1KEEfVtu2+vQ1KEAiIrEsAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIDp4th1PiNM6CeNjweLQd4O8eBYVzlycYhBUuqMPia9mrUxridjRuZ41nZUc1rhZzQfGFXa/bKNdHFi6lfrttp1scWLqRcZlrG+yacNqg7+4f/JV+1jHe9tX8g/+Sk30lSXv0vFf9gLl0pS+94viBc/8oUP/ANFD8pUPvIjF9rOOd7qr5F/8lT7Wsb73VXyL/wCSk70nSe9ofiBOkqT3rD8mFj5Po/Yx8o0fsRi+1rG+91V8i/8Akn2tY33uqvkX/wAlJ3pKk96wfJhOkqT3rB8mE+T6P2HyjR+xGL7Wsb731XyL/wCSfa1jfe6q+Rf/ACUnekqP3rB8mE6SpPesHyYT5Po/YfKNH7EYvtaxvvdVfIv/AJJ9rWN97qr5F/8AJSd6SpPesHyYTpKk96wfJhPk+j9h8o0fsRi+1rG+91V8i/8Akq/a1jne6q+Qf/JSc6TpPesPyYTpOk97Q/JhPk+j9h8o0fsRj+1nHO91X8g/+Sfazjne6q+Qf/JSc6Tpfe0PxAnSdL72h+IE+T6P2HyjR+xGP7Wcc73VXyD/AOSr9rGO97av5B/8lJzpSl97Q/ECdKUvveH4gWfk+j9jPyjR+xGP7WMd3YZVn9w/+Sp9rGPX/out9Hf/ACUnRS0w2U8XxAq9LU/5iL4gXpfxDT4xxGV/E9OvyYCyNlLFzjsD6mgnjjGlpGSF4Hau4tWfqeMRxBoFlVsUTO1jY3xNAXNX23bbXoIcEC92/b69FBxgERFYlgcZGMkbova1zTuIurAzxyTZVzMxzpKGKnmNuzhiiYfc79AnY1ZBRZjJxeUZTaeUQ2zjze8x4TPKcGbHVwG1tUj37G9zEBtJ8yxniuS804ZIWVWA4k22/pSUDdxaOK2KEA7QD410qvBsJq/wrDKKf+8ga75x4FZU7nZWsNZLCncbK1h9TXBNRVcLtGWmmafDGQuHS835mT4pWw6XIuS5TeTKeAvPF2HQn/lXH7Qcj/odl7/DIfqqR6v/APkkrdv/AMmvWOkqnuDY6aVzuAYSvZwXJeasTm6FSZfxR2l7sUcpbv3hp4FTzjyJkqN2kzKWAsI3tw6Ef8q9WiwjCqL8Ew2jp/7qBrfmHhK1WbpKXZGuzc3LsiK/J3zdsTnnZUZgc2OPXdgLg7Y4bHxfsqS+UMqYNliiFLhdHDC3eWxMaTrcdei0d0V7wAGwAIoF2onb7mV9t87e4XGUXYRtXJFoNJgHlJwLFpsddLBQVMrHW1thcfct4BWyMs48QCMLrTf/APHf/JSfkpqeQ3fBG48SwFUFLTDZBF8QLmtV/GqNTa7ZPqzndT/HaNRa7JPqyMT8s4+Gf0VXD/67/wCSzjyW4fVUGCMiqoXRv16i0j3T+I8Ku001ORYwRHxsC+jWtaLNaGjwBb9s2Gnb7XZB9zdt2xU6Cx2QfVlUIBFjsRFfF4Yu5U8kurW9PYdTh0zfcsZrPaDc2+wFYudlfHmkg4XWejv/AJKUJAcLEA+NfJ1LTO2wRH4AXOa/+N6fV3fFzhnPa7+O0au34mcMjF9q+Nk2OHVOv+xf/Jcm5YzAZNI4XV6O4inff5lJrpKk97Q/EC5CmpwLCCP4gUCP8PoUuLjIEP4jVGXEpsxTyWzZgw+TpSvoa3oZ2OdFJYdudp8YWWmm7b2suLYYm9rEweJoXNdPo9M9NXwZydNo9M9NX8PiyERFLJR8a1pdSvaASTbZ41CXlbyJm6sztV1VHl3FaiKXQs6KileNUbBtDeIKnAvmaeAm5hjJ/ZC30XumWUb6L3TLKNfJ5Oc7Ek/azi/oE31F9G8m+d4pGuOWMYcP1aCY/wDItgHSlN+Yi+IFUU8A/qY/ihSluNieSUtxsTyYw5umBYngeUI6XEqSamlF7tkjcw/fJTscBxCyoqNa1os0AeIKqgzm5ycmQbJucnJhEReDwEREAXl5mqaimwuV9NBNNILWbEwuPbDgvUQgEWIB8ayujMogXmnJ+f8AHsw1WIzZZxt75NC5NBOdjA3Vdp4LyY+TXPUjBGcrY0xg7UnD5gR/wLYMIIBshjHiaFXoMNrdCj+KFYQ3GcI8KRMjrZR6JGH+bZyfzZUwJ82J0rWVktrkxkOFnSj3TQdjgsxqjWhos0ADwBVUGybsk5Mi2TdknJhEReDwEREAXSxuhixHDpaWVgc19rggHYQeHgXdRE8BPBBPPPJRm2gxqoZQ4Hi1VC7R6H0CkleztW3vZltpPmXhO5OM69Ea52Vcc0he98Pmt9BbB3RRuN3RsJ8IXE08B2wx/FCsFuM0sYLBbhNLGDBPNhgzTheHuoccwvEqdgtoGenkZvlJ7a3ELqc6/AMcx6OkbhOFVtX0PT0ug0737eg9yDwPmUg2Rxs7SNrfELKj4on9vGx3jaCo3x38TjwRvjv4nHg1+P5N872a5uVsZsd32Pm+osj837I2asJzoyqxHA8QpYm37KWklYPvcg2uaN5HnUu+gQ/mY/ihchHGDcMaD4ApE9fOceFm+evnOPCyrO1F1ZHLLhtXimUKimo4JJpXaNmsYXH74w7geBV8Kjmhws4AjwhQoS4WmQ4y4Wma95OTnO4kI+1fGDf/APAm+ovQwTk4zqzGYHSZbxZjRpXc6hmAHYn9RT0NNTnbDH8UKvS8F79Bjv8AshT5bjNrGCc9wm1jB0stxSQ4RDHIxzXDSuCLe6K9FBq2Iq9vJAbyERFgweXmXBaLHMOfR1kLJGut2zQd4O8HgFDnlQ5GcyYRjj3YPhNVW0r7WFNTySEWazuYwNpPmKm0uLo2O7ZjT4wt9OolU+hvp1Eqn0NfLuTbPjmnRypjILeOHz6/F2CyByS0fKRkvMkMgwPG+knaWnEaSo0O0fbVqG111MXoUX5tnxQqGCBxuYYz42hSLNfKxYaJFmvlYsNHnzyyVmEFxgmY53uXMs7U7h5FCHN3J5nOozFVyQZbxeRh0NFzaGYg9g3eGqeIAAsALLh0CG9+gx3P6oWijUOl5RHpvdLyjX31Ns9BgLcr4vccaCb6iDk0zvtGVMVLv/8AnzW+gtgnQIfzTPihOgQ/mmfFClPcrH+CU9xsf4NfnU0z2dbss4wPAKCe30FTqa55/RjGfQJvqLYJ0CH80z4oToEP5qP4oWfVLPB69Ts8GvzqaZ5/RjGPQJvqKo5M89H/AHYxj0Cb6i2BdAh/NM+KFXoMX5tnxQnqdngeqWGvwcmWeibHK+NeMUE31FNLkdwytwnJtLR10UscrdO7ZGkEXked4G4hXkI4x/Vt8y5AAbAo1+rlcsSI2o1cr1iRafKzQVWJZGr6OjhfNM/oeixjS4m0rDsAJ2AqFLOTLPF3MOV8Y3WIoJvqLYEQCLEXXDoUX5tnxQmn1cqE1ExRqZUrCIH4HyYZ2OLRaWXMWiZr0nSUUwHam3uFOHLUElPhMUUrS1wvcEW90V6AjYNjG+Zcl5v1Eru5i/USu7lH9qdV1EDnEZHzVimcH12G4FiVZE62uCklk/q4h7lp3g+ZTAXF0cbu2Y0+MLzRc6ZcSPFNrqeUa+m8m2eZJDfKmONHhw+Yf8ilPzZsAxXAssugxSinpZNXYyxOYe3lOxwHEedZd6FF+bZ8ULk1rW9q0DxBbb9XK5YZtu1UrVhlURFEIoREQHwrqSCsp3QVETJGOtdrmgjbff4lEDlk5G8bwvMb58t4NPWYfPbTZT0r5HN0WMAsGRgDsi7zKYyo5rXds0HxhbarpVPKNtV0qnlEFMm8nOdKbMdNJNlrGGQR6Wg6WhmGjdjr3JZYa1OimBbC0EW2/OuXQ4+4b5lyXu/USuabPd+odzTZ0saY99BI1jS4m2oC52hQNr+TbOz6qUNy3irgy1iKGYg3H7Cn/tXzFPAL/cY9e3sQs6fUyozwijUyp7Gvwcm+d76RyvjRa7uaCa4/4FSLk3zy298rYu4frYfMT9BbBBBABYQx2/ZCdAh/NR/FClPc7H+CS9ysf4NffU1zzc/7LYuB/wDAm+onU1zx+i2L+gTfUWwToEP5pnxQnQIPzUfxQsepWeDHqNng199TTPP6LYv/AIfN9RV6meef0Xxf0Cb6i2B9Ah/NM+KE6BD+aZ8UJ6nZ4MrcrF+DX31M88/ovjA8IoJr/QWVubpkrNGD52FfieD4lTQDfNTSM/q5R7poG0hSu6BD+aZ8ULk2Nje1Y0eILxZr52R4WjXbrp2LDRyGxERQCEfGri6NCWWvf+awFnnKmLNxZ8kGG1M0brW6FA53uW8GqQa4uijd20bHeNqrtx26vXQ4ZkDX7fXrYcMyLwyxj9rnCq2523p3/wAk+1jH+9Vb6O/+Sk90vT/mIviBOl6f8xF8QKgf8P07/wDRQfKOn/YjD9rOPd6630d/8ly+1fHu9tX8g/8AkpOdK035iL4gVOlab3vF8QLHyfp/2Hyjp/2Iyfavj3e2r+Qf/JU+1fHu9tX8g/8AkpOdK03veL4gTpWm97xfECfJ+n/YfKOn/YjH9q+Pd7av5B/8k+1fHu9tX8g/+Sk70tTe94viBOlqb3vF8QJ8n6f9h8o6f9iMX2r473tq/kH/AMk+1fHe9tX8g/8AkpOdK03veL4gTpWm97xfECfJ+n/YfKOn/YjH9q+O97av5B/8k+1fHu9tX8g/+Sk70rTe94fiBOlqb3vF8QJ8n6f9h8o6f9iMf2r493tq/kH/AMlQ5Xx4C/2MrD+4f/JSd6Wpve8XxAnS1P73i+IE+T9P+w+UdP8AsRh+1jHu9Vb6O/8Akq/avj/eqt9Hf/JSd6Wp/wAxF8QKvS9P+Yi+IE+TtP8Asx8o6f8AYjD9quYSNWEVx/8ArSfVXN2VcxOs37EVzTxNNJ9VScEMI2RRj4IToUX5tnxQva/iGmSxxHtfxPTJe4wrk3k1q5JGS4lE1rBfsXNIPuh7pniWZMOooKGnEMEbWNHAAbyd3jXYAA2ADxKqv9Dt1Oijw1ovdFt9Ojjw1oIiKeTgiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgP/Z" alt="SPSKY Pilotos"></div>\n'
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
        '    <div class="nav-item active"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>Resumen</div>\n'
        '    <div class="nav-item"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>Horas bloque</div>\n'
        '    <div class="nav-item"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>Horas deber</div>\n'
        '    <div class="nav-item"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>DAN 121</div>\n'
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
        '    <div id="placeholder" style="display:flex;flex-direction:column;align-items:center;justify-content:center;flex:1;gap:14px;color:var(--dim);padding:60px 0;">\n'
        '      <svg width="44" height="44" fill="none" stroke="currentColor" stroke-width="1.2" viewBox="0 0 24 24" style="stroke:var(--border2)"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/></svg>\n'
        '      <div style="font-family:var(--display);font-size:18px;color:var(--sand-400)">SDC \u00b7 SPSKY Digital Copilot</div>\n'
        '      <div style="font-size:12px;text-align:center;max-width:300px;line-height:1.7;color:var(--muted)">Seleccione un cargo y un tripulante para visualizar sus indicadores de productividad.</div>\n'
        '      <div style="font-size:10px;font-family:var(--mono);color:var(--dim);margin-top:4px" id="periodsHint"></div>\n'
        '    </div>\n'
        '    <div id="dashboard" style="display:none;flex-direction:column;gap:16px;">\n'
        '      <div class="kpi-grid" id="kpiRow"></div>\n'
        '      <div class="card">\n'
        '        <div class="card-head">\n'
        '          <div><div class="card-title">Horas Bloque \u00b7 Evoluci\u00f3n mensual</div><div class="card-sub">Piloto vs. promedio del cargo (meses activos)</div></div>\n'
        '          <div class="legend">\n'
        '            <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--clay)" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="var(--clay)"/></svg><span>Efectuado</span></div>\n'
        '            <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--dusk)" stroke-width="1.5" stroke-dasharray="2 2"/><rect x="5.5" y="1.5" width="5" height="5" transform="rotate(45 9 4)" fill="var(--dusk)"/></svg><span style="color:var(--dusk)">Solo programado</span></div>\n'
        '            <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--sand-400)" stroke-width="1.5" stroke-dasharray="4 3"/><circle cx="9" cy="4" r="2.5" fill="var(--sand-400)"/></svg><span>Prom. cargo</span></div>\n'
        '            <div class="leg"><svg width="14" height="12"><polygon points="7,1 13,11 1,11" fill="none" stroke="var(--rust)" stroke-width="1.5"/></svg><span style="color:var(--rust)">Excluido prom.</span></div>\n'
        '          </div>\n'
        '        </div>\n'
        '        <div class="chart-wrap"><canvas id="blockChart"></canvas></div>\n'
        '        <div class="excl-note" id="exclNote" style="display:none">\n'
        '          <svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>\n'
        '          <span id="exclText"></span>\n'
        '        </div>\n'
        '      </div>\n'
        '      <div class="charts-row">\n'
        '        <div class="card">\n'
        '          <div class="card-head">\n'
        '            <div><div class="card-title">Rol Programado vs. Efectuado</div><div class="card-sub">Horas bloque por per\u00edodo</div></div>\n'
        '            <div class="legend">\n'
        '              <div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--dusk);display:inline-block"></span><span>Programado</span></div>\n'
        '              <div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--clay);display:inline-block"></span><span>Efectuado</span></div>\n'
        '            </div>\n'
        '          </div>\n'
        '          <div class="chart-wrap"><canvas id="compareChart"></canvas></div>\n'
        '        </div>\n'
        '        <div class="card">\n'
        '          <div class="card-head"><div class="card-title">Comparativo por Per\u00edodo</div><div class="card-sub">Programado vs. efectuado \u00b7 \u0394 horas</div></div>\n'
        '          <div id="compTableWrap"></div>\n'
        '        </div>\n'
        '      </div>\n'
        '      <div class="bottom-row">\n'
        '        <div class="card"><div class="card-head"><div class="card-title">Acumulado &amp; Proyecci\u00f3n</div><div class="card-sub">Basado en meses activos</div></div><div class="prog-list" id="progList"></div></div>\n'
        '        <div class="card"><div class="card-head"><div class="card-title">Cumplimiento DAN 121</div><div class="card-sub">\u00daltimo per\u00edodo disponible</div></div><div class="alert-list" id="alertList"></div></div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </div>\n'
        '</div>\n'
        '</div>\n'
        '<script>\n'
        + JS +
        '</script>\n'
        '</body>\n'
        '</html>\n'
    )
    return html

# ── MAIN ───────────────────────────────────────────────────
if __name__ == '__main__':
    records, periods = build_dataset()
    html = generate_html(records, periods)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print('\nDashboard generado: ' + str(OUTPUT_HTML))
    print('Tamano: ' + str(len(html)//1024) + ' KB')
