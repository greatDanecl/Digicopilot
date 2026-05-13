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
        '--bg:#0F0A18;--surface:#18102A;--s2:#221840;--s3:#2D2050;'
        '--border:rgba(103,30,119,0.35);--border2:rgba(103,30,119,0.55);'
        '--text:#F0E8F8;--text2:#C8B5DC;--muted:#8B6FA8;--dim:#5A4275;'
        '--r:10px;--r2:14px;'
        '--shadow:0 1px 4px rgba(0,0,0,.4),0 4px 20px rgba(103,30,119,.15);'
        '--shadow2:0 2px 12px rgba(0,0,0,.5),0 8px 32px rgba(103,30,119,.25);'
        "--font:'DM Sans',sans-serif;--display:'Playfair Display',serif;--mono:'DM Mono',monospace;"
        '}'
        '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}'
        'html{font-size:14px}'
        'body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.5}'
        '.shell{display:grid;grid-template-columns:230px 1fr;min-height:100vh}'
        '.sidebar{background:linear-gradient(180deg,#1A0B2E 0%,#0F0A18 100%);display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;border-right:1px solid var(--border)}'
        '.sidebar-top{padding:0 0 14px;border-bottom:1px solid var(--border)}'
        '.logo-wrap{width:100%;background:#000;display:flex;align-items:center;justify-content:center;padding:14px 18px}'
        '.logo-wrap img{width:100%;max-width:192px;height:auto;display:block}'
        '.brand-sub-line{font-size:9px;color:rgba(255,255,255,.3);letter-spacing:.1em;text-transform:uppercase;text-align:center;padding:5px 0 0;font-family:var(--mono)}'
        '.filters{padding:14px 16px;display:flex;flex-direction:column;gap:11px;border-bottom:1px solid var(--border)}'
        '.f-block{display:flex;flex-direction:column;gap:5px}'
        '.f-label{font-size:9px;color:rgba(255,255,255,.35);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono)}'
        '.f-select{appearance:none;background:rgba(103,30,119,.12);border:1px solid rgba(103,30,119,.4);border-radius:8px;color:var(--text);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;background-image:url(\"data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'12\' height=\'12\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%238B6FA8\' stroke-width=\'2\'%3E%3Cpolyline points=\'6 9 12 15 18 9\'/%3E%3C/svg%3E\");background-repeat:no-repeat;background-position:right 9px center}'
        '.f-select:focus,.f-select:hover{border-color:var(--green);box-shadow:0 0 0 2px rgba(38,216,0,.15)}'
        '.f-select option{background:#18102A;color:var(--text)}'
        '.sidebar-nav{padding:10px 8px;flex:1}'
        '.nav-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:var(--muted);cursor:pointer;transition:all .15s;margin-bottom:2px;border-left:2px solid transparent}'
        '.nav-item:hover{color:var(--text);background:var(--purple-dim);border-left-color:var(--purple-l)}'
        '.nav-item.active{color:var(--green-l);background:var(--green-dim2);border-left-color:var(--green)}'
        '.nav-item svg{width:14px;height:14px;flex-shrink:0}'
        '.sidebar-footer{padding:12px 16px;border-top:1px solid var(--border)}'
        '.pilot-badge{display:flex;align-items:center;gap:10px}'
        '.pilot-avatar{width:34px;height:34px;border-radius:50%;background:var(--purple);border:1.5px solid var(--purple-l);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}'
        '.pilot-name-s{font-size:11px;font-weight:500;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
        '.pilot-pos-s{font-size:10px;color:var(--muted);font-family:var(--mono)}'
        '.main{display:flex;flex-direction:column;min-height:100vh}'
        '.topbar{background:rgba(24,16,42,0.95);border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;backdrop-filter:blur(12px)}'
        '.page-title{font-family:var(--display);font-size:17px;color:var(--text)}'
        '.page-title span{color:var(--green-l)}'
        '.page-sub{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.topbar-right{display:flex;align-items:center;gap:8px}'
        '.pill{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--s2);color:var(--muted)}'
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
        '.comp-table th{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--s2)}'
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
        '::-webkit-scrollbar-thumb{background:rgba(103,30,119,.7);border-radius:2px}'
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
        '.sidebar-overlay{display:none;position:fixed;inset:0;background:rgba(15,10,24,.7);z-index:155;backdrop-filter:blur(2px)}'
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
        '    (convPct !== null ? \'<span class="delta \' + (convPct>=80?"d-up":convPct>=60?"d-warn":"d-down") + \'">\' + convPct.toFixed(0) + \'% conv.</span>\' : "") + \'</div></div>\' +\n'
        '    \'<div class="kpi k-g3"><div class="kpi-label">Block acum. YTD</div><div class="kpi-val">\' + fmt(accB,0) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">\' + actP.length + \' meses activos</span><span class="delta d-neu">/\' + PERIODS.length + \'m</span></div></div>\' +\n'
        '    \'<div class="kpi k-p2"><div class="kpi-label">D\u00edas blancos \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div>\' +\n'
        '    \'<div class="kpi-val">\' + (blancos !== null ? blancos : "\\u2014") + \'<span class="kpi-unit">d</span></div>\' +\n'
        '    \'<div class="kpi-footer"><span class="kpi-vs">Sin asignaci\u00f3n (rol ef.)</span><span class="delta \' + (blancos > 5 ? "d-warn" : "d-up") + \'">\' + (blancos !== null ? (blancos > 5 ? "\\u26a0 revisar" : "\\u2713 ok") : "\\u2014") + \'</span></div></div>\';\n'
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
        '    const ex = r.exclude_from_avg ? \'<span style="color:var(--rust);font-size:9px"> \\u2731</span>\' : "";\n'
        '    const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\\u2014";\n'
        '    const blCell = bl !== null ? (bl > 5 ? \'<span style="color:var(--warm-red)">\'+bl+\'</span>\' : bl) : "\\u2014";\n'
        '    tbl += \'<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">\' + (PERIOD_LABELS[p]||p) + ex + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (pg>0?pg.toFixed(1)+"h":"\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ac>0?ac.toFixed(1)+"h":"\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + (d>=0?"var(--sage)":"var(--warm-red)") + \'">\' + dstr + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (tp !== null ? tp : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (vp !== null ? vp : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ve !== null ? ve : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + blCell + \'</td>\'\n'
        '         + \'</tr>\';\n'
        '  });\n'
        '  if (excl.length) tbl += \'<tr><td colspan="8" style="font-size:9px;color:var(--muted);font-family:var(--mono);padding:6px 10px">\\u2731 Excluido del promedio comparativo</td></tr>\';\n'
        '  tbl += "</tbody></table>";\n'
        "  document.getElementById('compTableWrap').innerHTML = tbl;\n"
        '\n'
        '  // Progress bars\n'
        '  const pct1 = Math.min(accB/1000*100, 100);\n'
        '  const avgM = actP.length ? accB/actP.length : 0;\n'
        '  const proj = avgM * 12;\n'
        '  const pctP = Math.min(proj/1000*100, 100);\n'
        '  const totL = pr.reduce((s,r) => s+(r.libre_days||0), 0);\n'
        '  const avgL = pr.length ? totL/pr.length : 0;\n'
        "  document.getElementById('progList').innerHTML =\n"
        '    \'<div><div class="prog-head"><span class="prog-lbl">Horas bloque acumuladas</span><span class="prog-num" style="color:var(--clay)">\' + accB.toFixed(0) + \'h</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pct1 + \'%;background:var(--clay)"></div></div><div class="prog-note">L\\u00edmite DAN 121: 1.000h/a\\u00f1o \\u00b7 \' + (100-pct1).toFixed(1) + \'% disponible</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Proyecci\\u00f3n a 12 meses</span><span class="prog-num" style="color:var(--sand-500)">~\' + proj.toFixed(0) + \'h est.</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pctP + \'%;background:linear-gradient(90deg,var(--green),var(--teal))"></div></div><div class="prog-note">Prom. \' + avgM.toFixed(1) + \'h/mes en meses activos</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Descanso promedio</span><span class="prog-num" style="color:var(--sage)">\' + avgL.toFixed(1) + \' d/mes</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + Math.min(avgL/20*100,100) + \'%;background:var(--sage)"></div></div><div class="prog-note">M\\u00ednimo reglamentario DAN 121: 8 d\\u00edas/mes</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Meses activos</span><span class="prog-num">\' + actP.length + \' / \' + PERIODS.length + \'</span></div><div style="display:flex;gap:3px;margin-top:4px"><div style="height:5px;border-radius:2px 0 0 2px;background:var(--sage);flex:\' + actP.length + \'"></div><div style="height:5px;border-radius:0 2px 2px 0;background:var(--warn);opacity:.5;flex:\' + Math.max(PERIODS.length-actP.length,0) + \'"></div></div><div class="prog-note">\' + (excl.length?excl.map(p=>PERIOD_LABELS[p]||p).join(", ")+" excluidos":"Sin ausencias prolongadas") + \'</div></div>\';\n'
        '\n'
        '  // Alerts\n'
        '  function alrt(t,title,desc){return \'<div class="alert \'+t+\'"><div class="alert-dot"></div><div><div class="alert-title">\'+title+\'</div><div class="alert-desc">\'+desc+\'</div></div></div>\';}\n'
        '  let alerts = "";\n'
        '  alerts += alrt(mb>100?"danger":mb>85?"warn":"ok", "Bloque mensual \\u00b7 "+fmt(mb)+"h", mb>100?"Supera l\\u00edmite DAN 121 de 100h/mes":mb>85?"Cercano al l\\u00edmite de 100h/mes":"Dentro del l\\u00edmite (100h/mes)");\n'
        '  alerts += alrt(accB>900?"danger":accB>750?"warn":"ok", "Bloque acumulado \\u00b7 "+accB.toFixed(0)+"h", accB>900?"Muy cerca del l\\u00edmite anual de 1.000h":accB>750?"Supera el 75% del l\\u00edmite anual":"Sin riesgo l\\u00edmite anual ("+(1000-accB).toFixed(0)+"h disp.)");\n'
        '  alerts += alrt(ml<8?"danger":ml<10?"warn":"ok", "D\\u00edas libres \\u00b7 "+ml+"d", ml<8?"Bajo el m\\u00ednimo reglamentario (8d/mes)":ml<10?"Dentro del m\\u00ednimo, bajo el promedio del cargo":"Descanso adecuado seg\\u00fan DAN 121");\n'
        '  alerts += alrt(md>130?"danger":md>105?"warn":"ok", "Horas deber \\u00b7 "+fmt(md)+"h", md>130?"Horas deber muy elevadas, revisar FDPs":md>105?"Sobre promedio del cargo":"Dentro de rango normal");\n'
        '  alerts += \'<div style="margin-top:6px;padding:9px 11px;background:var(--sand-100);border-radius:7px;font-size:10px;color:var(--muted);line-height:1.5;font-family:var(--mono)">Alertas indicativas. El c\\u00e1lculo oficial de FDP y l\\u00edmites es responsabilidad de Operaciones.</div>\';\n'
        "  document.getElementById('alertList').innerHTML = alerts;\n"
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
        '    <div class="logo-wrap"><img src="data:image/png;base64,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" alt="SPSKY Pilotos"></div>\n'
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
