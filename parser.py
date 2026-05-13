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
    Soporta: efec_Feb_2026_SCL.xlsx, prog_Mar_2026_PMC.xlsx, y variantes.
    """
    fl = fname.lower()
    # Pattern: cualquier mes en texto + año de 4 dígitos
    month_names = r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|ene|abr|ago|dic)'
    year_names  = r'(20\d{2})'
    # mes antes que año: efec_feb_2026_SCL
    m = re.search(month_names + r'[_\-\s]*' + year_names, fl)
    if m:
        mon = m.group(1).capitalize()
        yr  = m.group(2)
        code = MONTH_MAP.get(mon, '00')
        if code != '00': return yr + '-' + code
    # año antes que mes: 2026_feb
    m2 = re.search(year_names + r'[_\-\s]*' + month_names, fl)
    if m2:
        mon = m2.group(2).capitalize()
        yr  = m2.group(1)
        code = MONTH_MAP.get(mon, '00')
        if code != '00': return yr + '-' + code
    # año + número de mes: 2026-02
    m3 = re.search(r'(20\d{2})[-_](0[1-9]|1[0-2])', fl)
    if m3: return m3.group(1) + '-' + m3.group(2)
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
    Convención: efec_MMM_AAAA_BASE.xlsx / prog_MMM_AAAA_BASE.xlsx
    """
    fl, sl = fname.lower(), sheet_name.lower()
    # Convención principal: empieza con efec o prog
    if fl.startswith('efec'):  return 'actual'
    if fl.startswith('prog'):  return 'programmed'
    # Variantes largas
    if fl.startswith('efectuado'):  return 'actual'
    if fl.startswith('programado'): return 'programmed'
    # Fallbacks por contenido del nombre
    if any(w in fl for w in ['horas','actual','efect','real','flown']): return 'actual'
    if any(w in fl for w in ['plan','sched','mando','master']):         return 'programmed'
    # Fallback por nombre de hoja
    if any(w in sl for w in ['hora','actual','efect']): return 'actual'
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
            cred_h  = parse_td(df.iloc[i+5, 1] if i+5 < len(df) else None)
            blk_h   = parse_td(df.iloc[i+6, 1] if i+6 < len(df) else None)
            duty_h  = parse_td(df.iloc[i+7, 1] if i+7 < len(df) else None)
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
        '--sand-50:#FAF7F2;--sand-100:#F3EDE2;--sand-200:#E8DCC8;--sand-300:#D4C4A8;'
        '--sand-400:#BFA882;--sand-500:#A68B5B;--sand-600:#8A7048;--sand-700:#6B5535;'
        '--earth-800:#4A3B25;--earth-900:#2E2416;'
        '--clay:#C4856A;--clay-dim:rgba(196,133,106,0.12);'
        '--rust:#B5603A;--sage:#7A9E7E;--sage-dim:rgba(122,158,126,0.12);'
        '--dusk:#8B7BA8;--dusk-dim:rgba(139,123,168,0.12);'
        '--warm-red:#C4534A;--warm-red-dim:rgba(196,83,74,0.10);'
        '--bg:#F7F3ED;--surface:#FDFAF6;--surface2:#F3EDE2;'
        '--border:#E2D8C8;--border2:#D4C4A8;'
        '--text:#2E2416;--text2:#6B5535;--muted:#A68B5B;--dim:#BFA882;'
        '--r:10px;--r2:16px;'
        '--shadow:0 1px 3px rgba(46,36,22,.06),0 4px 16px rgba(46,36,22,.04);'
        '--shadow2:0 2px 8px rgba(46,36,22,.08),0 8px 32px rgba(46,36,22,.06);'
        "--font:'DM Sans',sans-serif;--display:'Playfair Display',serif;--mono:'DM Mono',monospace;"
        '}'
        '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}'
        'html{font-size:14px}'
        'body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.5}'
        '.shell{display:grid;grid-template-columns:230px 1fr;min-height:100vh}'
        '.sidebar{background:var(--earth-800);display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto}'
        '.sidebar-top{padding:26px 22px 18px;border-bottom:1px solid rgba(255,255,255,.07)}'
        '.brand{display:flex;align-items:center;gap:10px;margin-bottom:3px}'
        '.brand-icon{width:30px;height:30px;background:var(--clay);border-radius:7px;display:flex;align-items:center;justify-content:center}'
        '.brand-icon svg{width:15px;height:15px;stroke:white}'
        '.brand-name{font-family:var(--display);font-size:15px;color:var(--sand-100);letter-spacing:.02em}'
        '.brand-sub{font-size:9px;color:rgba(255,255,255,.3);letter-spacing:.09em;text-transform:uppercase;margin-left:40px;font-family:var(--mono)}'
        '.filters{padding:18px 22px;display:flex;flex-direction:column;gap:13px;border-bottom:1px solid rgba(255,255,255,.07)}'
        '.f-block{display:flex;flex-direction:column;gap:5px}'
        '.f-label{font-size:9px;color:rgba(255,255,255,.35);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono)}'
        '.f-select{appearance:none;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:8px;'
        "color:var(--sand-100);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;"
        "background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%23BFA882' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E\");"
        'background-repeat:no-repeat;background-position:right 9px center}'
        '.f-select:focus,.f-select:hover{border-color:var(--clay);background-color:rgba(255,255,255,.1)}'
        '.f-select option{background:var(--earth-800);color:var(--sand-100)}'
        '.sidebar-nav{padding:14px 12px;flex:1}'
        '.nav-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:rgba(255,255,255,.4);cursor:pointer;transition:all .15s;margin-bottom:2px}'
        '.nav-item:hover,.nav-item.active{color:var(--sand-100);background:rgba(196,133,106,.16)}'
        '.nav-item svg{width:14px;height:14px;flex-shrink:0}'
        '.sidebar-footer{padding:14px 22px;border-top:1px solid rgba(255,255,255,.07)}'
        '.pilot-badge{display:flex;align-items:center;gap:10px}'
        '.pilot-avatar{width:34px;height:34px;border-radius:50%;background:var(--clay);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}'
        '.pilot-name-s{font-size:11px;font-weight:500;color:var(--sand-100);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
        '.pilot-pos-s{font-size:10px;color:rgba(255,255,255,.35);font-family:var(--mono)}'
        '.main{display:flex;flex-direction:column;min-height:100vh}'
        '.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:14px 30px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}'
        '.page-title{font-family:var(--display);font-size:17px;color:var(--text)}'
        '.page-title span{color:var(--clay)}'
        '.page-sub{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.topbar-right{display:flex;align-items:center;gap:8px}'
        '.pill{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--surface2);color:var(--text2)}'
        '.dot{width:6px;height:6px;border-radius:50%;background:var(--sage)}'
        '.content{padding:22px 30px;display:flex;flex-direction:column;gap:16px;flex:1}'
        '.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:11px}'
        '.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 16px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:box-shadow .2s,transform .15s}'
        '.kpi:hover{box-shadow:var(--shadow2);transform:translateY(-1px)}'
        ".kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--r2) var(--r2) 0 0}"
        '.kpi.k-clay::before{background:var(--clay)}.kpi.k-sage::before{background:var(--sage)}'
        '.kpi.k-dusk::before{background:var(--dusk)}.kpi.k-sand::before{background:var(--sand-400)}'
        '.kpi.k-rust::before{background:var(--rust)}'
        '.kpi-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;font-family:var(--mono)}'
        '.kpi-val{font-size:25px;font-weight:600;color:var(--text);font-family:var(--mono);letter-spacing:-.03em;line-height:1}'
        '.kpi-unit{font-size:12px;font-weight:400;color:var(--muted);margin-left:2px}'
        '.kpi-footer{display:flex;align-items:center;justify-content:space-between;margin-top:7px}'
        '.kpi-vs{font-size:10px;color:var(--muted)}.kpi-vs b{color:var(--text2);font-weight:500}'
        '.delta{font-size:10px;font-family:var(--mono);padding:2px 6px;border-radius:4px}'
        '.d-up{background:var(--sage-dim);color:var(--sage)}.d-down{background:var(--warm-red-dim);color:var(--warm-red)}'
        '.d-neu{background:var(--sand-100);color:var(--muted)}'
        '.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow)}'
        '.card-head{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}'
        '.card-title{font-size:13px;font-weight:500;color:var(--text)}'
        '.card-sub{font-size:10px;color:var(--muted);margin-top:2px;font-family:var(--mono)}'
        '.legend{display:flex;gap:12px;align-items:center;font-size:10px;color:var(--muted);font-family:var(--mono);flex-wrap:wrap}'
        '.leg{display:flex;align-items:center;gap:5px}'
        '.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}'
        '.chart-wrap{position:relative;height:220px}'
        '.comp-table{width:100%;border-collapse:collapse;font-size:12px}'
        '.comp-table th{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--sand-50)}'
        '.comp-table td{padding:8px 10px;border-bottom:1px solid rgba(226,216,200,.5)}'
        '.comp-table tr:last-child td{border-bottom:none}.comp-table tr:hover td{background:var(--sand-50)}'
        '.winner{display:inline-flex;align-items:center;font-size:10px;padding:2px 7px;border-radius:4px;font-family:var(--mono);font-weight:500}'
        '.w-prog{background:var(--dusk-dim);color:var(--dusk)}.w-act{background:var(--sage-dim);color:var(--sage)}.w-eq{background:var(--sand-100);color:var(--muted)}'
        '.bottom-row{display:grid;grid-template-columns:1fr 300px;gap:14px}'
        '.prog-list{display:flex;flex-direction:column;gap:13px}'
        '.prog-head{display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px}'
        '.prog-lbl{color:var(--text2)}.prog-num{font-family:var(--mono);font-size:11px}'
        '.prog-track{height:5px;background:var(--sand-200);border-radius:3px;overflow:hidden}'
        '.prog-fill{height:100%;border-radius:3px}'
        '.prog-note{font-size:9px;color:var(--dim);margin-top:2px;font-family:var(--mono)}'
        '.alert-list{display:flex;flex-direction:column;gap:7px}'
        '.alert{display:flex;align-items:flex-start;gap:9px;padding:9px 11px;border-radius:8px;border:1px solid}'
        '.alert.ok{background:var(--sage-dim);border-color:rgba(122,158,126,.25)}'
        '.alert.warn{background:rgba(181,96,58,.07);border-color:rgba(181,96,58,.2)}'
        '.alert.danger{background:var(--warm-red-dim);border-color:rgba(196,83,74,.2)}'
        '.alert-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px}'
        '.alert.ok .alert-dot{background:var(--sage)}.alert.warn .alert-dot{background:var(--rust)}.alert.danger .alert-dot{background:var(--warm-red)}'
        '.alert-title{font-size:11px;font-weight:500;color:var(--text)}'
        '.alert-desc{font-size:10px;color:var(--muted);margin-top:1px;font-family:var(--mono)}'
        '.excl-note{display:flex;align-items:center;gap:6px;padding:8px 11px;border-radius:7px;background:var(--sand-100);border:1px solid var(--border2);font-size:10px;color:var(--muted);margin-top:10px}'
        '.excl-note svg{width:12px;height:12px;flex-shrink:0}'
        '::-webkit-scrollbar{width:4px}::-webkit-scrollbar-thumb{background:var(--sand-300);border-radius:2px}'
        '@keyframes fadeUp{from{opacity:0;transform:translateY(7px)}to{opacity:1;transform:none}}'
        '.kpi,.card{animation:fadeUp .28s ease both}'
        '.kpi:nth-child(1){animation-delay:.04s}.kpi:nth-child(2){animation-delay:.08s}'
        '.kpi:nth-child(3){animation-delay:.12s}.kpi:nth-child(4){animation-delay:.16s}'
        '.kpi:nth-child(5){animation-delay:.20s}.kpi:nth-child(6){animation-delay:.24s}'
        '.turnos-bar{display:flex;align-items:center;gap:8px;margin-top:6px}'
        '.tbar-track{flex:1;height:7px;background:var(--sand-200);border-radius:4px;overflow:hidden;position:relative}'
        '.tbar-prog{height:100%;background:var(--dusk);border-radius:4px;position:absolute;left:0;top:0}'
        '.tbar-act{height:100%;background:var(--clay);border-radius:4px;position:absolute;left:0;top:0;opacity:.85}'
        '.tbar-label{font-size:10px;font-family:var(--mono);color:var(--muted);white-space:nowrap}'
        '.hamburger{display:none;position:fixed;top:12px;left:12px;z-index:200;width:38px;height:38px;border-radius:8px;border:1px solid var(--border2);background:var(--surface);cursor:pointer;align-items:center;justify-content:center;flex-direction:column;gap:5px;padding:9px}'
        '.hamburger span{display:block;width:16px;height:1.5px;background:var(--text2);border-radius:2px;transition:all .2s}'
        '.hamburger.open span:nth-child(1){transform:translateY(6.5px) rotate(45deg)}'
        '.hamburger.open span:nth-child(2){opacity:0}'
        '.hamburger.open span:nth-child(3){transform:translateY(-6.5px) rotate(-45deg)}'
        '.sidebar-overlay{display:none;position:fixed;inset:0;background:rgba(46,36,22,.45);z-index:155;backdrop-filter:blur(2px)}'
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
        "const selGroup = document.getElementById('selGroup');\n"
        "const selPilot = document.getElementById('selPilot');\n"
        '\n'
        "selGroup.addEventListener('change', () => {\n"
        '  const g = selGroup.value;\n'
        '  const names = [...new Set(RAW.filter(r => r.pos_group === g).map(r => r.name))].sort((a,b) => a.localeCompare(b, "es"));\n'
        "  selPilot.innerHTML = '<option value=\"\">— Seleccionar tripulante —</option>';\n"
        '  names.forEach(n => { const o = document.createElement("option"); o.value = o.textContent = n; selPilot.appendChild(o); });\n'
        '  selPilot.disabled = false;\n'
        "  document.getElementById('placeholder').style.display = 'flex';\n"
        "  document.getElementById('dashboard').style.display = 'none';\n"
        '});\n'
        "selPilot.addEventListener('change', () => { if (selPilot.value) render(selPilot.value, selGroup.value); });\n"
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
        "  document.getElementById('placeholder').style.display = 'none';\n"
        "  document.getElementById('dashboard').style.display = 'flex';\n"
        '  const pr = RAW.filter(r => r.name === pilotName);\n'
        '  const gr = RAW.filter(r => r.pos_group === group);\n'
        '  const latest = pr.filter(r => r.block_h_actual > 0).sort((a,b) => b.period.localeCompare(a.period))[0] || pr[0];\n'
        '  const lp = latest ? latest.period : PERIODS[PERIODS.length-1];\n'
        '  const init = pilotName.split(" ").filter((_,i) => i < 2).map(w => w[0]).join("");\n'
        "  document.getElementById('sideAvatar').textContent = init;\n"
        "  document.getElementById('sideName').textContent = pilotName.split(' ').slice(0,2).join(' ');\n"
        "  document.getElementById('sidePos').textContent = (latest ? latest.pos : group) + ' \u00b7 ' + (latest ? latest.base : '');\n"
        "  document.getElementById('pageTitle').innerHTML = '<span>' + pilotName.split(' ').slice(0,2).join(' ') + '</span> \u00b7 Productividad';\n"
        "  document.getElementById('pageSub').textContent = (latest ? latest.pos_group : group) + ' \u00b7 ' + (latest ? latest.base : '') + ' \u00b7 ' + Object.values(PERIOD_LABELS).join(' \u00b7 ');\n"
        '\n'
        '  const ga = gr.filter(r => r.period === lp && !r.exclude_from_avg && r.block_h_actual > 0);\n'
        '  const ab = avg(ga.map(r => r.block_h_actual));\n'
        '  const ad = avg(ga.map(r => r.duty_h_actual));\n'
        '  const al = avg(ga.map(r => r.libre_days));\n'
        '  const mb = latest ? (latest.block_h_actual || 0) : 0;\n'
        '  const md = latest ? (latest.duty_h_actual  || 0) : 0;\n'
        '  const ml = latest ? (latest.libre_days     || 0) : 0;\n'
        '  const bd = ab > 0 ? (mb-ab)/ab*100 : 0;\n'
        '  const dd = ad > 0 ? (md-ad)/ad*100 : 0;\n'
        '  const actP   = pr.filter(r => !r.exclude_from_avg && r.block_h_actual > 0);\n'
        '  const accB   = actP.reduce((s,r) => s + (r.block_h_actual    || 0), 0);\n'
        '  const accProg = pr.reduce((s,r)  => s + (r.block_h_programmed || 0), 0);\n'
        '  const accAct  = pr.reduce((s,r)  => s + (r.block_h_actual    || 0), 0);\n'
        '  const pva = accProg > 0 ? (accAct - accProg)/accProg*100 : 0;\n'
        '  const turnos  = latest ? (latest.turnos_programados || null) : null;\n'
        '  const vuelos  = latest ? (latest.vuelos_efectuados  || null) : null;\n'
        '  const vProg   = latest ? (latest.vuelos_programados || null) : null;\n'
        '  const blancos = latest ? (latest.dias_blancos        || null) : null;\n'
        '  // Conversion: turnos efectivamente volados = vuelos en el rol efectuado / turnos en el programado\n'
        '  const convPct = (turnos != null && vuelos != null && turnos > 0) ? vuelos/turnos*100 : null;\n'
        '\n'
        "  document.getElementById('kpiRow').innerHTML =\n"
        '    \'<div class="kpi k-clay"><div class="kpi-label">Block hours \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + fmt(mb) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Prom.: <b>\' + fmt(ab) + \'h</b></span><span class="delta \' + dc(bd) + \'">\' + ds(bd) + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-sand"><div class="kpi-label">Duty hours \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + fmt(md) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Prom.: <b>\' + fmt(ad) + \'h</b></span><span class="delta \' + dc(dd) + \'">\' + ds(dd) + \'</span></div></div>\' +\n'
        '    \'<div class="kpi k-sage"><div class="kpi-label">D\u00edas libres \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div><div class="kpi-val">\' + ml + \'<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Prom.: <b>\' + fmt(al,0) + \'d</b></span><span class="delta \' + dc(ml-al) + \'">\' + (ml-al>=0?"+":"") + (ml-al).toFixed(0) + \'d</span></div></div>\' +\n'
        '    \'<div class="kpi k-dusk"><div class="kpi-label">Turnos prog. \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div>\' +\n'
        '    \'<div class="kpi-val">\' + (turnos !== null ? turnos : "\\u2014") + \'<span class="kpi-unit">\' + (vuelos !== null ? " / "+vuelos+" ef." : "") + \'</span></div>\' +\n'
        '    \'<div class="kpi-footer"><span class="kpi-vs">\' + (vProg !== null ? vProg+" vuelos prog." : "Sin datos prog.") + \'</span>\' +\n'
        '    (convPct !== null ? \'<span class="delta \' + (convPct>=80?"d-up":convPct>=60?"d-warn":"d-down") + \'">\' + convPct.toFixed(0) + \'% conv.</span>\' : "") + \'</div></div>\' +\n'
        '    \'<div class="kpi k-rust"><div class="kpi-label">Block acum.</div><div class="kpi-val">\' + fmt(accB,0) + \'<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">\' + actP.length + \' meses activos</span><span class="delta d-neu">/\' + PERIODS.length + \'m</span></div></div>\' +\n'
        '    \'<div class="kpi k-sand"><div class="kpi-label">D\u00edas blancos \u00b7 \' + (PERIOD_LABELS[lp]||lp) + \'</div>\' +\n'
        '    \'<div class="kpi-val">\' + (blancos !== null ? blancos : "\\u2014") + \'<span class="kpi-unit">d</span></div>\' +\n'
        '    \'<div class="kpi-footer"><span class="kpi-vs">Das sin asignaci\\u00f3n (rol ef.)</span><span class="delta \' + (blancos > 5 ? "d-warn" : "d-up") + \'">\' + (blancos !== null ? (blancos > 5 ? "\\u26a0 revisar" : "\\u2713 ok") : "\\u2014") + \'</span></div></div>\';\n'
        '\n'
        '  // Line chart\n'
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
        "      { label:'Piloto', data:pData, borderColor:'#C4856A',\n"
        "        backgroundColor(c) { return makeGrad(bc, c.chart.chartArea, 'rgba(196,133,106,.15)', 'rgba(196,133,106,.01)'); },\n"
        '        borderWidth:2.5,\n'
        '        pointRadius(c)          { return excl.includes(PERIODS[c.dataIndex]) ? 6 : 4; },\n'
        "        pointStyle(c)           { return excl.includes(PERIODS[c.dataIndex]) ? 'triangle' : progOnlyIdx.includes(c.dataIndex) ? 'rectRot' : 'circle'; },\n"
        "        pointBackgroundColor(c) { return excl.includes(PERIODS[c.dataIndex]) ? '#B5603A' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#C4856A'; },\n"
        "        pointBorderColor(c)     { return excl.includes(PERIODS[c.dataIndex]) ? '#B5603A' : progOnlyIdx.includes(c.dataIndex) ? '#8B7BA8' : '#C4856A'; },\n"
        '        pointHoverRadius:7, tension:.35, fill:true, spanGaps:true, order:1 },\n'
        "      { label:'Prom. cargo', data:gData, borderColor:'#BFA882', borderWidth:1.5, borderDash:[5,4],\n"
        "        pointBackgroundColor:'#BFA882', pointRadius:3, pointHoverRadius:5,\n"
        '        tension:.35, fill:false, spanGaps:false, order:2 }\n'
        '    ]},\n'
        '    options: { responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false},\n'
        '      plugins: { legend:{display:false}, tooltip:{\n'
        "        backgroundColor:'#FAF7F2', borderColor:'#E2D8C8', borderWidth:1,\n"
        "        titleColor:'#2E2416', bodyColor:'#A68B5B', padding:11,\n"
        "        titleFont:{family:\"'DM Sans',sans-serif\",size:12,weight:500},\n"
        "        bodyFont:{family:\"'DM Mono',monospace\",size:11},\n"
        '        callbacks:{\n'
        '          title(i)     { const p=PERIODS[i[0].dataIndex]; const ex=excl.includes(p); const po=progOnlyIdx.includes(i[0].dataIndex); return (PERIOD_LABELS[p]||p)+(ex?" \u00b7 \u26a0 excluido del prom.":po?" \u00b7 solo programado":""); },\n'
        '          label(i)     { if(i.raw==null||i.raw===0)return null; const po=progOnlyIdx.includes(i.dataIndex)&&i.datasetIndex===0; return "  "+i.dataset.label+(po?" (prog.)":"")+": "+i.raw.toFixed(1)+"h"; },\n'
        '          afterBody(i) { const p=PERIODS[i[0].dataIndex]; const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex]; if(av==null||my==null||my===0)return[]; const d=my-av; return["  vs prom. cargo: "+(d>=0?"+":"")+d.toFixed(1)+"h"]; }\n'
        '        }\n'
        '      }},\n'
        "      scales:{\n"
        "        x:{grid:{color:'rgba(226,216,200,.6)',drawBorder:false},ticks:{color:'#A68B5B',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},\n"
        "        y:{min:0,grid:{color:'rgba(226,216,200,.6)',drawBorder:false},ticks:{color:'#A68B5B',font:{size:11,family:\"'DM Mono',monospace\"},callback:function(v){return v+'h';}},border:{display:false}}\n"
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
        "      {label:'Programado',data:prog,backgroundColor:'rgba(139,123,168,.55)',borderColor:'#8B7BA8',borderWidth:1,borderRadius:5,borderSkipped:false},\n"
        "      {label:'Efectuado', data:act, backgroundColor:'rgba(196,133,106,.55)',borderColor:'#C4856A',borderWidth:1,borderRadius:5,borderSkipped:false}\n"
        '    ]},\n'
        "    options:{responsive:true,maintainAspectRatio:false,\n"
        '      plugins:{legend:{display:false},tooltip:{\n'
        "        backgroundColor:'#FAF7F2',borderColor:'#E2D8C8',borderWidth:1,titleColor:'#2E2416',bodyColor:'#A68B5B',padding:11,\n"
        "        bodyFont:{family:\"'DM Mono',monospace\",size:11},\n"
        '        callbacks:{\n'
        '          label(i){return "  "+i.dataset.label+": "+i.raw.toFixed(1)+"h";},\n'
        '          afterBody(i){const idx=i[0].dataIndex;const d=act[idx]-prog[idx];if(prog[idx]===0&&act[idx]===0)return["  Sin datos"];const w=d>.5?"\u25b2 Efectuado mayor":d<-.5?"\u25b2 Programado mayor":"\u2248 Similares";return["  \u0394: "+(d>=0?"+":"")+d.toFixed(1)+"h  "+w];}\n'
        '        }\n'
        '      }},\n'
        "      scales:{\n"
        "        x:{grid:{display:false},ticks:{color:'#A68B5B',font:{size:11,family:\"'DM Mono',monospace\"}},border:{display:false}},\n"
        "        y:{min:0,grid:{color:'rgba(226,216,200,.6)',drawBorder:false},ticks:{color:'#A68B5B',font:{size:11,family:\"'DM Mono',monospace\"},callback:function(v){return v+'h';}},border:{display:false}}\n"
        '      }\n'
        '    }\n'
        '  });\n'
        '\n'
        '  // Comparison table\n'
        '  let tbl = \'<table class="comp-table"><thead><tr><th>Per\\u00edodo</th><th>Block prog.</th><th>Block ef.</th><th>\\u0394 Block</th><th>Turnos</th><th>Vuelos prog.</th><th>Vuelos ef.</th><th>Conv.</th><th>Blancos</th></tr></thead><tbody>\';\n'
        '  PERIODS.forEach(p => {\n'
        '    const r = pr.find(x => x.period===p); if(!r) return;\n'
        '    const pg=r.block_h_programmed||0, ac=r.block_h_actual||0, d=ac-pg;\n'
        '    const tp=r.turnos_programados, vp=r.vuelos_programados;\n'
        '    const ve=r.vuelos_efectuados,  bl=r.dias_blancos;\n'
        '    const cp = (tp != null && ve != null && tp > 0) ? ve/tp*100 : null;\n'
        '    const ex = r.exclude_from_avg ? \'<span style="color:var(--rust);font-size:9px"> \\u2731</span>\' : "";\n'
        '    const cpCell = cp !== null\n'
        '      ? \'<span style="font-family:var(--mono);font-size:11px;color:\' + (cp>=80?"var(--sage)":cp>=60?"var(--rust)":"var(--warm-red)") + \'">\' + cp.toFixed(0) + \'%</span>\'\n'
        '      : \'<span style="color:var(--dim)">\\u2014</span>\';\n'
        '    const dstr = pg > 0 ? ((d>=0?"+":"")+d.toFixed(1)+"h") : "\\u2014";\n'
        '    const blCell = bl !== null ? (bl > 5 ? \'<span style="color:var(--warm-red)">\'+bl+\'</span>\' : bl) : "\\u2014";\n'
        '    tbl += \'<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">\' + (PERIOD_LABELS[p]||p) + ex + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (pg>0?pg.toFixed(1)+"h":"\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ac>0?ac.toFixed(1)+"h":"\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono);color:\' + (d>=0?"var(--sage)":"var(--warm-red)") + \'">\' + dstr + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (tp !== null ? tp : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (vp !== null ? vp : "\\u2014") + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + (ve !== null ? ve : "\\u2014") + \'</td>\'\n'
        '         + \'<td>\' + cpCell + \'</td>\'\n'
        '         + \'<td style="font-family:var(--mono)">\' + blCell + \'</td>\'\n'
        '         + \'</tr>\';\n'
        '  });\n'
        '  if (excl.length) tbl += \'<tr><td colspan="9" style="font-size:9px;color:var(--muted);font-family:var(--mono);padding:6px 10px">\\u2731 Excluido del promedio comparativo</td></tr>\';\n'
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
        '    \'<div><div class="prog-head"><span class="prog-lbl">Proyecci\\u00f3n a 12 meses</span><span class="prog-num" style="color:var(--sand-500)">~\' + proj.toFixed(0) + \'h est.</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + pctP + \'%;background:linear-gradient(90deg,var(--clay),var(--sand-400))"></div></div><div class="prog-note">Prom. \' + avgM.toFixed(1) + \'h/mes en meses activos</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Descanso promedio</span><span class="prog-num" style="color:var(--sage)">\' + avgL.toFixed(1) + \' d/mes</span></div><div class="prog-track"><div class="prog-fill" style="width:\' + Math.min(avgL/20*100,100) + \'%;background:var(--sage)"></div></div><div class="prog-note">M\\u00ednimo reglamentario DAN 121: 8 d\\u00edas/mes</div></div>\' +\n'
        '    \'<div><div class="prog-head"><span class="prog-lbl">Meses activos</span><span class="prog-num">\' + actP.length + \' / \' + PERIODS.length + \'</span></div><div style="display:flex;gap:3px;margin-top:4px"><div style="height:5px;border-radius:2px 0 0 2px;background:var(--sage);flex:\' + actP.length + \'"></div><div style="height:5px;border-radius:0 2px 2px 0;background:var(--rust);opacity:.5;flex:\' + Math.max(PERIODS.length-actP.length,0) + \'"></div></div><div class="prog-note">\' + (excl.length?excl.map(p=>PERIOD_LABELS[p]||p).join(", ")+" excluidos":"Sin ausencias prolongadas") + \'</div></div>\';\n'
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
        '    <div class="brand">\n'
        '      <div class="brand-icon"><svg fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/></svg></div>\n'
        '      <span class="brand-name">SPSKY</span>\n'
        '    </div>\n'
        '    <div class="brand-sub" style="margin-left:38px">Digital Copilot</div>\n'
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
