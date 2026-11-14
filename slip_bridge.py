"""Punch ID -> attendance + leave balance slip on a network thermal printer.
57 mm / 80 mm thermal paper. Works with ZKTeco network devices (LP400, UFace 900, ...) that allow live punch capture.

Setup : pip install pyzk            (xlrd only if you ever get an old binary .xls)
Folder: put the two reports in the reports folder (or in the linked folder). Rules:
          - the file name must contain  WH  (attendance)  or  LB  (leave balance)  as a separate word
            (WH.xls, WH 05 Nov.xlsx, Updated_LB_report.csv ... anything else in the name is fine)
          - the file can be the HTML-table .xls export, .xlsx, .xlsm or .csv - it is read as it is, no manual conversion
Source: the browser extension downloads 'WH Report' / 'LB Report' from Decibel every 5 minutes into the 'source' folder next to the exe;
        the app can also log in to Decibel itself and download them (see DL below); it copies them from there and its Status Bar shows when WH / LB / Salary were last updated (tick or cross).
Daily : drop the new WH and LB reports in the folder. They are picked up automatically within about 30 seconds,
        no restart needed. The newest file of each kind wins. If a new file is unreadable (half-copied, wrong format)
        the old data stays in use.
Test  : python slip_bridge.py --test 401708     (prints slip on screen, no hardware needed)
Live  : python slip_bridge.py                   (listens to the device and prints slips)
"""
import csv, glob, html, http.cookiejar, io, itertools, json, os, re, shutil, socket, sys, threading, time, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timedelta

DEVICE_IP, DEVICE_PORT, COMM_KEY = '192.168.1.201', 4370, 0
PRINTER_IP, PRINTER_PORT = '192.168.1.250', 9100   # network thermal printer (raw ESC/POS)
COLS, DAYS = 32, 30                                 # 57 mm paper: 32 | 80 mm paper: 40 | A4: 48
DUPLICATE_SECONDS = 600                             # same person ignored for 10 minutes
SOURCE_DIR = ''                                     # source folder the downloaded reports are copied from (checked every 5 s); the window app sets it to the 'source' folder next to the exe
REPORT_NAMES = {'WH': 'WH Report', 'LB': 'LB Report'}   # the browser extension saves its downloads under these names (any extension), replacing the old file every 5 min
REFRESH_MIN = 5                                     # the extension re-downloads WH + LB every 5 minutes
STALE_MIN = 12                                      # a report not refreshed for this many minutes shows a red cross in the Status Bar (2 missed downloads + margin)
PORTAL = {'site': 'https://bpo.mydecibel.com/',     # Decibel portal the extension downloads from (login details are kept in the extension, not here)
          'WH': 'https://mydecibel.com/TMS/OriginateAt02.aspx?AR=none&ReportIndex=846',
          'LB': 'https://mydecibel.com/TMS/OriginateAt02.aspx?AR=none&ReportIndex=39'}
PRINT_MODE = 'raw'                                  # 'raw' = ESC/POS thermal printer | 'text' = normal text via the Windows driver (any printer)
PRINTER_NAME = ''                                  # USB / Windows printer name; empty = use the network IP above
SALARY_PUNCH_STATE = 4                              # punch state (the machine's function key) that prints the SALARY slip
ATTENDANCE_PUNCH_STATES = None                      # None = every other state (3, 0, 1 ...) prints the attendance + leave balance slip; {3} = only state 3 does (others are ignored)
PREVIEW_ONLY = False                                # True (set by the window app, password protected): a punch only shows the attendance + salary slip, nothing prints
PAGE_FONT_PT, PAGE_MARGIN_IN = 12, 0.5                  # text printing on a wide page (A4 / Letter on a Windows printer): font size and margin; narrow rolls use 9 pt and no margin
PRINT_RED = True                                    # Absent / Fullday Ded. lines in red: two-colour ESC/POS paper, or colour on a Windows driver (a one-colour printer just prints black). False = never send the colour command
MACHINES = []                                       # set by the window app: [{'name', 'ip', 'port', 'key', 'printer_ip'}] one entry per attendance machine; each runs live() in its own thread
ID_MAP = {}                                         # only if device IDs differ: {'1': '401708'}
SHOW_SHIFT = True                                   # Shift (roster code) column on the attendance slip
SHOW_OT = True                                      # 'Calculated Overtime' column on the attendance slip (just before Remarks)
FEED_LINES = 3                                      # line feeds after the slip before the cut (less = less wasted paper; if the last line is cut off, make it 4)
TIGHT_LINES = True                                  # True = ESC 3 tighter line spacing (about 13% shorter slip). If lines overlap on your printer set False
LINE_DOTS = 26                                      # line pitch in dots when TIGHT_LINES (default printer pitch is 30)
SAL_FIELDS = (('name', 'Name'), ('id', 'ID'), ('dept', 'Dept'), ('pos', 'Post'))   # details printed at the top of the salary slip (more: 'grade', 'type', 'doj', 'bank')


def flags(rem):   # L = Late, E = Early going, H = Half day
    r = str(rem).lower()
    return ' '.join(c for c, k in (('L', 'late'), ('E', 'early'), ('H', 'half')) if k in r)


LEAVE_RULES = [   # order matters: specific first, general last. Only these 17 leave types (+ 'Off Day') are shown on the slip.
    (r'half\s*day\s*casual|halfday.*casual', 'Half Day Casual'),
    (r'short\s*casual', 'Short Casual Leave'),
    (r'casual', 'Casual Leave'),
    (r'half\s*day\s*sick|halfday.*sick', 'Half Day Sick'),
    (r'short\s*sick', 'Short Sick Leave'),
    (r'sick', 'Sick Leave'),
    (r'annual', 'Annual Leave'),
    (r'compensatory', 'CP Leave'),
    (r'leave\s*without\s*pay', 'Leave Without Pay'),
    (r'social\s*security', 'Social Security Leave'),
    (r'off?ic(i)?al\s*visit', 'Official Visit'),
    (r'paternity', 'Paternity Leave'),
    (r'maternity', 'Maternity Leave'),
    (r'hajj|pilgrimage', 'Hajj / Pilgrimage'),
    (r'suspended', 'Suspended'),
    (r'iddat', 'Iddat Leave'),
    (r'lay\s*off', 'Lay Off Leave'),
]
# Words in the WH remark that make the day an 'Off Day' on the slip. Public / gazetted holidays (Eid Miladun Nabi ...) are NOT here: their remark stays as it was.
OFF_DAY_WORDS = ('off[\\s_-]*day', 'sunday', 'off[\\s_-]*saturday', 'weekly[\\s_-]*off', 'rest[\\s_-]*day')
OFF_DAY = r'\b(?:' + '|'.join(OFF_DAY_WORDS) + r')\b'
ABSENT = r'\babsent\b'
FULLDAY_DED = r'full[\s_-]*day[\s_-]*ded'
RED_REMARKS = ('Absent', 'Fullday Ded.')       # printed / shown in red


def clean_remark(text):
    """WH remark -> 'Off Day', one of the 17 leave types, 'Off Day-(Leave type)' when both are present, 'Absent', 'Fullday Ded.' or ''.
    e.g. 'Sunday' / 'Off Saturday On Time-Overtime' -> 'Off Day' | '(Half Day Casual)-On Time' -> 'Half Day Casual' | 'Off Saturday Compensatory Leave' -> 'Off Day-(CP Leave)'."""
    s = str(text or '').lower()
    leave = next((name for pat, name in LEAVE_RULES if re.search(pat, s)), '')
    if re.search(OFF_DAY, s):
        return f'Off Day-({leave})' if leave else 'Off Day'
    if leave:
        return leave
    if re.search(ABSENT, s):
        return 'Absent'
    return 'Fullday Ded.' if re.search(FULLDAY_DED, s) else ''


BOLD_STARTS = ('Total Earnings', 'Total Deductions', 'Total Provident Fund')


def is_big(line):
    """The NET SALARY line of the salary slip: printed / shown in a double size font."""
    return bool(re.match(r'^NET( SALARY)? ', line))


def is_bold(line):
    """Totals of the salary slip (bold); the big NET SALARY line is bold too."""
    return line.startswith(BOLD_STARTS) or is_big(line)


def is_red(line):
    """True for an attendance line of the slip whose remark is Absent or Fullday Ded. (shown / printed in red)."""
    return bool(re.match(r'^\d{2}-[A-Za-z]{3} ', line)) and line.rstrip().endswith(RED_REMARKS)


def fmt_ot(v):
    """Calculated overtime as short text: '01:30:00' -> '1:30', '1.50' -> '1.5', zero / empty -> ''."""
    s = str(v if v is not None else '').strip()
    if not s or s.lower() in ('nan', 'none', '-', '--'):
        return ''
    m = re.fullmatch(r'(\d{1,3}):(\d{2})(?::\d{2})?', s)
    if m:
        return '' if int(m.group(1)) == 0 and m.group(2) == '00' else f'{int(m.group(1))}:{m.group(2)}'
    try:
        x = float(s.replace(',', ''))
    except ValueError:
        return s[:5]
    return '' if x == 0 else f'{x:.2f}'.rstrip('0').rstrip('.')


def shift_code(v):
    """Roster code as short text: None -> '', 3.0 -> '3'."""
    if v is None or (isinstance(v, float) and v != v):
        return ''
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return str(v).strip()


SLIP_HOOK = None   # set by the window app: SLIP_HOOK(uid, text, state, error) with state 'show' | 'ok' | 'err' | 'notfound' (attendance slip), 'sshow' | 'sok' | 'serr' | 'snotfound' (salary slip) or 'pshow' (preview mode: text = (attendance slip, salary slip))
STOP = threading.Event()   # set by the window app to stop live()
STATS = {'error': '', 'loading': False, 'connected': False, 'slips': 0, 'last': '-', 'printer': '-', 'machines': {}}   # shown on the app overview
LOADED = {}                # {'WH': file name, 'LB': file name} of the reports in use


# ======================================================================================
#  Finding the report files.  A file is a WH / LB report when its name contains WH / LB as a separate word.
# ======================================================================================
REPORT_EXT = ('.xls', '.xlsx', '.xlsm', '.csv', '.tsv', '.txt', '.htm', '.html')
_MON = {m: i for i, m in enumerate(('jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'), 1)}


def kind_of(name):
    """'WH' or 'LB' when the file name has that word in it (any case, anywhere), else None."""
    base = os.path.basename(name)
    stem, ext = os.path.splitext(base)
    if base.startswith(('~$', '.')) or ext.lower() not in REPORT_EXT:
        return None
    low = stem.lower()
    wh, lb = re.search(r'(?<![a-z])wh(?![a-z])', low), re.search(r'(?<![a-z])lb(?![a-z])', low)
    return 'WH' if wh and not lb else 'LB' if lb and not wh else None


def fdate(p):
    """Date written in the file name ('05 Nov 2026', '05_Nov__2026' ...) or None."""
    m = re.search(r'(\d{1,2})[\s_\-.]+([A-Za-z]{3,9})[\s_,.\-]*(\d{4})', os.path.basename(p))
    try:
        return datetime(int(m.group(3)), _MON[m.group(2)[:3].lower()], int(m.group(1)))
    except Exception:
        return None


def rank(p):
    """Newest first: the date in the file name when there is one, else the file's modified time."""
    try:
        mt = os.path.getmtime(p)
    except OSError:
        mt = 0
    return (fdate(p) or datetime.fromtimestamp(mt), mt)


def candidates(folder, kind):
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    return [os.path.join(folder, n) for n in names if kind_of(n) == kind and os.path.isfile(os.path.join(folder, n))]


def newest(kind, folder='.'):
    files = candidates(folder, kind)
    if not files:
        raise FileNotFoundError(f'No {kind} report found (the file name must contain {kind})')
    return max(files, key=rank)


def report_dates():
    out = {}
    for k in ('WH', 'LB'):
        try:
            out[k] = rank(newest(k))[0]
        except Exception:
            pass
    return out


def sig():
    return [[p, os.path.getmtime(p), os.path.getsize(p)] for p in (newest('WH'), newest('LB'))]


_obs, _obs_lock = {}, threading.Lock()


def _stable(p, wait=3.0):
    """True when the file has not changed (size / time) for `wait` seconds: it is fully copied."""
    try:
        st = os.stat(p)
    except OSError:
        return False
    cur, now = (st.st_size, st.st_mtime_ns), time.time()
    with _obs_lock:
        o = _obs.get(p)
        if o is None or o[0] != cur:
            _obs[p] = (cur, now)
            return False
        return now - o[1] >= wait


def _mark_stable(p):
    """A file this program copied and checked itself is complete: no need to wait for it again."""
    try:
        st = os.stat(p)
    except OSError:
        return
    with _obs_lock:
        _obs[p] = ((st.st_size, st.st_mtime_ns), time.time() - 60)


# ======================================================================================
#  Reading any report format into plain rows (lists of text):  HTML-table .xls | .xlsx | .csv | old binary .xls
# ======================================================================================
_ROW = re.compile(r'<tr[^>]*>(.*?)</tr>', re.S | re.I)
_CELL = re.compile(r'<t[dh][^>]*>(.*?)</t[dh]>', re.S | re.I)
_TAG = re.compile(r'<[^>]+>')


def _enc(head):
    return 'utf-16' if head[:2] in (b'\xff\xfe', b'\xfe\xff') else 'utf-8-sig'


def html_rows(path, enc='utf-8-sig'):
    """Stream the rows of an HTML-table '.xls' export (e.g. Decibel reports) as lists of cell text."""
    buf, first = '', True
    with open(path, encoding=enc, errors='replace') as f:
        while True:
            chunk = f.read(4 << 20)
            if not chunk:
                break
            if first and re.search(r'<frameset|Excel Workbook Frameset', chunk[:20000], re.I):
                raise ValueError('web-page style .xls (needs its _files folder): open it in Excel and save as .xlsx')
            first, buf = False, buf + chunk
            cut = buf.lower().rfind('</tr>')
            if cut < 0:
                continue
            part, buf = buf[:cut + 5], buf[cut + 5:]
            for m in _ROW.finditer(part):
                yield [html.unescape(_TAG.sub('', c)).replace('\xa0', ' ').strip() for c in _CELL.findall(m.group(1))]


def csv_rows(path, enc='utf-8-sig'):
    try:
        with open(path, encoding=enc) as f:
            sample = f.read(1 << 16)
    except UnicodeDecodeError:
        enc = 'cp1252'
        with open(path, encoding=enc, errors='replace') as f:
            sample = f.read(1 << 16)
    first = sample.split('\n', 1)[0]
    delim = max(',;\t|', key=lambda d: first.count(d))
    if first.count(delim) == 0:
        delim = ','
    with open(path, encoding=enc, errors='replace', newline='') as f:
        for r in csv.reader(f, delimiter=delim):
            yield [c.strip() for c in r]


_DATE_BUILTIN = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))
_EPOCH = datetime(1899, 12, 30)


def _is_date_fmt(code):
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', '', code)
    return bool(re.search(r'[ymdhs]', code, re.I))


def _serial(v):
    """Excel date/time number -> text: 2026-11-05 | 07:45:00 | 2026-11-05 07:45:00."""
    r = _sercache.get(v)
    if r is None:
        r = _sercache[v] = _serial_calc(v)
    return r


def _serial_calc(v):
    try:
        x = float(v)
        dt = _EPOCH + timedelta(seconds=round(x * 86400))
    except Exception:
        return v
    if x < 1:
        return dt.strftime('%H:%M:%S')
    return dt.strftime('%Y-%m-%d') if x == int(x) else dt.strftime('%Y-%m-%d %H:%M:%S')


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
        else:
            break
    return n - 1


_CELL_X = re.compile(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(?:<f\b[^>]*/>|<f\b[^>]*>.*?</f>)?(?:<v>([^<]*)</v>|<is>(.*?)</is>)?</c>)', re.S)
_A_T, _A_S = re.compile(r'\bt="(\w+)"'), re.compile(r'\bs="(\d+)"')
_T = re.compile(r'<t\b[^>]*>(.*?)</t>', re.S)
_colcache, _sercache = {}, {}


def xlsx_rows(path):
    """Fast streaming reader for .xlsx / .xlsm (first sheet). Dates and times come out as text."""
    import zipfile, codecs
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    ss = []
    if 'xl/sharedStrings.xml' in names:
        data = re.sub(r'<si\s*/>', '<si></si>', z.read('xl/sharedStrings.xml').decode('utf-8', 'replace'))
        for m in re.finditer(r'<si\b[^>]*>(.*?)</si>', data, re.S):
            ss.append(html.unescape(''.join(_T.findall(re.sub(r'<rPh\b.*?</rPh>', '', m.group(1), flags=re.S)))).strip())
        del data
    isdate = []                                                    # per cell style: is it a date / time format?
    if 'xl/styles.xml' in names:
        st = z.read('xl/styles.xml').decode('utf-8', 'replace')
        custom = {}
        for tag in re.findall(r'<numFmt\b[^>]*>', st):
            i, c = re.search(r'numFmtId="(\d+)"', tag), re.search(r'formatCode="([^"]*)"', tag)
            if i and c:
                custom[int(i.group(1))] = html.unescape(c.group(1))
        blk = re.search(r'<cellXfs\b[^>]*>(.*?)</cellXfs>', st, re.S)
        for tag in re.findall(r'<xf\b[^>]*>', blk.group(1) if blk else ''):
            m = re.search(r'numFmtId="(\d+)"', tag)
            n = int(m.group(1)) if m else 0
            isdate.append(n in _DATE_BUILTIN or (n in custom and _is_date_fmt(custom[n])))
    sheet = None
    try:
        rid = re.search(r'<sheet\b[^>]*\br:id="([^"]+)"', z.read('xl/workbook.xml').decode('utf-8', 'replace')).group(1)
        for tag in re.findall(r'<Relationship\b[^>]*>', z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'replace')):
            if f'Id="{rid}"' in tag:
                tgt = re.search(r'Target="([^"]+)"', tag).group(1)
                sheet = tgt.lstrip('/') if tgt.startswith('/') else 'xl/' + tgt
    except Exception:
        sheet = None
    if sheet not in names:
        sheets = sorted((n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)), key=lambda n: int(re.search(r'(\d+)\.xml', n).group(1)))
        if not sheets:
            raise ValueError('no worksheet found in this .xlsx file')
        sheet = sheets[0]
    dec, buf, acache = codecs.getincrementaldecoder('utf-8')('replace'), '', {}
    with z.open(sheet) as f:
        while True:
            raw = f.read(4 << 20)
            buf += dec.decode(raw, final=not raw)
            cut = buf.rfind('</row>')
            if not raw or cut >= 0:
                part, buf = (buf, '') if not raw else (buf[:cut + 6], buf[cut + 6:])
                for rm in re.finditer(r'<row\b[^>]*?(?<!/)>(.*?)</row>', part, re.S):
                    got, top = [], -1
                    for col, attrs, v, inl in _CELL_X.findall(rm.group(1)):
                        if not v and not inl:
                            continue
                        k = acache.get(attrs)
                        if k is None:
                            t, st = _A_T.search(attrs), _A_S.search(attrs)
                            si = int(st.group(1)) if st else 0
                            k = acache[attrs] = (t.group(1) if t else 'n', bool(isdate[si]) if si < len(isdate) else False)
                        t, dflag = k
                        if t == 's':
                            val = ss[int(v)] if v.isdigit() and int(v) < len(ss) else ''
                        elif t == 'inlineStr' or (inl and not v):
                            val = html.unescape(''.join(_T.findall(inl))).strip()
                        elif t == 'e':
                            continue
                        elif t == 'str':
                            val = html.unescape(v).strip()
                        elif t == 'b':
                            val = 'TRUE' if v == '1' else 'FALSE'
                        elif dflag:
                            val = _serial(v)
                        else:
                            val = v
                        c = _colcache.get(col)
                        if c is None:
                            c = _colcache[col] = _col_index(col)
                        got.append((c, val))
                        if c > top:
                            top = c
                    if top >= 0:
                        row = [''] * (top + 1)
                        for c, val in got:
                            row[c] = val
                        yield row
            if not raw:
                break


def xls_rows(path):
    """Old binary .xls (needs the xlrd library)."""
    try:
        import xlrd
    except ImportError:
        raise ValueError('old binary .xls needs the xlrd library: open the file in Excel and save as .xlsx (or .csv)')
    book = xlrd.open_workbook(path, on_demand=True)
    sh = book.sheet_by_index(0)
    for i in range(sh.nrows):
        row = []
        for c in sh.row(i):
            if c.ctype == xlrd.XL_CELL_DATE:
                row.append(_serial(c.value))
            elif c.ctype == xlrd.XL_CELL_NUMBER:
                row.append(str(int(c.value)) if c.value == int(c.value) else repr(c.value))
            elif c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                row.append('')
            else:
                row.append(str(c.value).strip())
        yield row


def iter_rows(path):
    """Look at what is really inside the file (not at its extension) and read it."""
    with open(path, 'rb') as f:
        head = f.read(4096)
    if head[:2] == b'PK':
        return xlsx_rows(path)
    if head[:4] == b'\xd0\xcf\x11\xe0':
        return xls_rows(path)
    enc = _enc(head)
    probe = head.decode(enc if enc == 'utf-16' else 'utf-8', 'ignore').lstrip('\ufeff \t\r\n\x00').lower()
    if probe.startswith(('<html', '<!doctype', '<table', '<?xml', '<meta', '<head', '<body')) or '<table' in probe[:3000]:
        return html_rows(path, enc)
    return csv_rows(path, enc)


# ======================================================================================
#  Turning the rows into the tables the slip needs
# ======================================================================================
def _norm(s):
    return re.sub(r'\s+', ' ', str(s)).strip().lower()


WH_NEED = ('attendance date', 'emp id', 'employee name', 'emp in', 'emp out', 'remarks')


def wh_table(rows):
    """-> [(date text, emp id, name, in, out, remarks, roster code, calculated overtime)]  (columns are found by their header names)."""
    idx = None
    out = []
    for i, r in enumerate(rows):
        if idx is None:
            low = [_norm(x) for x in r]
            if 'emp id' in low and 'attendance date' in low:
                miss = [c for c in WH_NEED if c not in low]
                if miss:
                    raise ValueError('missing columns in WH report: ' + ', '.join(miss))
                ros = low.index('roster code') if 'roster code' in low else next((j for j, x in enumerate(low) if 'roster' in x), -1)
                ot = next((j for j, x in enumerate(low) if 'overtime' in x and 'calc' in x), -1)        # 'Calculated Overtime' (optional column)
                if ot < 0:
                    ot = next((j for j, x in enumerate(low) if 'overtime' in x or x in ('ot', 'ot hours')), -1)
                idx = [low.index(c) for c in WH_NEED] + [ros, ot]
            elif i > 120:
                raise ValueError('WH report header row (Attendance Date, Emp Id ...) not found')
            continue
        if len(r) <= max(idx[0], idx[1]):                 # (xlsx rows may be shorter: trailing empty cells are not stored)
            continue
        out.append(tuple(r[j] if 0 <= j < len(r) else '' for j in idx))
    if idx is None:
        raise ValueError('WH report header row (Attendance Date, Emp Id ...) not found')
    return out


def _lb_cols(top, sub, nd):
    """Positions of: emp id, name, and (allowed, availed, closing) for Casual, Sick, Annual leave."""
    t = [_norm(x) for x in top]
    idc = t.index('emp id')
    nmc = t.index('employee name') if 'employee name' in t else max(0, idc - 1)
    s = [_norm(x) for x in sub]
    blocks = []                                           # (label, allowed col, availed col, closing col)
    if any(s):
        nd = max(nd, len(s))
        off = nd - len(s) if len(s) < nd else 0
        label = [''] * nd
        if off > 0:                                       # HTML export: group names are not aligned with the columns
            groups = t[off:]
            span = max(1, (nd - off) // max(1, len(groups)))
            for c in range(nd):
                label[c] = t[c] if c < off else (groups[min((c - off) // span, len(groups) - 1)] if groups else '')
        else:                                             # merged cells: a group name covers the blank cells after it
            cur = ''
            for c in range(nd):
                cur = t[c] if c < len(t) and t[c] else cur
                label[c] = cur
        for j, x in enumerate(s):
            if x == 'allowed':
                a = j + off
                blk = {k: next((j + off + q for q in range(4) if j + q < len(s) and s[j + q] == k), None) for k in ('availed', 'closing')}
                blocks.append((label[a] if a < nd else '', a, blk['availed'] if blk['availed'] is not None else a + 1, blk['closing'] if blk['closing'] is not None else a + 3))
    else:                                                 # one header row such as 'Casual Leave Allowed'
        for lv in ('casual', 'sick', 'annual'):
            f = lambda w: next((j for j, x in enumerate(t) if lv in x and w in x), None)
            a, u, c = f('allowed'), f('availed'), f('closing')
            if a is not None and u is not None and c is not None:
                blocks.append((lv, a, u, c))
    by = {}
    for lab, a, u, c in blocks:
        for lv in ('casual', 'sick', 'annual'):
            if lv in lab and lv not in by:
                by[lv] = (a, u, c)
    if len(by) < 3:
        if len(blocks) >= 3:
            by = {lv: blocks[i][1:] for i, lv in enumerate(('casual', 'sick', 'annual'))}
        elif nd >= 18:                                    # the standard layout, by position
            by = {'casual': (6, 7, 9), 'sick': (10, 11, 13), 'annual': (14, 15, 17)}
        else:
            raise ValueError('unexpected leave balance layout (Casual / Sick / Annual Allowed, Availed, Closing not found)')
    return idc, nmc, [x for lv in ('casual', 'sick', 'annual') for x in by[lv]]


def _num(x):
    try:
        return round(float(str(x).replace(',', '').strip() or 0), 2)
    except ValueError:
        return 0.0


def lb_table(rows):
    """-> [(emp id, name, [9 numbers: Casual allowed/used/left, Sick ..., Annual ...])]"""
    top = sub = cols = None
    out = []
    for i, r in enumerate(rows):
        if top is None:
            if 'emp id' in [_norm(x) for x in r]:
                top = r
            elif i > 120:
                raise ValueError('LB report header row (Emp Id ...) not found')
            continue
        if sub is None:
            if any(_norm(x) in ('allowed', 'availed', 'closing') for x in r):
                sub = r
                continue
            sub = []
        if cols is None:
            cols = _lb_cols(top, sub, len(r))
        idc, nmc, pos = cols
        if len(r) <= idc or not re.search(r'\d', r[idc]):
            continue
        out.append((r[idc], r[nmc] if nmc < len(r) else '', [_num(r[p]) if p < len(r) else 0.0 for p in pos]))
    if top is None:
        raise ValueError('LB report header row (Emp Id ...) not found')
    return out


def read_report(path, kind, quick=False):
    rows = iter_rows(path)
    if quick:
        rows = itertools.islice(rows, 400)
    t = (wh_table if kind == 'WH' else lb_table)(rows)
    if quick and not t:
        raise ValueError(f'no data rows found in the {kind} report')
    return t


def eid(v):
    """Employee ID as plain text, letters removed: 219033.0 -> '219033', '219033x' -> '219033', '906137 (906137)' -> '906137'."""
    s = str(v).strip()
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    return re.sub(r'^[A-Za-z]+|[A-Za-z]+$', '', re.sub(r'\s*\(.*\)$', '', s))


def hm(v):
    """Any time value (time, datetime, text like '8:45' or '08:45:10', Excel fraction) -> 'HH:MM' or ''."""
    if v is None or (isinstance(v, float) and v != v) or v == '':
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%H:%M')
    s = str(v).strip()
    if s.lower() in ('', 'nan', 'nat', 'none'):
        return ''
    if re.fullmatch(r'0?\.\d+', s) or isinstance(v, (int, float)) and 0 <= v < 1:
        m = round(float(s) * 1440)
        return f'{m // 60 % 24:02d}:{m % 60:02d}'
    m = re.search(r'(?<!\d)(\d{1,2}):(\d{2})', s)
    return f'{int(m.group(1)):02d}:{m.group(2)}' if m else s[:5]


def parse_date(s):
    """'26 Aug 2026' | '2026-08-26' | '26/08/2026' (day first) -> datetime, or None."""
    s = str(s).strip()
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    try:
        if m:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.match(r'^(\d{1,2})[\s\-/.,]+([A-Za-z]{3,9})[\s\-/.,]*(\d{2,4})', s)
        if m:
            y = int(m.group(3))
            return datetime(y + 2000 if y < 100 else y, _MON[m.group(2)[:3].lower()], int(m.group(1)))
        m = re.match(r'^(\d{1,2})[\-/.](\d{1,2})[\-/.](\d{2,4})', s)
        if m:
            y = int(m.group(3))
            return datetime(y + 2000 if y < 100 else y, int(m.group(2)), int(m.group(1)))
    except (ValueError, KeyError):
        return None
    return None


def load():
    wp, lp = newest('WH'), newest('LB')
    lb, wh = read_report(lp, 'LB'), read_report(wp, 'WH')
    d, dups = {}, []
    for uid, n, num in lb:
        uid = eid(uid)
        if uid in d:
            dups.append(uid)
        d[uid] = {'n': n, 'att': [], 'lv': {'Casual': num[0:3], 'Sick': num[3:6], 'Annual': num[6:9]}}
    seen, rows = {}, []
    for ds, emp, n, i, o, r, sh, ot in wh:
        if not str(emp).strip():
            continue
        if ds not in seen:
            seen[ds] = parse_date(ds)
        if seen[ds] is not None:
            rows.append((seen[ds], emp, n, i, o, r, sh, ot))
    rows.sort(key=lambda t: t[0])                         # by date, equal dates keep the file order
    _last_span[0] = (rows[-1][0] - rows[0][0]).days + 1 if rows else 0
    for dt, emp, n, i, o, r, sh, ot in rows:
        uid = eid(emp)
        d.setdefault(uid, {'n': n, 'att': [], 'lv': {}})['att'].append([dt.strftime('%d-%b'), hm(i), hm(o), r, flags(r), shift_code(sh), clean_remark(r), fmt_ot(ot)])
    if dups:
        print(f'Warning: {len(dups)} duplicate employee IDs after removing letters, e.g. {dups[:3]} (the last one is used)')
    return d


DATA, today, SIG = {}, {}, None   # today = {id: [first punch, last punch]} (resets if program restarts)
SPAN = [0]                        # days the loaded WH report covers: the slip always shows at least all of them (DAYS is the minimum)
_last_span = [0]
MAX_SPAN = 62                     # never print more than two months, whatever the report holds


CACHE, BAD, _lock, _last_err = 'cache.json', None, threading.Lock(), ['']


def refresh():
    """Reload the reports when newer / changed files appear. Parsed data is cached (restarts are instant);
    old data is kept while the new file is still being copied or if it fails, and a failed file is not retried until it changes."""
    global DATA, SIG, BAD
    if not _lock.acquire(blocking=False):
        return
    s = None
    try:
        s = sig()
        if s == SIG or s == BAD:
            return
        key = s + [['v', 11]]                              # bump v when the parsing rules change
        try:
            c = json.load(open(CACHE, encoding='utf-8'))
            d, span = (c['data'], c.get('span', 0)) if c['sig'] == key else (None, 0)
        except Exception:
            d, span = None, 0
        if d is None:
            ok = [_stable(p) for p, _, _ in s]
            if not all(ok):
                return                                    # still being copied: look again at the next check
            STATS['loading'] = True
            t0 = time.time()
            d = load()
            span = _last_span[0]
            try:
                json.dump({'sig': key, 'data': d, 'span': span}, open(CACHE + '.t', 'w', encoding='utf-8'))
                os.replace(CACHE + '.t', CACHE)
            except Exception:
                pass
            print(f'  reports read in {time.time() - t0:.1f} s')
        if not d:
            raise ValueError('no employees found in the reports')
        DATA, SIG, SPAN[0] = d, s, span
        LOADED.update({'WH': os.path.basename(s[0][0]), 'LB': os.path.basename(s[1][0])})
        STATS['error'], _last_err[0] = '', ''
        print('Reports loaded:', [os.path.basename(p) for p, _, _ in s], f'({len(d)} employees)')
    except Exception as ex:
        if s:
            BAD = s
        STATS['error'] = str(ex)[:90]
        if _last_err[0] != str(ex):
            _last_err[0] = str(ex)
            print('Report reload failed (old data kept):', ex)
    finally:
        STATS['loading'] = False
        _lock.release()


SYNC, _sync_lock = {'last': '-', 'result': 'Linked folder not set'}, threading.Lock()
SYNC_ERR = {'WH': '', 'LB': ''}      # last problem while copying / reading each report from the linked folder ('' = none)
STATUS = {k: {'time': '-', 'ok': False, 'note': 'not checked yet'} for k in ('WH', 'LB', 'SAL')}   # Status Bar rows: when each was last updated + tick / cross
STATUS_ROWS = (('WH', 'Work Hour Report'), ('LB', 'Leave Balance Report'), ('SAL', 'Salary Slip'))


def _newest_file(kind):
    """Newest WH / LB file of the linked folder (the extension's download); falls back to the reports folder. -> (path, folder_problem)."""
    problem = ''
    if SOURCE_DIR:
        if os.path.isdir(SOURCE_DIR):
            files = candidates(SOURCE_DIR, kind)
            if files:
                return max(files, key=rank), ''
            problem = f'no {kind} file in {SOURCE_DIR}'
        else:
            problem = f'folder {SOURCE_DIR} not reachable'
    files = candidates('.', kind)
    return (max(files, key=rank) if files else None), problem


def update_status():
    """Refresh the three Status Bar rows (called every few seconds in the background, so the window never waits on the network folder).
    WH / LB: time = when the newest file was written (the extension's last download); tick = fresh (within STALE_MIN), no copy / read error, and loaded.
    Salary : time = when the salary PDFs were last read; tick = salary slips are loaded without error."""
    now = time.time()
    for k in ('WH', 'LB'):
        try:
            p, problem = _newest_file(k)
            if not p:
                STATUS[k].update(time='-', ok=False, note=problem or f'no {k} report found'); continue
            mt = os.path.getmtime(p)
            age = (now - mt) / 60
            err = (DL_ERR.get(k) if DL['on'] else '') or SYNC_ERR.get(k) or (STATS['error'] if STATS['error'] and k in LOADED else '')
            ok = age <= STALE_MIN and not err and not problem.startswith('folder') and k in LOADED
            note = err or (f'older than {STALE_MIN} min - the extension has not downloaded a new file' if age > STALE_MIN else problem or 'updated')
            STATUS[k].update(time=datetime.fromtimestamp(mt).strftime('%d %b %Y %H:%M'), ok=bool(ok), note=note)
        except Exception as ex:
            STATUS[k].update(ok=False, note=str(ex)[:60])
    try:
        t = SAL['last'] if SAL['last'] != '-' else '-'
        if t == '-':
            pdfs = [os.path.join(f, n) for f in dict.fromkeys(x for x in ('.', SOURCE_DIR) if x) if os.path.isdir(f)
                    for n in os.listdir(f) if n.lower().endswith('.pdf') and not n.startswith(('~$', '.'))]
            if pdfs:
                t = datetime.fromtimestamp(max(os.path.getmtime(x) for x in pdfs)).strftime('%d %b %Y %H:%M')
        ok = SAL['n'] > 0 and not SAL['error']
        STATUS['SAL'].update(time=t, ok=bool(ok), note=SAL['error'] or ('updated' if ok else 'no salary PDF loaded'))
    except Exception as ex:
        STATUS['SAL'].update(ok=False, note=str(ex)[:60])




def check_file(path, kind):
    """Quick sanity check of the first rows; raises ValueError with a readable message."""
    try:
        read_report(path, kind, quick=True)
    except ValueError:
        raise
    except Exception as ex:
        raise ValueError(f'cannot read file ({ex})')


def prune():
    """Keep only the 3 newest local copies of each report."""
    for k in ('WH', 'LB'):
        for p in sorted(candidates('.', k), key=rank, reverse=True)[3:]:
            try:
                os.remove(p)
            except OSError:
                pass


def sync_reports():
    """Copy the newest WH / LB file of the linked folder into the reports folder when it is new or changed."""
    if not _sync_lock.acquire(blocking=False):
        return
    msg, changed = [], False
    try:
        SYNC['last'] = datetime.now().strftime('%d %b %H:%M:%S')
        if not SOURCE_DIR:
            SYNC['result'] = 'Linked folder not set'
            return
        if not os.path.isdir(SOURCE_DIR):
            SYNC['result'] = 'Linked folder not found'
            SYNC_ERR.update(WH='linked folder not reachable', LB='linked folder not reachable')
            return
        for kind in ('WH', 'LB'):
            tmp = None
            try:
                files = candidates(SOURCE_DIR, kind)
                if not files:
                    msg.append(f'{kind}: no file'); SYNC_ERR[kind] = f'no {kind} file in the linked folder'; continue
                src = max(files, key=rank)
                f = os.path.basename(src)
                dst = os.path.join('.', f)
                cur = report_dates().get(kind)
                if cur and rank(src)[0] < cur:
                    msg.append(f'{kind}: {f} is older than the loaded one'); continue
                if os.path.exists(dst) and os.path.getsize(dst) == os.path.getsize(src) and int(os.path.getmtime(dst)) == int(os.path.getmtime(src)):
                    msg.append(f'{kind}: {f} up to date'); SYNC_ERR[kind] = ''; continue
                if not _stable(src):
                    msg.append(f'{kind}: {f} is still being copied, waiting'); continue
                tmp = dst + '.tmp'
                shutil.copy2(src, tmp); check_file(tmp, kind); os.replace(tmp, dst); _mark_stable(dst)
                msg.append(f'{kind}: {f} updated'); changed = True; SYNC_ERR[kind] = ''
            except Exception as ex:
                msg.append(f'{kind}: {str(ex)[:60]}'); SYNC_ERR[kind] = str(ex)[:60]
                if tmp and os.path.exists(tmp):
                    os.remove(tmp)
        SYNC['result'] = '; '.join(msg)
    finally:
        _sync_lock.release()
    if changed:
        prune()
        refresh()


# ======================================================================================
#  Auto-download of the WH / LB reports from Decibel (no browser needed): log in, open the report page, pick the client / service status /
#  dates, tick "Export to Excel", press "Get Report", and save the file as 'WH Report' / 'LB Report' in the source folder (old file replaced).
#  Repeats every DL['minutes'] minutes (adjustable in the window app).
# ======================================================================================
DL = {'on': False, 'minutes': 5, 'user': '', 'pw': '', 'force': False, 'running': False, 'last': '-', 'next': '-', 'result': 'Auto-download not run yet'}
DL_ERR = {'WH': '', 'LB': ''}
CLIENT_WORD, STATUS_WORD, CYCLE_DAY = 'service tyre', 'in service', 26     # client name to pick, service status to pick, the WH period runs from the 26th to the 25th
DEBUG_PAGE = 'download_debug.html'                                          # the last page the site sent when a download failed (open it to see why)


class LoginFailed(ValueError):
    pass


_TAG_RE = re.compile(r'<(/?)(form|input|select|option|textarea)\b((?:"[^"]*"|\'[^\']*\'|[^\'">])*)>', re.I)
_ATTR_RE = re.compile(r'([\w:.-]+)(?:\s*=\s*("[^"]*"|\'[^\']*\'|[^\s"\'>]+))?')


def _attrs(s):
    out = {}
    for m in _ATTR_RE.finditer(s):
        v = m.group(2)
        out[m.group(1).lower()] = html.unescape(v[1:-1] if v and v[0] in '"\'' else v) if v is not None else ''
    return out


def _parse_forms(page):
    """Forms of an HTML page as [{'action', 'els': [...]}] (plain regex, so nothing beyond the standard 're' / 'html' modules is needed)."""
    forms, f, sel, opt = [], None, None, None
    ms = list(_TAG_RE.finditer(page))
    for k, m in enumerate(ms):
        close, tag, a = m.group(1), m.group(2).lower(), _attrs(m.group(3))
        nxt = ms[k + 1].start() if k + 1 < len(ms) else len(page)
        if close:
            if tag == 'select':
                sel = None
            elif tag == 'form':
                f = None
            continue
        if tag == 'form':
            f = {'action': a.get('action', ''), 'els': []}; forms.append(f)
        elif f is None:
            continue
        elif tag == 'input':
            t, v = (a.get('type') or 'text').lower(), a.get('value')
            f['els'].append({'t': t, 'n': a.get('name'), 'v': ('on' if v is None and t in ('checkbox', 'radio') else v or ''), 'c': 'checked' in a, 'dis': 'disabled' in a})
        elif tag == 'select':
            sel = {'t': 'select', 'n': a.get('name'), 'opts': [], 'dis': 'disabled' in a}; f['els'].append(sel)
        elif tag == 'option' and sel is not None:
            text = html.unescape(page[m.end():nxt].split('<', 1)[0])
            sel['opts'].append([a['value'] if 'value' in a else text.strip(), text, 'selected' in a])
        elif tag == 'textarea':
            f['els'].append({'t': 'textarea', 'n': a.get('name'), 'v': html.unescape(page[m.end():nxt].split('<', 1)[0]), 'dis': 'disabled' in a})
    return forms


def _has_login(page):
    return any(e['t'] == 'password' for f in _parse_forms(page) for e in f['els'])


def _payload(form, over=None, button=None):
    """What a browser would send for this form: hidden / text fields, ticked boxes, the selected option of every list; `over` {name: value | None} changes
    or removes fields; `button` = name of the submit button that was pressed."""
    out = []
    for e in form['els']:
        n, t = e.get('n'), e['t']
        if not n or e.get('dis'):
            continue
        if t in ('submit', 'button', 'image', 'reset', 'file'):
            if t in ('submit', 'image') and n == button:
                out += [(n + '.x', '1'), (n + '.y', '1')] if t == 'image' else [(n, e['v'])]
            continue
        if t in ('checkbox', 'radio'):
            if e['c']:
                out.append((n, e['v']))
        elif t == 'select':
            sel = next((o for o in e['opts'] if o[2]), e['opts'][0] if e['opts'] else None)
            if sel:
                out.append((n, sel[0]))
        else:
            out.append((n, e['v']))
    for n, v in (over or {}).items():
        out = [(a, b) for a, b in out if a != n]
        if v is not None:
            out.append((n, v))
    return out


def _option(form, name, want=None, text=None):
    """Value of the option of list `name` whose value == want or whose text contains `text` (case-insensitive); None when the list or option is missing."""
    for e in form['els']:
        if e['t'] == 'select' and e.get('n') and e['n'].endswith(name):
            for v, tx, _ in e['opts']:
                if (want is not None and v == str(want)) or (text and text in tx.strip().lower()):
                    return e['n'], v
            return e['n'], None
    return None, None


def _name(form, suffix):
    return next((e['n'] for e in form['els'] if e.get('n') and e['n'].endswith(suffix)), None)


class _Session:
    def __init__(self):
        self.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        self.op.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) SlipKiosk')]
        self.url = ''
        self.page = ''

    def go(self, url, fields=None):
        """GET (fields=None) or POST; returns (headers, body bytes). Redirects are followed; self.url = final address."""
        data = urllib.parse.urlencode(fields).encode('utf-8') if fields is not None else None
        with self.op.open(urllib.request.Request(url, data=data), timeout=90) as r:
            self.url, body, hd = r.geturl(), r.read(), r.headers
        ctype = hd.get('Content-Type', '').lower()
        disp = hd.get('Content-Disposition', '').lower()
        if 'text/html' in ctype and 'attachment' not in disp:
            enc = re.search(r'charset=([\w-]+)', ctype)
            self.page = body.decode(enc.group(1) if enc else 'utf-8', 'replace')
        return hd, body

    def login_on(self):
        """Fill the user / password of the login form on the current page and submit it."""
        form = next(f for f in _parse_forms(self.page) if any(e['t'] == 'password' for e in f['els']))
        els = form['els']
        pw = next(i for i, e in enumerate(els) if e['t'] == 'password')
        usr = next((e for e in reversed(els[:pw]) if e['t'] in ('text', 'email', 'tel') and e.get('n')), None)
        if usr is None:
            raise ValueError('login form not recognised')
        btn = next((e['n'] for e in els if e['t'] in ('submit', 'image') and e.get('n')), None)
        self.go(urllib.parse.urljoin(self.url, form['action']), _payload(form, {usr['n']: DL['user'], els[pw]['n']: DL['pw']}, btn))


def _wh_dates(today=None):
    """From = the 26th that started the current cycle, To = yesterday (on the 26th itself: today). e.g. 14 Nov -> 26 Oct .. 13 Nov, 26 Nov -> 26 Nov .. 26 Nov."""
    t = (today or datetime.now()).date()
    start = t.replace(day=CYCLE_DAY) if t.day >= CYCLE_DAY else (t.replace(day=1) - timedelta(days=1)).replace(day=CYCLE_DAY)
    return start, max(t - timedelta(days=1), start)


def _debug(page, why):
    try:
        open(DEBUG_PAGE, 'w', encoding='utf-8').write(f'<!-- {why} -->\n' + page)
    except OSError:
        pass


def download_report(kind, ses=None):
    """Download one report ('WH' or 'LB') and save it into the source folder. Raises on any problem (message is shown in the window)."""
    ses = ses or _Session()
    url = PORTAL[kind]
    ses.go(url)
    if _has_login(ses.page):                                   # session missing: log in with the form the site shows (once); else with the portal page (once)
        ses.login_on()
        ses.go(url)
        if _has_login(ses.page) and PORTAL.get('site'):
            ses.go(PORTAL['site'])
            if _has_login(ses.page):
                ses.login_on()
            ses.go(url)
    if _has_login(ses.page):
        _debug(ses.page, 'still on the login page')
        raise LoginFailed('Decibel login failed - check the User ID / password')
    form = next((f for f in _parse_forms(ses.page) if _name(f, 'pb_getreport')), None)
    if not form:
        _debug(ses.page, 'report form not found')
        raise ValueError('report page not recognised (see download_debug.html)')
    action = urllib.parse.urljoin(ses.url, form['action'])
    n, cv = _option(form, 'ddl_Client', text=CLIENT_WORD)           # 1) client: the page reloads itself when it changes
    if n is None or cv is None:
        raise ValueError(f'client "{CLIENT_WORD}" not in the list')
    ses.go(action, _payload(form, {n: cv, '__EVENTTARGET': n, '__EVENTARGUMENT': ''}))
    form = next((f for f in _parse_forms(ses.page) if _name(f, 'pb_getreport')), None)
    if not form:
        _debug(ses.page, 'report form lost after choosing the client')
        raise ValueError('report page not recognised after choosing the client')
    over = {'__EVENTTARGET': '', '__EVENTARGUMENT': '', _name(form, 'cb_ExportToExcel'): 'on', n: cv}
    n, sv = _option(form, 'ddl_ServiceStatus', text=STATUS_WORD)    # 2) service status = In Service
    if n:
        if sv is None:
            raise ValueError(f'service status "{STATUS_WORD}" not in the list')
        over[n] = sv
    if kind == 'WH':                                               # 3) date range (LB has none)
        a, b = _wh_dates()
        for pre, d in (('F', a), ('T', b)):
            for suf, val in (('Day', d.day), ('Month', d.month), ('Year', d.year)):
                n, ov = _option(form, f'ddl_{pre}{suf}', want=val)
                if n is None or ov is None:
                    raise ValueError(f'{pre}{suf} {val} not in the list')
                over[n] = ov
    btn = _name(form, 'pb_getreport')
    hd, body = ses.go(action, _payload(form, over, btn))           # 4) Get Report
    disp, ctype = hd.get('Content-Disposition', ''), hd.get('Content-Type', '').lower()
    is_file = 'attachment' in disp.lower() or ('text/html' not in ctype) or (b'aspnetForm' not in body[:200000] and b'<table' in body.lower())
    if not is_file or not body.strip():
        _debug(ses.page, 'no file came back')
        raise ValueError('no report file came back (see download_debug.html)')
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', disp, re.I)
    ext = os.path.splitext(urllib.parse.unquote(m.group(1)).strip())[1].lower() if m else ''
    if ext not in REPORT_EXT:
        ext = '.xlsx' if body[:2] == b'PK' else '.xls'
    if not SOURCE_DIR:
        raise ValueError('source folder not set')
    os.makedirs(SOURCE_DIR, exist_ok=True)
    dst = os.path.join(SOURCE_DIR, REPORT_NAMES[kind] + ext)
    tmp = dst + '.tmp'
    try:
        with open(tmp, 'wb') as f:
            f.write(body)
        check_file(tmp, kind)                                      # an unreadable download never replaces the good file
        os.replace(tmp, dst)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    for e in REPORT_EXT:                                           # same report saved earlier with another extension: remove it
        old = os.path.join(SOURCE_DIR, REPORT_NAMES[kind] + e)
        if e != ext and os.path.exists(old):
            try:
                os.remove(old)
            except OSError:
                pass
    return dst


def download_all():
    """Download WH then LB (one login). Safe to call from any thread; two runs never overlap."""
    if DL['running']:
        return
    DL['running'] = True
    try:
        if not (DL['user'] and DL['pw']):
            DL['result'] = 'Enter the Decibel User ID and Password first'
            DL_ERR.update(WH='no Decibel login set', LB='no Decibel login set'); return
        who = (DL['user'], DL['pw'])
        if DL.get('bad') == who and not DL.get('manual'):
            return                                                   # this login was refused: never retry it every 5 minutes (could lock the account); it is retried after a change or 'Download now'
        ses, msg = _Session(), []
        for kind in ('WH', 'LB'):
            try:
                p = download_report(kind, ses)
                DL_ERR[kind] = ''; msg.append(f'{kind}: saved {os.path.basename(p)}'); DL['bad'] = None
            except LoginFailed as ex:
                DL['bad'] = who
                DL_ERR.update(WH=str(ex), LB=str(ex)); msg.append(str(ex)); break
            except Exception as ex:
                DL_ERR[kind] = str(ex)[:80] or ex.__class__.__name__
                msg.append(f'{kind}: {DL_ERR[kind]}')
        DL['result'] = '; '.join(msg)
        print('Decibel download:', DL['result'])
    finally:
        DL['last'] = datetime.now().strftime('%d %b %H:%M:%S')
        DL['running'] = False


def download_loop():
    """Every DL['minutes'] minutes (and on 'Download now') run download_all()."""
    last = 0.0
    time.sleep(8)
    while True:
        try:
            if DL['force'] or (DL['on'] and SOURCE_DIR and time.time() - last >= max(1, DL['minutes']) * 60):
                DL['manual'], DL['force'], last = DL['force'], False, time.time()
                download_all()
            DL['next'] = (datetime.fromtimestamp(last + max(1, DL['minutes']) * 60).strftime('%H:%M:%S') if DL['on'] and last else '-')
        except Exception as ex:
            DL['result'] = f'download error: {str(ex)[:80]}'
        time.sleep(1)


# ======================================================================================
#  Salary slips: PDF pay statements (one page per employee) -> read once, kept in a cache, shown + printed as a slip
#  Put the salary PDF files (any name, e.g. Administration.pdf, one per department, 5500+ employees in total) into the reports folder or the linked folder:
#  they are indexed by themselves in the background. If an employee is in several PDFs (old months), the newest pay month wins.
# ======================================================================================
SAL_CACHE = 'salary_cache.json'                       # the window app moves this into its hidden data folder
SALARY = {}                                            # employee id -> pay statement (parsed page)
SAL = {'n': 0, 'files': 0, 'month': '', 'loading': False, 'error': '', 'result': 'No salary PDF found yet', 'last': '-', 'done': 0, 'total': 0}
_sal_files, _sal_bad, _sal_lock, _sal_cache_loaded = {}, {}, threading.Lock(), [False]
_NUMBER = re.compile(r'^-?[\d,]+(?:\.\d+)?$')


def _pymupdf():
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        try:
            import fitz
            return fitz
        except ImportError:
            raise ValueError('PDF support is missing (PyMuPDF): rebuild the exe with the new requirements')


def _squash(s):
    return re.sub(r'\s+', '', str(s)).lower()


def _page_lines(page):
    out = []
    for b in page.get_text('dict')['blocks']:
        for ln in b.get('lines', []):
            t = ''.join(s['text'] for s in ln['spans']).replace('\xa0', ' ').strip()
            if t:
                x0, y0, x1, y1 = ln['bbox']
                out.append((x0, y0, x1, y1, t))
    return out


def _group_rows(items, tol=3.0):
    rows = []
    for it in sorted(items, key=lambda i: (i[1], i[0])):
        if rows and abs(rows[-1][0][1] - it[1]) <= tol:
            rows[-1].append(it)
        else:
            rows.append([it])
    return rows


def _salary_page(L):
    """Text lines [(x0, y0, x1, y1, text)] of one PDF page -> pay statement dict, or None when the page is not a pay statement."""
    info, bottom = {}, 0
    for x0, y0, x1, y1, t in L:                        # 'Label :' with its value right below it (same left edge)
        if y0 > 260 or not re.search(r':\s*$', t):
            continue
        val = next((v for v in sorted(L, key=lambda v: v[1]) if abs(v[0] - x0) < 4 and 0 < v[1] - y0 < 25 and not re.search(r':\s*$', v[4])), None)
        if val:
            info[_squash(re.sub(r'\s*:\s*$', '', t))] = val[4]
            bottom = max(bottom, val[3])
    if 'employeeid' not in info:
        return None

    def anchor(name):
        return next((it for it in L if _squash(it[4]) == name and it[1] > bottom - 1), None)

    month = next((t for x0, y0, x1, y1, t in L if y0 < 90 and re.fullmatch(r'[A-Za-z]+ \d{4}', t)), '')
    company = next((t for x0, y0, x1, y1, t in L if y0 < 90 and 100 < x0 < 400 and 'statement' not in t.lower()), 'Service Tyres')
    left, right = (lambda it: it[2] <= 300), (lambda it: it[0] >= 310)
    te, td = anchor('totalearnings'), anchor('totaldeductions')
    earn, ded = [], []
    if te:
        for row in _group_rows([it for it in L if bottom + 1 < it[1] < te[1] - 2]):
            for side, out in ((left, earn), (right, ded)):
                cells = sorted((it for it in row if side(it)), key=lambda i: i[0])
                lab = ' '.join(it[4] for it in cells if not _NUMBER.match(it[4]))
                amt = next((it[4] for it in reversed(cells) if _NUMBER.match(it[4])), None)
                if lab and amt is not None:
                    out.append((lab, amt))

    def row_amount(a, side):
        cells = [it for it in L if a and abs(it[1] - a[1]) <= 3 and _NUMBER.match(it[4]) and side(it)]
        return max(cells, key=lambda i: i[0])[4] if cells else ''

    def table(first_header):                           # header row (found by its first cell) + the numbers on the row below, matched by column position
        h = anchor(first_header)
        if not h:
            return {}
        below = [it for it in L if h[1] + 6 < it[1] < h[1] + 25 and _NUMBER.match(it[4])]
        out = {}
        for hd in sorted((it for it in L if abs(it[1] - h[1]) <= 3), key=lambda i: i[0]):
            cx = (hd[0] + hd[2]) / 2
            v = min(below, key=lambda b: abs((b[0] + b[2]) / 2 - cx), default=None)
            if v and abs((v[0] + v[2]) / 2 - cx) < 60:
                out[hd[4]] = v[4]
        return out

    loans, lh = [], anchor('loanname')
    if lh:
        hb = max(it[3] for it in L if abs(it[1] - lh[1]) < 12 and _squash(it[4]) in ('loanname', 'opening', 'deduction', 'balanceamount', 'remaining', 'instalment'))
        foot = next((it for it in L if it[4].upper().startswith('THIS IS COMPUTER')), None)
        for row in _group_rows([it for it in L if it[1] >= hb - 1 and (not foot or it[1] < foot[1] - 1)]):
            cells = sorted(row, key=lambda i: i[0])
            nums, name = [it[4] for it in cells if _NUMBER.match(it[4])], ' '.join(it[4] for it in cells if not _NUMBER.match(it[4]))
            if name and len(nums) >= 3:
                loans.append([name] + (nums + ['', '', '', ''])[:4])           # name, opening, deduction, balance, remaining instalments
    g = info.get
    return {'id': eid(g('employeeid', '')), 'name': g('employeename', ''), 'type': g('employementtype', g('employmenttype', '')), 'doj': g('dateofjoining', ''),
            'grade': g('grade', ''), 'dept': g('departmentname', ''), 'pos': g('positionname', ''), 'bank': g('bankname', ''), 'acct': g('bankaccountno', ''),
            'gross': g('grosssalary', ''), 'month': month, 'company': company, 'earn': earn, 'te': row_amount(te, left), 'ded': ded, 'td': row_amount(td, right),
            'net': row_amount(anchor('totalnetsalary') or anchor('netsalary'), lambda it: it[0] > 400), 'tax': table('taxableincome'), 'pf': table('membercontribution'), 'loans': loans}


def parse_pdf(path):
    """Every pay statement page of a PDF -> {employee id: statement}. Raises ValueError when the file holds none."""
    mu, recs = _pymupdf(), {}
    with mu.open(path) as doc:
        SAL['total'], SAL['done'] = len(doc), 0
        for i, page in enumerate(doc):
            r = _salary_page(_page_lines(page))
            if r and r['id']:
                recs[r['id']] = r
            SAL['done'] = i + 1
            if i == 4 and not recs:                    # five pages and no pay statement: not a salary file
                break
    if not recs:
        raise ValueError('no pay statements found in this PDF')
    return recs


def _month_key(m):
    mm = re.match(r'([A-Za-z]{3})[A-Za-z]* (\d{4})', m or '')
    return (int(mm.group(2)), _MON.get(mm.group(1).lower(), 0)) if mm else (0, 0)


def _merge_salary():
    """Best statement per employee over all PDFs: newest pay month, then newest file."""
    global SALARY
    best = {}
    for p, f in _sal_files.items():
        for uid, rec in f['recs'].items():
            k = (_month_key(rec.get('month')), f['sig'][1])
            if uid not in best or k > best[uid][0]:
                best[uid] = (k, rec)
    SALARY = {u: r for u, (k, r) in best.items()}
    top = max((k[0] for k, r in best.values()), default=None)
    SAL.update(n=len(SALARY), files=sum(1 for f in _sal_files.values() if f['recs']),
               month=next((r['month'] for k, r in best.values() if k[0] == top), '') if top else '')


def _save_sal_cache():
    try:
        json.dump({'v': 1, 'files': _sal_files}, open(SAL_CACHE + '.t', 'w', encoding='utf-8'))
        os.replace(SAL_CACHE + '.t', SAL_CACHE)
    except Exception as ex:
        print('Could not save the salary cache:', ex)


def refresh_salary():
    """Index new / changed salary PDFs of the reports folder and the linked folder (in the background; old data keeps working)."""
    if not _sal_lock.acquire(blocking=False):
        return
    try:
        if not _sal_cache_loaded[0]:
            _sal_cache_loaded[0] = True
            try:
                c = json.load(open(SAL_CACHE, encoding='utf-8'))
                if c.get('v') == 1:
                    _sal_files.update(c['files'])
                    _merge_salary()
            except Exception:
                pass
        cur = {}
        for folder in dict.fromkeys(f for f in ('.', SOURCE_DIR) if f):
            try:
                names = os.listdir(folder)
            except OSError:
                continue
            for n in names:
                p = os.path.abspath(os.path.join(folder, n))
                if n.lower().endswith('.pdf') and not n.startswith(('~$', '.')) and os.path.isfile(p):
                    try:
                        st = os.stat(p)
                        cur[p] = [st.st_size, st.st_mtime]
                    except OSError:
                        pass
        changed = False
        for p in list(_sal_files):                      # a PDF that was deleted is forgotten (unless its folder is unreachable, e.g. the share is offline)
            if p not in cur and os.path.isdir(os.path.dirname(p)):
                del _sal_files[p]
                changed = True
        for p, sg in cur.items():
            if _sal_files.get(p, {}).get('sig') == sg or _sal_bad.get(p) == sg or not _stable(p):
                continue
            SAL['loading'], SAL['error'] = True, ''
            t0 = time.time()
            try:
                recs = parse_pdf(p)
                _sal_files[p], changed = {'sig': sg, 'recs': recs}, True
                SAL['result'] = f'{os.path.basename(p)}: {len(recs)} salary slips in {time.time() - t0:.0f} s'
                SAL['last'] = datetime.now().strftime('%d %b %H:%M:%S')
                print('Salary PDF loaded:', SAL['result'])
            except Exception as ex:
                _sal_bad[p] = sg                        # not retried until the file changes (or the program restarts)
                SAL['error'] = f'{os.path.basename(p)}: {str(ex)[:70]}'
                SAL['result'] = SAL['error']
                print('Salary PDF skipped:', SAL['error'])
        if changed:
            _merge_salary()
            _save_sal_cache()
    finally:
        SAL['loading'] = False
        _sal_lock.release()


def _amount(v):
    """'44568' -> '44,568' (numbers that already have commas stay as they are)."""
    s = str(v).strip()
    return f'{int(s):,}' if re.fullmatch(r'-?\d{4,}', s) else s


def format_salary(rec, W=40):
    """Pay statement -> compact plain text slip, W characters wide (32 = 57 mm paper, 40 = 80 mm). Only what is needed, no blank lines."""
    import textwrap
    L = []
    line = lambda: L.append('-' * W)
    centre = lambda t: L.append(str(t)[:W].center(W).rstrip())

    def kv(label, value):
        value = str(value)
        room = W - len(value) - 1
        label = label if len(label) <= room else label[:max(1, room - 1)] + '.'
        L.append(label + ' ' * (W - len(label) - len(value)) + value)

    def field(label, value):
        pre = f'{label:<4}: '
        w = textwrap.wrap(str(value), max(8, W - len(pre))) or ['']
        L.append(pre + w[0])
        L.extend(' ' * len(pre) + x for x in w[1:])

    def item(label):                                    # 'Earned Basic Salary (Days: 26)' -> 'Basic Salary (26d)'
        label = re.sub(r'^Earned\s+', '', label)
        return re.sub(r'\(Hrs:\s*(\d+)\)', r'(\1h)', re.sub(r'\(Days:\s*(\d+)\)', r'(\1d)', label))

    centre('SALARY SLIP'); centre(rec.get('month', ''))
    line()
    for k, lab in SAL_FIELDS:
        if rec.get(k):
            field(lab, rec[k])
    if rec.get('gross'):
        kv('Gross', _amount(rec['gross']))
    line()
    L.append('EARNINGS')
    for lab, amt in rec.get('earn', []):
        kv(item(lab), amt)
    kv('Total Earnings', rec.get('te', ''))
    if rec.get('ded'):                                  # no 'None' line when there are no deductions
        L.append('DEDUCTIONS')
        for lab, amt in rec['ded']:
            kv(item(lab), amt)
    kv('Total Deductions', rec.get('td', '') or '0')
    line()
    net = str(rec.get('net', ''))                       # double size font: half as many characters fit on the line
    half = W // 2
    label = 'NET SALARY' if len('NET SALARY') + 1 + len(net) <= half else 'NET'
    L.append(label + ' ' * max(1, half - len(label) - len(net)) + net)
    line()
    tax = rec.get('tax') or {}
    if tax:
        L.append('TAX')
        for k, v in tax.items():
            if k.lower().startswith('taxable') or str(v).strip('0,.') != '':
                kv(k, v)
        line()
    if rec.get('pf'):
        L.append('PROVIDENT FUND')
        for k, v in rec['pf'].items():
            kv(k, v)
        line()
    for name, opening, deduction, balance, remaining in rec.get('loans', []):
        L.append(f'LOAN: {name}'[:W])
        kv(' Opening', opening); kv(' Deduction', deduction); kv(' Balance', balance); kv(' Instalments left', remaining)
        line()
    centre('Computer generated slip')
    return '\n'.join(L)


def salary_text(uid, cols=None):
    rec = SALARY.get(eid(uid))
    return format_salary(rec, cols or COLS) if rec else None



_started = []


def start_watchers():
    """Background checks: the reports folder and the linked folder, both every 5 s (a new file is live in about 15-25 s)."""
    if _started:
        return
    _started.append(1)

    def loop(fn, secs):
        while True:
            try:
                fn()
            except Exception as ex:
                print('Auto-update error:', ex)
            time.sleep(secs)
    threading.Thread(target=loop, args=(refresh, 5), daemon=True).start()
    threading.Thread(target=loop, args=(sync_reports, 5), daemon=True).start()
    threading.Thread(target=download_loop, daemon=True).start()                          # Decibel auto-download every DL['minutes'] minutes
    threading.Thread(target=loop, args=(update_status, 5), daemon=True).start()           # Status Bar (WH / LB / salary: last update + tick / cross)
    threading.Thread(target=loop, args=(refresh_salary, 5), daemon=True).start()         # salary slip PDFs (read once, then cached)


def slip(uid, ts):
    """Attendance + leave balance slip, compact for 57 mm paper (32 characters): Date, Shift, In, Out, Calculated Overtime, Remarks."""
    import textwrap
    uid = eid(uid)
    e, L = DATA.get(uid), '-' * COLS
    if not e:
        return None
    rows = e['att'][-max(DAYS, min(SPAN[0], MAX_SPAN)):]
    shw = min(5, max([2] + [len(a[5]) for a in rows if len(a) > 5]) + 1) if SHOW_SHIFT else 0     # Shift column width (0 = hidden)
    otw = 6 if SHOW_OT else 0                                                  # Overtime column width (0 = hidden)
    space = COLS - (7 + shw + 12 + otw)                                        # room left for Remarks on the same line
    left, right = f'ID: {uid}', f'{ts:%d-%b %H:%M}'                            # ID on the left, punch time on the right: one line
    idl = [left + ' ' * (COLS - len(left) - len(right)) + right] if len(left) + len(right) < COLS else [left, right]
    o = ['ATTENDANCE SLIP'.center(COLS).rstrip(), L] + textwrap.wrap(f'Name: {e["n"]}', COLS, subsequent_indent='      ') + idl + [L]
    o.append(f'{"Date":<7}' + (f'{"Sh":<{shw}}' if shw else '') + f'{"In":<6}{"Out":<6}' + (f'{"OT":<{otw}}' if otw else '') + ('Remarks' if space >= 7 else 'Rem'))
    for a in rows:
        sc = (a[5] if len(a) > 5 else '')
        ot = (a[7] if len(a) > 7 else '')
        rem = (a[6] if len(a) > 6 else clean_remark(a[3])) or (a[4] if a[1] or a[2] else '')   # leave type, else L/E/H flags
        if not a[1] and not a[2] and rem:       # no punches (Off Day, Absent, leave ...): date + shift + remark only, no '-- --'
            base = f'{a[0]:<7}' + (f'{sc[:shw - 1]:<{shw}}' if shw else '')
            o += [base + rem] if len(base) + len(rem) <= COLS else [base.rstrip(), '  ' + rem[:COLS - 2]]
            continue
        txt = f'{a[0]:<7}' + (f'{(sc or "--")[:shw - 1]:<{shw}}' if shw else '') + f'{a[1] or "--":<6}{a[2] or "--":<6}' + (f'{ot:<{otw}}' if otw else '')
        if len(rem) <= space:
            o.append((txt + rem).rstrip())
        else:                                   # long leave name: continue on the next line
            o += [txt.rstrip(), '  ' + rem[:COLS - 2]]
    if uid in today:
        f, l = today[uid]
        o.append((f'{"Today":<7}' + ' ' * shw + f'{f:<6}{(l if l != f else "--"):<6}').rstrip())
    if rows:
        o += [L, f'Late(L): {sum("L" in a[4].split() for a in rows)}  Early(E): {sum("E" in a[4].split() for a in rows)}']
    w = (COLS - 8) // 3
    o += [L, f'{"LEAVE":<8}{"Allowed":>{w}}{"Used":>{w}}{"Left":>{w}}']
    o += [f'{k:<8}{v[0]:>{w}g}{v[1]:>{w}g}{v[2]:>{w}g}' for k, v in e['lv'].items()] or ['No leave data found']
    return '\n'.join(o)


def printer_problem(h):
    """(problem text, jobs waiting) for an open Windows printer handle; empty text = no problem found."""
    import win32print
    info = win32print.GetPrinter(h, 2)
    flags = ((1, 'paused'), (2, 'in error'), (8, 'jammed'), (0x10, 'out of paper'), (0x40, 'having a paper problem'),
             (0x80, 'offline'), (0x400000, 'cover open'))
    bad = [n for b, n in flags if info['Status'] & b]
    if info['Attributes'] & 0x400:
        bad.append('set to "Use Printer Offline"')
    return ', '.join(bad), info['cJobs']


def gdi_print(text, pname=None):
    """Print as normal text through the Windows printer driver (works on any installed printer). Thermal rolls: 9 pt from the left edge;
    a wide page (A4 / Letter): bigger font, margins, and the slip is kept as one block on the page (a long slip continues on the next page)."""
    import win32con, win32ui
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(pname or PRINTER_NAME)
    try:
        width, height, dpi = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES), dc.GetDeviceCaps(win32con.LOGPIXELSY)
        dpx = dc.GetDeviceCaps(win32con.LOGPIXELSX) or dpi
        wide = width / dpx > 5                                    # wider than 5 inch printable = A4 / Letter, not a thermal roll
        mx, my = (int(dpx * PAGE_MARGIN_IN), int(dpi * PAGE_MARGIN_IN)) if wide else (0, 0)
        cols = max(COLS, max((len(ln.rstrip()) for ln in text.split('\n')), default=0))   # widest line, so nothing is cut off
        pt = PAGE_FONT_PT if wide else 9
        em = int(min((width - 2 * mx) / (cols * 0.6), dpi * pt / 72))                     # chosen size, or smaller so every character fits the width
        fonts = {k: win32ui.CreateFont({'name': 'Courier New', 'height': -h, 'weight': w}) for k, h, w in (('n', em, 400), ('b', em, 700), ('big', em * 2 if not wide else int(em * 1.5), 700))}
        step, y = int(em * 1.15), my
        limit = height - my
        dc.StartDoc('Attendance slip')
        dc.StartPage()
        for line in text.split('\n'):
            k = 'big' if is_big(line) else 'b' if is_bold(line) else 'n'
            h = (step * 2 if not wide else int(step * 1.5)) if k == 'big' else step
            if y + h > limit:
                dc.EndPage(); dc.StartPage(); y = my
            dc.SelectObject(fonts[k])
            dc.SetTextColor(0x0000FF if PRINT_RED and is_red(line) else 0)      # COLORREF is BGR: 0x0000FF = red
            dc.TextOut(mx, y, line.rstrip()); y += h
        dc.EndPage()
        dc.EndDoc()
    finally:
        dc.DeleteDC()


def pair_mode(printer_ip=None, printer_name=None):
    """True when a slip goes to a Windows / USB printer in normal-text mode (office printer, A4): attendance + salary slip then share ONE page."""
    return PRINT_MODE == 'text' and bool(printer_name or (not printer_ip and PRINTER_NAME))


def gdi_print_pair(att, sal, pname=None):
    """ONE page: attendance slip on the left, salary slip on the right (either may be None: the other one is printed alone, on the left).
    The font is as big as fits (max PAGE_FONT_PT); a slip that is too long for the page continues on the next page, both columns together."""
    import win32con, win32ui
    texts = [t for t in (att, sal) if t]
    if not texts:
        return
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(pname or PRINTER_NAME)
    try:
        width, height, dpi = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES), dc.GetDeviceCaps(win32con.LOGPIXELSY)
        dpx = dc.GetDeviceCaps(win32con.LOGPIXELSX) or dpi
        mx, my = (int(dpx * PAGE_MARGIN_IN), int(dpi * PAGE_MARGIN_IN)) if width / dpx > 5 else (0, 0)
        widths = [max(COLS, max(len(ln.rstrip()) for ln in t.split('\n'))) for t in texts]
        gap = 4 if len(texts) > 1 else 0
        em = int(min((width - 2 * mx) / ((sum(widths) + gap) * 0.6), dpi * PAGE_FONT_PT / 72))
        fonts = {k: win32ui.CreateFont({'name': 'Courier New', 'height': -h, 'weight': w}) for k, h, w in (('n', em, 400), ('b', em, 700), ('big', int(em * 1.5), 700))}
        step, limit = int(em * 1.15), height - my
        xs = [mx + int((sum(widths[:i]) + gap * i) * 0.6 * em) for i in range(len(texts))]
        rows = []
        for i, t in enumerate(texts):
            col = []
            for line in t.split('\n'):
                k = 'big' if is_big(line) else 'b' if is_bold(line) else 'n'
                if k == 'big' and len(line.rstrip()) * 0.9 * em > widths[i] * 0.6 * em:
                    k = 'b'                                      # too long for the bigger font inside its column
                col.append((line.rstrip(), k, int(step * 1.5) if k == 'big' else step))
            rows.append(col)
        pos = [0] * len(rows)
        dc.StartDoc('Attendance + salary slip')
        while any(pos[i] < len(rows[i]) for i in range(len(rows))):
            dc.StartPage()
            for i, col in enumerate(rows):
                y = my
                while pos[i] < len(col):
                    line, k, h = col[pos[i]]
                    if y + h > limit and y > my:
                        break
                    dc.SelectObject(fonts[k])
                    dc.SetTextColor(0x0000FF if PRINT_RED and is_red(line) else 0)
                    dc.TextOut(xs[i], y, line); y += h; pos[i] += 1
            dc.EndPage()
        dc.EndDoc()
    finally:
        dc.DeleteDC()


def print_pair(att, sal, printer_ip=None, printer_name=None):
    """Attendance + salary slip on one page of a Windows printer (see pair_mode)."""
    import win32print
    pname = printer_name or PRINTER_NAME
    h = win32print.OpenPrinter(pname)
    try:
        bad, jobs = printer_problem(h)
        if bad:
            raise RuntimeError(f'Printer "{pname}" is {bad}. Fix it in Windows (Settings > Printers & scanners)')
        gdi_print_pair(att, sal, pname)
        print(f'  Windows print queue: {win32print.GetPrinter(h, 2)["cJobs"]} job(s) waiting')
    finally:
        win32print.ClosePrinter(h)


def print_slip(text, printer_ip=None, printer_name=None):
    """Print on the default printer, or on this machine's own printer when printer_ip / printer_name is given."""
    pname, pip = (printer_name, PRINTER_IP) if printer_name else ('', printer_ip) if printer_ip else (PRINTER_NAME, PRINTER_IP)
    def raw_line(ln):
        pre = post = b''
        if PRINT_RED and is_red(ln): pre, post = b'\x1br\x01', b'\x1br\x00'              # ESC r 1 = red ink, ESC r 0 = back to black
        if is_big(ln): pre, post = pre + b'\x1bE\x01\x1d!\x11', b'\x1d!\x00\x1bE\x00' + post   # bold + double width / height (GS ! 0x11)
        elif is_bold(ln): pre, post = pre + b'\x1bE\x01', b'\x1bE\x00' + post             # ESC E 1 = bold
        return pre + ln.encode('ascii', 'replace') + post
    body = b'\n'.join(raw_line(ln) for ln in text.split('\n'))
    init = b'\x1b@' + (b'\x1b3' + bytes([LINE_DOTS]) if TIGHT_LINES else b'')         # init (+ tighter line spacing)
    data = init + body + b'\n' * FEED_LINES + b'\x1d\x56\x00'                        # text, short feed, cut
    if pname:                                    # printer installed in Windows (USB, shared, office ...)
        import win32print
        h = win32print.OpenPrinter(pname)
        try:
            bad, jobs = printer_problem(h)
            if bad:
                raise RuntimeError(f'Printer "{pname}" is {bad}. Fix it in Windows (Settings > Printers & scanners)')
            if PRINT_MODE == 'text':
                gdi_print(text, pname)
            else:
                win32print.StartDocPrinter(h, 1, ('Slip', None, 'RAW'))
                win32print.StartPagePrinter(h)
                win32print.WritePrinter(h, data)
                win32print.EndPagePrinter(h)
                win32print.EndDocPrinter(h)
            print(f'  Windows print queue: {win32print.GetPrinter(h, 2)["cJobs"]} job(s) waiting')
        finally:
            win32print.ClosePrinter(h)
        return
    with socket.create_connection((pip, PRINTER_PORT), timeout=10) as s:
        s.sendall(data)


LATEST = {}        # machine name -> what its LCD shows now: {'seq', 'id', 'att', 'sal', 'status', 'msg', 'ts'}
WEB = {'on': False, 'running': False, 'port': 0, 'key': '', 'clear': 20, 'error': ''}


def _latest(uid, text, state, machine):
    d = LATEST.setdefault(machine, {'seq': 0, 'id': '', 'att': None, 'sal': None, 'status': '', 'msg': '', 'ts': 0})
    nf, snf = f'ID {uid} was not found in the reports', f'No salary slip for ID {uid}'
    if state == 'show':        d.update(att=text, sal=None, status='Printing...', msg='')
    elif state == 'notfound':  d.update(att=None, sal=None, status='', msg=nf)
    elif state == 'sshow':     d.update(att=None, sal=text, status='Printing...', msg='')
    elif state == 'snotfound': d.update(att=None, sal=None, status='', msg=snf)
    elif state == 'pshow':     d.update(att=text[0], sal=text[1], status='Preview only', msg='' if text[0] or text[1] else nf)
    elif state in ('ok', 'sok'):   d['status'] = 'Printed'
    elif state in ('err', 'serr'): d['status'] = 'Print failed'
    else:
        return
    d.update(id=uid, ts=time.time(), seq=d['seq'] + 1)


def _hook(uid, text, state, err, machine=''):
    try:
        _latest(uid, text, state, machine)
    except Exception:
        pass
    try:
        if SLIP_HOOK: SLIP_HOOK(uid, text, state, err, machine)
    except Exception:
        pass


def _marked(text):
    """Slip text -> [[line, 'r'|'b'|'g'|'']]: red / bold / big lines, for the web LCD page."""
    if not text:
        return None
    return [[ln, 'g' if is_big(ln) else 'b' if is_bold(ln) else 'r' if is_red(ln) else ''] for ln in text.split('\n')]


LCD_PAGE = """<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>__NAME__ - LCD</title><style>
html,body{margin:0;height:100%;background:#fff;font-family:'Segoe UI',Arial,sans-serif;color:#1a1a1a;overflow:hidden}
#top{display:flex;justify-content:space-between;align-items:center;padding:10px 22px;border-bottom:4px solid #c8000f}
#top b{font-size:26px}#id{color:#c8000f;font-weight:700;font-size:26px}
#wrap{display:flex;gap:10px;padding:10px;height:calc(100% - 76px);box-sizing:border-box}
.panel{flex:1;border:1px solid #e3e3e3;display:flex;flex-direction:column;min-width:0}
.pt{padding:8px 14px;font-weight:700}.pt span{float:right;color:#555}
.box{flex:1;background:#f2f2f2;display:flex;justify-content:center;align-items:flex-start;overflow:hidden;padding:8px}
.paper{background:#fff;border:1px solid #e3e3e3;padding:8px 12px;height:100%;box-sizing:border-box;overflow:hidden;display:flex;align-items:center;justify-content:center}
.pre{white-space:pre;font-family:'Courier New',monospace;display:inline-block;align-self:flex-start;color:#000}
.pre div{min-height:1em}.r{color:#d00000}.b{font-weight:700}.g{font-weight:700;font-size:2em}
.hint{color:#9a9a9a;font-weight:700;font-size:26px;text-align:center;white-space:pre-wrap}
</style></head><body>
<div id="top"><b>__NAME__</b><span id="id"></span></div>
<div id="wrap">
<div class="panel"><div class="pt">Attendance &amp; Leave balance<span id="sa"></span></div><div class="box"><div class="paper" id="pa"></div></div></div>
<div class="panel"><div class="pt">Salary Slip<span id="ss"></span></div><div class="box"><div class="paper" id="ps"></div></div></div>
</div><script>
var N="__N__",last="",st={};
function fit(p){var pre=p.firstChild;if(!pre||pre.className!=="pre")return;var lo=6,hi=70,best=6;
 while(lo<=hi){var m=(lo+hi)>>1;pre.style.fontSize=m+"px";
  if(pre.offsetHeight<=p.clientHeight-18&&pre.offsetWidth<=p.clientWidth-26){best=m;lo=m+1}else hi=m-1}
 pre.style.fontSize=best+"px"}
function draw(p,lines,msg){p.innerHTML="";
 if(!lines){var h=document.createElement("div");h.className="hint";h.textContent=msg||"Slip Preview";p.appendChild(h);return}
 var pre=document.createElement("div");pre.className="pre";
 lines.forEach(function(l){var d=document.createElement("div");d.textContent=l[0]||" ";d.className=l[1];pre.appendChild(d)});
 p.appendChild(pre);fit(p)}
function tick(){fetch("/api/"+N+location.search,{cache:"no-store"}).then(function(r){return r.json()}).then(function(d){
 var idle=d.clear>0&&d.age>d.clear&&d.status!=="Preview only"||!d.seq,k=idle?"idle":String(d.seq)+d.status;
 if(k!==last){last=k;st=d;
  document.getElementById("id").textContent=idle?"":"Employee ID: "+d.id;
  document.getElementById("sa").textContent=idle?"":(d.att?d.status:"");document.getElementById("ss").textContent=idle?"":(d.sal?d.status:"");
  draw(document.getElementById("pa"),idle?null:d.att,idle?"":d.att?"":d.msg);draw(document.getElementById("ps"),idle?null:d.sal,"")}
 }).catch(function(){})}
window.onresize=function(){fit(document.getElementById("pa"));fit(document.getElementById("ps"))};
draw(document.getElementById("pa"),null,"");draw(document.getElementById("ps"),null,"");tick();setInterval(tick,1000);
</script></body></html>"""


def start_web(port, key, clear=20):
    """LCD over the network: every machine gets its own page  http://<this PC>:<port>/lcd/<n>?k=<key>  that any browser (mini PC, TV, tablet) can show full screen."""
    import hmac, urllib.parse
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    WEB.update(on=True, key=key, clear=clear)
    if WEB['running']:
        return

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def send(self, code, body, ctype):
            b = body.encode('utf-8')
            self.send_response(code)
            self.send_header('Content-Type', ctype + '; charset=utf-8'); self.send_header('Content-Length', str(len(b))); self.send_header('Cache-Control', 'no-store')
            self.end_headers(); self.wfile.write(b)

        def do_GET(self):
            u = urllib.parse.urlparse(self.path)
            k = urllib.parse.parse_qs(u.query).get('k', [''])[0]
            if not WEB['on'] or not hmac.compare_digest(k.encode(), str(WEB['key']).encode()):
                return self.send(403, 'Access key missing or wrong', 'text/plain')
            m = re.fullmatch(r'/(lcd|api)/(\d+)', u.path)
            n = int(m.group(2)) if m else 0
            if not m or not 1 <= n <= len(MACHINES):
                return self.send(404, 'Unknown machine', 'text/plain')
            name = MACHINES[n - 1]['name']
            if m.group(1) == 'lcd':
                return self.send(200, LCD_PAGE.replace('__NAME__', html.escape(name)).replace('__N__', str(n)), 'text/html')
            d = LATEST.get(name) or {}
            self.send(200, json.dumps({'seq': d.get('seq', 0), 'id': d.get('id', ''), 'status': d.get('status', ''), 'msg': d.get('msg', ''),
                                       'age': round(time.time() - d.get('ts', 0)), 'clear': int(WEB['clear']),
                                       'att': _marked(d.get('att')), 'sal': _marked(d.get('sal'))}), 'application/json')

    try:
        srv = ThreadingHTTPServer(('0.0.0.0', int(port)), H)
    except OSError as ex:
        WEB['error'] = f'Web LCD could not start on port {port}: {ex}'
        print(WEB['error'])
        return
    WEB.update(running=True, port=int(port), error='')
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    print(f'Web LCD running on port {port}')


def live(dev=None):
    """Listen to ONE attendance machine (dev = {'name', 'ip', 'port', 'key', 'printer_ip', 'printer_name'}); the window app starts one live() thread per machine."""
    from zk import ZK
    dev = dev or {}
    ip, port, key = dev.get('ip', DEVICE_IP), int(dev.get('port', DEVICE_PORT)), int(dev.get('key', COMM_KEY))
    name, pip = dev.get('name', 'Machine'), dev.get('printer_ip') or None          # pip: this machine's own network printer (None = the default printer)
    pname = dev.get('printer_name') or None                                        # or this machine's own Windows / USB printer; the slip prints ONLY there (no fallback to another printer)
    hook = lambda uid, text, state, err: _hook(uid, text, state, err, name)
    tag = f'[{name}] ' if dev else ''

    def connected(ok):
        STATS['machines'][name] = ok
        STATS['connected'] = any(STATS['machines'].values())
    last = {}
    while not STOP.is_set():
        try:
            conn = ZK(ip, port=port, timeout=10, password=key).connect()
            connected(True)
            print(f'{tag}Device connected - waiting for punches...')
            for att in conn.live_capture():
                if STOP.is_set():
                    break
                if att is None:
                    continue
                uid = eid(ID_MAP.get(str(att.user_id), str(att.user_id)))
                if PREVIEW_ONLY:                                          # preview mode: both slips are only shown on the dashboard / LCD; the Print slip buttons print them
                    if time.time() - last.get((uid, 'p'), 0) < DUPLICATE_SECONDS:
                        continue
                    last[(uid, 'p')] = time.time()
                    refresh()
                    hm = att.timestamp.strftime('%H:%M')
                    today[uid] = [today.get(uid, [hm])[0], hm]
                    print(f'{tag}{att.timestamp} ID {uid}: preview only (nothing is printed)')
                    hook(uid, (slip(uid, att.timestamp), salary_text(uid)), 'pshow', '')
                    continue
                punch = getattr(att, 'punch', None)                      # punch state set on the machine: 0 in, 1 out, 2 break out, 3 break in, 4 OT in, 5 OT out
                salary = punch == SALARY_PUNCH_STATE
                if not salary and ATTENDANCE_PUNCH_STATES is not None and punch not in ATTENDANCE_PUNCH_STATES:
                    print(f'{tag}{att.timestamp} ID {uid}: punch state {punch} prints nothing (ignored)')
                    continue
                if time.time() - last.get((uid, salary), 0) < DUPLICATE_SECONDS:      # same person + same kind of slip only once per 10 minutes (per machine)
                    continue
                last[(uid, salary)] = t0 = time.time()
                if salary:
                    text = salary_text(uid)
                    print(f'{tag}{att.timestamp} ID {uid} punch state {punch}:', 'printing the SALARY slip' if text else 'no salary slip for this ID')
                    hook(uid, text, 'sshow' if text else 'snotfound', '')
                    if text:
                        try:
                            if pair_mode(pip, pname):                            # A4 / office printer: salary + attendance slip on ONE page
                                refresh(); print_pair(slip(uid, att.timestamp), text, pip, pname)
                            else:
                                print_slip(text, pip, pname)
                            hook(uid, text, 'sok', '')
                            STATS.update(slips=STATS['slips'] + 1, printer='OK', last=f'{att.timestamp:%H:%M:%S}  {name}  ID {uid}  salary slip')
                            print(f'{tag}  salary slip sent to printer in {time.time() - t0:.2f} s')
                        except Exception as ex:
                            STATS['printer'] = f'Error: {ex}'
                            print(f'{tag}  Printer error:', ex)
                            hook(uid, text, 'serr', str(ex))
                    continue
                refresh()
                hm = att.timestamp.strftime('%H:%M')
                today[uid] = [today.get(uid, [hm])[0], hm]
                text = slip(uid, att.timestamp)
                print(f'{tag}{att.timestamp} ID {uid} punch state {punch}:', 'printing' if text else 'ID not found in reports')
                hook(uid, text, 'show' if text else 'notfound', '')
                if text:
                    try:
                        if pair_mode(pip, pname):                                # A4 / office printer: attendance + salary slip on ONE page
                            print_pair(text, salary_text(uid), pip, pname)
                        else:
                            print_slip(text, pip, pname)
                        hook(uid, text, 'ok', '')
                        STATS.update(slips=STATS['slips'] + 1, printer='OK',
                                     last=f'{att.timestamp:%H:%M:%S}  {name}  ID {uid}  {DATA[uid]["n"]}')
                        print(f'{tag}  slip sent to printer in {time.time() - t0:.2f} s')
                    except Exception as ex:
                        STATS['printer'] = f'Error: {ex}'
                        print(f'{tag}  Printer error:', ex)
                        hook(uid, text, 'err', str(ex))
            try:
                conn.end_live_capture(); conn.disconnect()
            except Exception:
                pass
            connected(False)
        except Exception as ex:
            connected(False)
            print(f'{tag}Connection problem, retrying in 10 s:', ex)
            time.sleep(10)


if __name__ == '__main__':
    refresh()
    if len(sys.argv) > 2 and sys.argv[1] == '--test':
        print(slip(sys.argv[2], datetime.now()) or 'ID not found in the reports')
    else:
        start_watchers()
        live()
