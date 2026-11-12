"""Punch ID -> attendance + leave balance slip on a network thermal printer.
Works with ZKTeco network devices (LP400, UFace 900, ...) that allow live punch capture.

Setup : pip install pyzk            (xlrd only if you ever get an old binary .xls)
Folder: put the two reports in the reports folder (or in the linked folder). Rules:
          - the file name must contain  WH  (attendance)  or  LB  (leave balance)  as a separate word
            (WH.xls, WH 05 Nov.xlsx, Updated_LB_report.csv ... anything else in the name is fine)
          - the file can be the HTML-table .xls export, .xlsx, .xlsm or .csv - it is read as it is, no manual conversion
Daily : drop the new WH and LB reports in the folder. They are picked up automatically within about 30 seconds,
        no restart needed. The newest file of each kind wins. If a new file is unreadable (half-copied, wrong format)
        the old data stays in use.
Test  : python slip_bridge.py --test 401708     (prints slip on screen, no hardware needed)
Live  : python slip_bridge.py                   (listens to the device and prints slips)
"""
import csv, glob, html, io, itertools, json, os, re, shutil, socket, sys, threading, time
from datetime import datetime, timedelta

DEVICE_IP, DEVICE_PORT, COMM_KEY = '192.168.1.201', 4370, 0
PRINTER_IP, PRINTER_PORT = '192.168.1.250', 9100   # network thermal printer (raw ESC/POS)
COLS, DAYS = 40, 30                                 # 80 mm paper: 40 | 58 mm paper: 32
DUPLICATE_SECONDS = 600                             # same person ignored for 10 minutes
SOURCE_DIR = ''                                     # linked folder the daily reports are copied from (checked every 5 s)
PRINT_MODE = 'raw'                                  # 'raw' = ESC/POS thermal printer | 'text' = normal text via the Windows driver (any printer)
PRINTER_NAME = ''                                  # USB / Windows printer name; empty = use the network IP above
SALARY_PUNCH_STATE = 4                              # punch state (the machine's function key) that prints the SALARY slip
ATTENDANCE_PUNCH_STATES = None                      # None = every other state (3, 0, 1 ...) prints the attendance + leave balance slip; {3} = only state 3 does (others are ignored)
PREVIEW_ONLY = False                                # True (set by the window app, password protected): a punch only shows the attendance + salary slip, nothing prints
PRINT_RED = True                                    # Absent / Fullday Ded. lines in red: two-colour ESC/POS paper, or colour on a Windows driver (a one-colour printer just prints black). False = never send the colour command
ID_MAP = {}                                         # only if device IDs differ: {'1': '401708'}


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


def shift_code(v):
    """Roster code as short text: None -> '', 3.0 -> '3'."""
    if v is None or (isinstance(v, float) and v != v):
        return ''
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return str(v).strip()


SLIP_HOOK = None   # set by the window app: SLIP_HOOK(uid, text, state, error) with state 'show' | 'ok' | 'err' | 'notfound' (attendance slip), 'sshow' | 'sok' | 'serr' | 'snotfound' (salary slip) or 'pshow' (preview mode: text = (attendance slip, salary slip))
STOP = threading.Event()   # set by the window app to stop live()
STATS = {'error': '', 'loading': False, 'connected': False, 'slips': 0, 'last': '-', 'printer': '-'}   # shown on the app overview
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
    """-> [(date text, emp id, name, in, out, remarks, roster code)]  (columns are found by their header names)."""
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
                idx = [low.index(c) for c in WH_NEED] + [ros]
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
    for ds, emp, n, i, o, r, sh in wh:
        if not str(emp).strip():
            continue
        if ds not in seen:
            seen[ds] = parse_date(ds)
        if seen[ds] is not None:
            rows.append((seen[ds], emp, n, i, o, r, sh))
    rows.sort(key=lambda t: t[0])                         # by date, equal dates keep the file order
    _last_span[0] = (rows[-1][0] - rows[0][0]).days + 1 if rows else 0
    for dt, emp, n, i, o, r, sh in rows:
        uid = eid(emp)
        d.setdefault(uid, {'n': n, 'att': [], 'lv': {}})['att'].append([dt.strftime('%d-%b'), hm(i), hm(o), r, flags(r), shift_code(sh), clean_remark(r)])
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
        key = s + [['v', 10]]                              # bump v when the parsing rules change
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
            return
        for kind in ('WH', 'LB'):
            tmp = None
            try:
                files = candidates(SOURCE_DIR, kind)
                if not files:
                    msg.append(f'{kind}: no file'); continue
                src = max(files, key=rank)
                f = os.path.basename(src)
                dst = os.path.join('.', f)
                cur = report_dates().get(kind)
                if cur and rank(src)[0] < cur:
                    msg.append(f'{kind}: {f} is older than the loaded one'); continue
                if os.path.exists(dst) and os.path.getsize(dst) == os.path.getsize(src) and int(os.path.getmtime(dst)) == int(os.path.getmtime(src)):
                    msg.append(f'{kind}: {f} up to date'); continue
                if not _stable(src):
                    msg.append(f'{kind}: {f} is still being copied, waiting'); continue
                tmp = dst + '.tmp'
                shutil.copy2(src, tmp); check_file(tmp, kind); os.replace(tmp, dst); _mark_stable(dst)
                msg.append(f'{kind}: {f} updated'); changed = True
            except Exception as ex:
                msg.append(f'{kind}: {str(ex)[:60]}')
                if tmp and os.path.exists(tmp):
                    os.remove(tmp)
        SYNC['result'] = '; '.join(msg)
    finally:
        _sync_lock.release()
    if changed:
        prune()
        refresh()


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
    """Pay statement -> plain text slip, W characters wide (40 = 80 mm paper, 32 = 58 mm)."""
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
        pre = f'{label:<6}: '
        w = textwrap.wrap(str(value), max(8, W - len(pre))) or ['']
        L.append(pre + w[0])
        L.extend(' ' * len(pre) + x for x in w[1:])

    def item(label):                                    # 'Earned Basic Salary (Days: 26)' -> 'Basic Salary (26 days)'
        label = re.sub(r'^Earned\s+', '', label)
        return re.sub(r'\(Hrs:\s*(\d+)\)', r'(\1 hrs)', re.sub(r'\(Days:\s*(\d+)\)', r'(\1 days)', label))

    centre('SALARY SLIP'); centre(rec.get('company', '')); centre(rec.get('month', ''))
    line()
    for k, lab in (('name', 'Name'), ('id', 'ID'), ('dept', 'Dept'), ('pos', 'Post'), ('grade', 'Grade'), ('type', 'Type'), ('doj', 'Joined'), ('bank', 'Bank')):
        if rec.get(k):
            field(lab, rec[k])
    if rec.get('gross'):
        field('Gross', _amount(rec['gross']))
    line()
    L.append('EARNINGS')
    for lab, amt in rec.get('earn', []):
        kv(item(lab), amt)
    kv('Total Earnings', rec.get('te', ''))
    line()
    L.append('DEDUCTIONS')
    for lab, amt in rec.get('ded', []):
        kv(item(lab), amt)
    if not rec.get('ded'):
        L.append('None')
    kv('Total Deductions', rec.get('td', ''))
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
    centre('Computer generated slip'); centre('No signature required')
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
    threading.Thread(target=loop, args=(refresh_salary, 5), daemon=True).start()         # salary slip PDFs (read once, then cached)


def slip(uid, ts):
    uid = eid(uid)
    e, L = DATA.get(uid), '-' * COLS
    if not e:
        return None
    rows = e['att'][-max(DAYS, min(SPAN[0], MAX_SPAN)):]
    shw = min(8, max([5] + [len(a[5]) for a in rows if len(a) > 5]) + 1)      # Shift column width (fits the longest roster code)
    space = COLS - (7 + shw + 12)                                              # room left for Remarks on the same line
    o = ['ATTENDANCE SLIP'.center(COLS), L, f'Name : {e["n"]}', f'ID   : {uid}',
         f'Punch: {ts:%d-%b-%Y %H:%M}', L, f'{"Date":<7}{"Shift":<{shw}}{"In":<6}{"Out":<6}Remarks']
    for a in rows:
        sh = (a[5] if len(a) > 5 else '') or '--'
        rem = (a[6] if len(a) > 6 else clean_remark(a[3])) or (a[4] if a[1] or a[2] else '')   # leave type, else L/E/H flags
        line = f'{a[0]:<7}{sh[:shw - 1]:<{shw}}{a[1] or "--":<6}{a[2] or "--":<6}'
        if len(rem) <= space:
            o.append((line + rem).rstrip())
        else:                                   # long leave name: continue on the next line
            o += [line.rstrip(), '  ' + rem[:COLS - 2]]
    if uid in today:
        f, l = today[uid]
        o.append(f'{"Today":<7}{"":<{shw}}{f:<6}{(l if l != f else "--"):<6}')
    if rows:
        o += [L, f'Late: {sum("L" in a[4].split() for a in rows)}   Early going: {sum("E" in a[4].split() for a in rows)}', 'L=Late  E=Early going']
    o += [L, 'LEAVE BALANCE', f'{"":<9}{"Allowed":>8}{"Used":>8}{"Left":>8}']
    o += [f'{k:<9}{v[0]:>8g}{v[1]:>8g}{v[2]:>8g}' for k, v in e['lv'].items()] or ['No leave data found']
    return '\n'.join(o + [L, 'Thank you'.center(COLS)])


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


def gdi_print(text):
    """Print as normal text through the Windows printer driver (works on any installed printer)."""
    import win32con, win32ui
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(PRINTER_NAME)
    try:
        width, height, dpi = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES), dc.GetDeviceCaps(win32con.LOGPIXELSY)
        em = int(min(width / (COLS * 0.6), dpi * 9 / 72))        # 9 pt, or smaller so COLS characters fit the paper width
        fonts = {k: win32ui.CreateFont({'name': 'Courier New', 'height': -h, 'weight': w}) for k, h, w in (('n', em, 400), ('b', em, 700), ('big', em * 2, 700))}
        step, y = int(em * 1.25), 0
        dc.StartDoc('Attendance slip')
        dc.StartPage()
        for line in text.split('\n'):
            k = 'big' if is_big(line) else 'b' if is_bold(line) else 'n'
            h = step * 2 if k == 'big' else step
            if y + h > height:
                dc.EndPage(); dc.StartPage(); y = 0
            dc.SelectObject(fonts[k])
            dc.SetTextColor(0x0000FF if PRINT_RED and is_red(line) else 0)      # COLORREF is BGR: 0x0000FF = red
            dc.TextOut(0, y, line.rstrip()); y += h
        dc.EndPage()
        dc.EndDoc()
    finally:
        dc.DeleteDC()


def print_slip(text):
    def raw_line(ln):
        pre = post = b''
        if PRINT_RED and is_red(ln): pre, post = b'\x1br\x01', b'\x1br\x00'              # ESC r 1 = red ink, ESC r 0 = back to black
        if is_big(ln): pre, post = pre + b'\x1bE\x01\x1d!\x11', b'\x1d!\x00\x1bE\x00' + post   # bold + double width / height (GS ! 0x11)
        elif is_bold(ln): pre, post = pre + b'\x1bE\x01', b'\x1bE\x00' + post             # ESC E 1 = bold
        return pre + ln.encode('ascii', 'replace') + post
    body = b'\n'.join(raw_line(ln) for ln in text.split('\n'))
    data = b'\x1b@' + body + b'\n\n\n\n\x1d\x56\x00'   # init, text, feed, cut
    if PRINTER_NAME:                                    # printer installed in Windows (USB, shared, office ...)
        import win32print
        h = win32print.OpenPrinter(PRINTER_NAME)
        try:
            bad, jobs = printer_problem(h)
            if bad:
                raise RuntimeError(f'Printer "{PRINTER_NAME}" is {bad}. Fix it in Windows (Settings > Printers & scanners)')
            if PRINT_MODE == 'text':
                gdi_print(text)
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
    with socket.create_connection((PRINTER_IP, PRINTER_PORT), timeout=10) as s:
        s.sendall(data)


def _hook(uid, text, state, err):
    try:
        if SLIP_HOOK: SLIP_HOOK(uid, text, state, err)
    except Exception:
        pass


def live():
    from zk import ZK
    last = {}
    while not STOP.is_set():
        try:
            conn = ZK(DEVICE_IP, port=DEVICE_PORT, timeout=10, password=COMM_KEY).connect()
            STATS['connected'] = True
            print('Device connected - waiting for punches...')
            for att in conn.live_capture():
                if STOP.is_set():
                    break
                if att is None:
                    continue
                uid = eid(ID_MAP.get(str(att.user_id), str(att.user_id)))
                if PREVIEW_ONLY:                                          # preview mode: both slips are only shown on the dashboard; the Print slip buttons print them
                    if time.time() - last.get((uid, 'p'), 0) < DUPLICATE_SECONDS:
                        continue
                    last[(uid, 'p')] = time.time()
                    refresh()
                    hm = att.timestamp.strftime('%H:%M')
                    today[uid] = [today.get(uid, [hm])[0], hm]
                    print(f'{att.timestamp} ID {uid}: preview only (nothing is printed)')
                    _hook(uid, (slip(uid, att.timestamp), salary_text(uid)), 'pshow', '')
                    continue
                punch = getattr(att, 'punch', None)                      # punch state set on the machine: 0 in, 1 out, 2 break out, 3 break in, 4 OT in, 5 OT out
                salary = punch == SALARY_PUNCH_STATE
                if not salary and ATTENDANCE_PUNCH_STATES is not None and punch not in ATTENDANCE_PUNCH_STATES:
                    print(f'{att.timestamp} ID {uid}: punch state {punch} prints nothing (ignored)')
                    continue
                if time.time() - last.get((uid, salary), 0) < DUPLICATE_SECONDS:      # same person + same kind of slip only once per 10 minutes
                    continue
                last[(uid, salary)] = t0 = time.time()
                if salary:
                    text = salary_text(uid)
                    print(f'{att.timestamp} ID {uid} punch state {punch}:', 'printing the SALARY slip' if text else 'no salary slip for this ID')
                    _hook(uid, text, 'sshow' if text else 'snotfound', '')
                    if text:
                        try:
                            print_slip(text)
                            _hook(uid, text, 'sok', '')
                            STATS.update(slips=STATS['slips'] + 1, printer='OK', last=f'{att.timestamp:%H:%M:%S}  ID {uid}  salary slip')
                            print(f'  salary slip sent to printer in {time.time() - t0:.2f} s')
                        except Exception as ex:
                            STATS['printer'] = f'Error: {ex}'
                            print('  Printer error:', ex)
                            _hook(uid, text, 'serr', str(ex))
                    continue
                refresh()
                hm = att.timestamp.strftime('%H:%M')
                today[uid] = [today.get(uid, [hm])[0], hm]
                text = slip(uid, att.timestamp)
                print(f'{att.timestamp} ID {uid} punch state {punch}:', 'printing' if text else 'ID not found in reports')
                _hook(uid, text, 'show' if text else 'notfound', '')
                if text:
                    try:
                        print_slip(text)
                        _hook(uid, text, 'ok', '')
                        STATS.update(slips=STATS['slips'] + 1, printer='OK',
                                     last=f'{att.timestamp:%H:%M:%S}  ID {uid}  {DATA[uid]["n"]}')
                        print(f'  slip sent to printer in {time.time() - t0:.2f} s')
                    except Exception as ex:
                        STATS['printer'] = f'Error: {ex}'
                        print('  Printer error:', ex)
                        _hook(uid, text, 'err', str(ex))
            try:
                conn.end_live_capture(); conn.disconnect()
            except Exception:
                pass
        except Exception as ex:
            STATS['connected'] = False
            print('Connection problem, retrying in 10 s:', ex)
            time.sleep(10)


if __name__ == '__main__':
    refresh()
    if len(sys.argv) > 2 and sys.argv[1] == '--test':
        print(slip(sys.argv[2], datetime.now()) or 'ID not found in the reports')
    else:
        start_watchers()
        live()
