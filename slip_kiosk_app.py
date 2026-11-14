"""Slip Kiosk - prints an attendance + leave slip for every punch. Needs slip_bridge.py next to it.
Reports: copy the WH / LB files exactly as downloaded (.xls, .xlsx or .csv; the name only has to contain WH or LB) into the 'reports' folder
(or the linked folder). They are read as they are and are live about 15-30 seconds later - no conversion, no restart."""
import base64, hashlib, hmac, io, json, os, queue, re, secrets, shutil, socket, sys, threading, time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
from PIL import Image                      # needed for the system tray icon - add to exe build
import pystray                              # needed for the system tray icon - add to exe build

LOGO_B64 = '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'
BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
DATA = os.path.join(BASE, 'reports')
SOURCE = os.path.join(BASE, 'source')                  # source folder: the browser extension saves WH Report / LB Report here, inside the exe folder
try: os.makedirs(SOURCE, exist_ok=True)
except OSError: pass
# settings + working files live in %APPDATA%\SlipKiosk (hidden from the exe folder); an old config.json next to the exe is moved there once
CFG_DIR = os.path.join(os.environ.get('APPDATA') or BASE, 'SlipKiosk')
CFG, OLD_CFG = os.path.join(CFG_DIR, 'config.json'), os.path.join(BASE, 'config.json')
try:
    os.makedirs(CFG_DIR, exist_ok=True)
    if not os.path.exists(CFG) and os.path.exists(OLD_CFG): shutil.move(OLD_CFG, CFG)
    elif os.path.exists(OLD_CFG): os.remove(OLD_CFG)                      # stale copy: the settings in AppData win
except OSError:
    if not os.path.exists(CFG): CFG = OLD_CFG
STATE = os.path.join(CFG_DIR, 'data')                  # parsed-data cache and old working files: kept here, so the reports folder only shows WH / LB
try: os.makedirs(STATE, exist_ok=True)
except OSError: STATE = CFG_DIR
os.makedirs(DATA, exist_ok=True)
os.chdir(DATA)                                         # the engine looks for the reports in the current folder = the reports folder
q = queue.Queue()


class W:                                   # send all print() output to the log window
    def write(self, s): q.put(s)
    def flush(self): pass


sys.stdout = W()
import slip_bridge as sb                   # engine: finds + reads the newest WH / LB reports (any name, .xls / .xlsx / .csv) and keeps them up to date
sb.CACHE = os.path.join(STATE, 'cache.json')           # parsed data of the loaded reports (instant restart)
sb.SAL_CACHE = os.path.join(STATE, 'salary_cache.json')  # parsed salary slips of all PDFs (the PDFs are read only once)

try:                                                   # one-time tidy-up of what older versions left behind
    for _f in os.listdir(DATA):
        if _f.startswith('cache.json') or re.match(r'^(WH|LB).*\.(tmp|norm)$', _f, re.I):
            try: os.remove(os.path.join(DATA, _f))
            except OSError: pass
    shutil.rmtree(os.path.join(DATA, '.cache'), ignore_errors=True)
    _old = os.path.join(STATE, 'clean')                # converted copies of the older version: reused only if the reports folder has no such report
    for _k in ('WH', 'LB'):
        if os.path.isdir(_old) and not sb.candidates(DATA, _k) and sb.candidates(_old, _k):
            shutil.copy2(max(sb.candidates(_old, _k), key=sb.rank), DATA)
except Exception:
    pass

DEFAULT_MACHINES = [{'on': True, 'name': f'Machine {i}', 'ip': f'192.168.1.{100 + i}', 'port': '4370', 'key': '0', 'pip': '', 'screen': 0} for i in range(1, 6)]
cfg = {'device_ip': '192.168.1.201', 'device_port': '4370', 'comm_key': '0', 'printer_ip': '192.168.1.250',
       'paper': '57', 'days': '30', 'mode': 'network', 'printer_name': '', 'source_dir': '', 'dl_on': True, 'dl_min': '5', 'dc_site': 'https://bpo.mydecibel.com/', 'dc_user': 'ST.404473', 'dc_pass': 'Ahmed!@4190', 'plain': False, 'autorun': False, 'preview_only': False, 'show_keypad': False,
       'machines': DEFAULT_MACHINES, 'lcd_web': False, 'lcd_port': '8080', 'lcd_clear': '20', 'preview_clear': '30', 'lcd_key': ''}
if os.path.exists(CFG):
    cfg.update(json.load(open(CFG)))
cfg['source_dir'] = cfg.get('source_dir') or SOURCE   # default: the 'source' folder next to the exe (can be changed in the Data tab)
if cfg.get('paper_v') != 2: cfg['paper'], cfg['paper_v'] = '57', 2      # one-time switch of older installs from 80 mm to 57 mm paper
cfg.pop('sync_min', None); cfg.pop('auto_salary', None)
if not cfg.get('lcd_key'): cfg['lcd_key'] = secrets.token_hex(4)      # secret in the web LCD links, so only the people given the link can open them
BAT = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup', 'SlipKiosk.bat')
RED, NAVY, BG, FN = '#c8000f', '#111111', '#ffffff', ('Segoe UI', 10)          # Servis Tyres logo: red, black, white
INK, MUTED, LINE = '#1a1a1a', '#6b6b6b', '#e3e3e3'
HEAD, PRIMARY, PRIMARY_D = '#ffffff', '#c8000f', '#a3040c'                       # black header | logo red | darker red (hover)

try:                                        # crisp text on Windows (no blurry scaling)
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

root = tk.Tk(); root.title('Slip Kiosk - Servis Tyres'); root.configure(bg=BG)
root.option_add('*TCombobox*Listbox.font', FN)


def setup_style():
    s = ttk.Style()
    try: s.theme_use('clam')
    except tk.TclError: pass
    s.configure('.', font=FN, background=BG, foreground=INK, focuscolor=BG)
    # tabs: soft pills, the active one is solid navy
    s.configure('TNotebook', background=BG, borderwidth=0, bordercolor=BG, lightcolor=BG, darkcolor=BG, tabmargins=(16, 10, 16, 0))
    s.configure('TNotebook.Tab', font=('Segoe UI', 11, 'bold'), padding=(30, 11), background='#f1f1f1', foreground=MUTED,
                bordercolor=BG, lightcolor=BG, darkcolor=BG, borderwidth=1)
    s.map('TNotebook.Tab', background=[('selected', PRIMARY), ('active', '#e2e2e2')], foreground=[('selected', 'white')],
          bordercolor=[('selected', PRIMARY)], lightcolor=[('selected', PRIMARY)], darkcolor=[('selected', PRIMARY)],
          expand=[('selected', (0, 0, 0, 0))], padding=[('selected', (30, 11))])
    # buttons: default (outlined), Primary (navy), Ghost (for the dark header)
    s.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=(16, 8), background='#f2f2f2', foreground=NAVY, bordercolor='#d0d0d0',
                lightcolor='#f2f2f2', darkcolor='#f2f2f2', relief='flat', focusthickness=0)
    s.map('TButton', background=[('pressed', '#e2e2e2'), ('active', '#e2e2e2')], foreground=[('disabled', '#a8a8a8')])
    s.configure('Primary.TButton', background=PRIMARY, foreground='white', bordercolor=PRIMARY, lightcolor=PRIMARY, darkcolor=PRIMARY)
    s.map('Primary.TButton', background=[('pressed', '#82030a'), ('active', '#a3040c')], foreground=[('disabled', '#f7a8ae')])
    s.configure('Ghost.TButton', background='#f2f2f2', foreground='#111111', bordercolor='#d0d0d0', lightcolor='#f2f2f2', darkcolor='#f2f2f2')
    s.map('Ghost.TButton', background=[('pressed', '#d8d8d8'), ('active', '#e2e2e2')])
    # inputs
    for w in ('TEntry', 'TCombobox'):
        s.configure(w, padding=(8, 6), fieldbackground='white', bordercolor=LINE, lightcolor=LINE, darkcolor=LINE, insertcolor=PRIMARY)
        s.map(w, bordercolor=[('focus', PRIMARY)], lightcolor=[('focus', PRIMARY)], darkcolor=[('focus', PRIMARY)])
    s.configure('TCombobox', background='white', arrowcolor=NAVY, selectbackground='white', selectforeground=INK)
    s.map('TCombobox', fieldbackground=[('readonly', 'white')], selectbackground=[('readonly', 'white')], selectforeground=[('readonly', INK)])
    for w in ('Card.TRadiobutton', 'Card.TCheckbutton'):
        s.configure(w, background='white', foreground=INK, font=FN, indicatorbackground='white', indicatorforeground=NAVY, focuscolor='white', indicatorsize=18, indicatormargin=(0, 0, 8, 0))
        s.map(w, background=[('active', 'white')], indicatorbackground=[('selected', PRIMARY), ('!selected', 'white')],
              indicatorforeground=[('selected', 'white')])
    s.configure('Vertical.TScrollbar', background='#c9c9c9', troughcolor=BG, bordercolor=BG, lightcolor='#c9c9c9', darkcolor='#c9c9c9',
                arrowcolor=NAVY, gripcount=0, arrowsize=16)
    s.map('Vertical.TScrollbar', background=[('active', '#a8a8a8')])


setup_style()


def make_card(parent, title=None, px=18, py=14):
    """White panel with a 1px border. Returns (outer, body): place `outer`, fill `body`."""
    outer = tk.Frame(parent, bg=LINE)
    inner = tk.Frame(outer, bg='white'); inner.pack(fill='both', expand=True, padx=1, pady=1)
    if title:
        tk.Label(inner, text=title, font=('Segoe UI', 11, 'bold'), fg=NAVY, bg='white').pack(anchor='w', padx=px, pady=(py, 0))
        tk.Frame(inner, bg=LINE, height=1).pack(fill='x', padx=px, pady=(8, 0))
    body = tk.Frame(inner, bg='white'); body.pack(fill='both', expand=True, padx=px, pady=py)
    return outer, body


def hover(w, normal, over):
    w.bind('<Enter>', lambda e: str(w['state']) != 'disabled' and w.config(bg=over))
    w.bind('<Leave>', lambda e: w.config(bg=normal))


def center(win):
    win.update_idletasks()
    win.geometry(f'+{(win.winfo_screenwidth() - win.winfo_width()) // 2}+{(win.winfo_screenheight() - win.winfo_height()) // 3}')
    win.attributes('-topmost', True); win.lift(); win.focus_force(); win.grab_set()


def dialog(title):
    win = tk.Toplevel(root); win.title(title); win.configure(bg='white'); win.resizable(False, False)
    tk.Frame(win, bg=RED, height=2).pack(fill='x')
    body = tk.Frame(win, bg='white', padx=40, pady=28); body.pack()
    return win, body

# ---- login (password only). First run: set password + confirm. Forgot it: type "Block HR" as the password to set a new one.
RESET_WORD = 'block hr'
SHARED_PW = None  # PACK_PW   <- pack.py puts the password from password.json (repo) here, so every PC uses the same password


def _active_pw():
    """Password record in use: the shared one from the link; a local one only if no shared password exists
    or it was set on this PC while the current shared one was active (it stops working once the repo password changes)."""
    loc = cfg.get('pw') if isinstance(cfg.get('pw'), dict) and cfg['pw'].get('hash') else None
    if SHARED_PW and isinstance(SHARED_PW, dict) and SHARED_PW.get('hash'):
        return loc if loc and loc.get('base') == SHARED_PW['hash'] else SHARED_PW
    return loc


def _hash(pw, salt):
    return hashlib.pbkdf2_hmac('sha256', pw.encode('utf-8'), bytes.fromhex(salt), 200000).hex()


def _save_pw(pw):
    salt = secrets.token_hex(16)
    cfg['pw'] = {'salt': salt, 'hash': _hash(pw, salt)}        # only a salted hash is stored, never the password itself
    if SHARED_PW and SHARED_PW.get('hash'): cfg['pw']['base'] = SHARED_PW['hash']
    json.dump(cfg, open(CFG, 'w'))
    return {'salt': cfg['pw']['salt'], 'hash': cfg['pw']['hash']}


def show_shared_pw(rec):
    """After a password is set: show the text to save as password.json in the code repo -> all PCs then use this password."""
    win, body = dialog('Use this password on all PCs')
    txt = json.dumps(rec)
    tk.Label(body, text='Password saved on this PC', font=('Segoe UI', 16, 'bold'), fg=NAVY, bg='white').pack(anchor='w')
    tk.Label(body, text='To use the same password on every PC, with no setup on the other PCs:', fg=MUTED, bg='white', font=FN,
             justify='left', wraplength=500).pack(anchor='w', pady=(4, 12))
    for i, t in enumerate(['Click Copy below.', 'In the code repo create or replace  password.json  and paste it.',
                           'Commit and wait for "Publish code to Pages" to turn green.'], 1):
        row = tk.Frame(body, bg='white'); row.pack(anchor='w', fill='x', pady=3)
        tk.Label(row, text=str(i), font=('Segoe UI', 9, 'bold'), fg='white', bg=PRIMARY, width=2).pack(side='left', padx=(0, 12))
        tk.Label(row, text=t, bg='white', fg=INK, font=FN, justify='left', wraplength=460).pack(side='left')
    tk.Label(body, text='Other PCs pick it up the next time the exe starts.', fg=MUTED, bg='white', font=FN).pack(anchor='w', pady=(10, 12))
    e = ttk.Entry(body, width=64, font=('Consolas', 9)); e.insert(0, txt); e.pack(fill='x'); e.state(['readonly'])

    def copy():
        root.clipboard_clear(); root.clipboard_append(txt); root.update()
        cp.config(text='Copied')
    bt = tk.Frame(body, bg='white'); bt.pack(anchor='e', pady=(16, 0))
    cp = ttk.Button(bt, text='Copy', style='Primary.TButton', command=copy); cp.pack(side='left', padx=(0, 8))
    ttk.Button(bt, text='Close', command=win.destroy).pack(side='left')
    center(win)
    root.wait_window(win)


def _check_pw(pw):
    try:
        r = _active_pw()
        return hmac.compare_digest(_hash(pw, r['salt']), r['hash'])
    except Exception:
        return False


def login():
    win, body = dialog('Slip Kiosk - Login')
    ok, mode = [False], ['login' if _active_pw() else 'setup']
    p1, p2, msg = tk.StringVar(), tk.StringVar(), tk.StringVar()
    try: win.logo = tk.PhotoImage(data=LOGO_B64, master=win)       # same Servis Tyres logo as the main window
    except Exception: win.logo = None

    def build():
        for w in body.winfo_children(): w.destroy()
        p1.set(''); p2.set('')
        setup = mode[0] == 'setup'
        if win.logo: tk.Label(body, image=win.logo, bg='white').pack(pady=(0, 10 if setup else 20))
        if setup:                                                   # login screen: logo + password only, no extra heading
            tk.Label(body, text='Create password', font=('Segoe UI', 19, 'bold'), fg=NAVY, bg='white').pack()
            tk.Label(body, text='It protects the settings and the exit button', font=FN, fg=MUTED, bg='white').pack(pady=(2, 16))
        tk.Label(body, text='Password', bg='white', font=FN, fg=MUTED).pack(anchor='w')
        e1 = ttk.Entry(body, textvariable=p1, show='\u2022', width=30, font=('Segoe UI', 12)); e1.pack(fill='x', pady=(3, 10))
        if setup:
            tk.Label(body, text='Confirm password', bg='white', font=FN, fg=MUTED).pack(anchor='w')
            ttk.Entry(body, textvariable=p2, show='\u2022', width=30, font=('Segoe UI', 12)).pack(fill='x', pady=(3, 10))
        tk.Label(body, textvariable=msg, fg=RED, bg='white', font=FN, wraplength=300).pack()
        ttk.Button(body, text='Save password' if setup else 'Login', style='Primary.TButton', command=submit).pack(fill='x', pady=(10, 0))
        e1.focus_set()

    def submit(event=None):
        pw = p1.get()
        if mode[0] == 'setup':
            if len(pw) < 4: msg.set('Password must be at least 4 characters'); return
            if pw.strip().casefold() == RESET_WORD: msg.set('This word cannot be used as a password'); return
            if pw != p2.get(): msg.set('Passwords do not match'); return
            rec = _save_pw(pw); ok[0] = True; win.destroy(); show_shared_pw(rec)
        elif pw.strip().casefold() == RESET_WORD:
            mode[0] = 'setup'; build(); msg.set('Set a new password')
        elif _check_pw(pw):
            ok[0] = True; win.destroy()
        else:
            msg.set('Wrong password'); p1.set('')

    build()
    win.bind('<Return>', submit)
    center(win)
    root.wait_window(win)
    return ok[0]


root.withdraw()
if not login():
    root.destroy(); sys.exit(0)
root.deiconify()
def _work_area():                                    # screen minus the taskbar (the old size let the taskbar cover the bottom of the window)
    try:
        r = (ctypes.c_long * 4)(); ctypes.windll.user32.SystemParametersInfoW(48, 0, r, 0)      # SPI_GETWORKAREA -> left, top, right, bottom
        if r[2] > r[0] and r[3] > r[1]: return r[0], r[1], r[2], r[3]
    except Exception:
        pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight()


_l, _t, _r, _b = _work_area(); _sw, _sh = _r - _l, _b - _t
_w, _h = min(1240, _sw - 24), min(820, _sh - 60)       # 60 px = title bar + window borders
root.geometry(f'{_w}x{_h}+{_l + (_sw - _w) // 2}+{_t + max(0, (_sh - _h - 50) // 3)}'); root.minsize(min(980, _w), min(600, _h)); root.after(300, root.focus_force)
root.protocol('WM_DELETE_WINDOW', lambda: root.iconify())   # X button minimizes to taskbar, never closes
v = {k: tk.StringVar(value=str(cfg[k])) for k in ('device_ip', 'device_port', 'comm_key', 'printer_ip', 'paper', 'days', 'printer_name', 'mode', 'source_dir', 'dl_min', 'dc_site', 'dc_user', 'dc_pass', 'lcd_port', 'lcd_clear')}
dl_on = tk.BooleanVar(value=bool(cfg['dl_on']))
auto, plain, lcdweb = tk.BooleanVar(value=cfg['autorun']), tk.BooleanVar(value=cfg['plain']), tk.BooleanVar(value=bool(cfg['lcd_web']))

# ---- header: navy bar, logo badge, title, status pill
hd = tk.Frame(root, bg=HEAD); hd.pack(fill='x')
badge = tk.Frame(hd, bg='white'); badge.pack(side='left', padx=(18, 14), pady=10)
try:
    logo = tk.PhotoImage(data=LOGO_B64)
    try:                                                              # smooth, smaller logo for the header (falls back to the original)
        from PIL import ImageTk
        logo = ImageTk.PhotoImage(Image.open(io.BytesIO(base64.b64decode(LOGO_B64))).convert('RGBA').resize((57, 54), Image.LANCZOS))
    except Exception:
        pass
    tk.Label(badge, image=logo, bg='white').pack(padx=6, pady=4)
except Exception:
    pass
tb = tk.Frame(hd, bg=HEAD); tb.pack(side='left')
tk.Label(tb, text='Attendance Slip Kiosk', font=('Segoe UI', 18, 'bold'), fg='#111111', bg=HEAD).pack(anchor='w')
tk.Label(tb, text='Servis Tyres  |  Attendance and leave slips', font=('Segoe UI', 9), fg='#6b6b6b', bg=HEAD).pack(anchor='w')
pill = tk.Label(hd, text='\u25cf  STOPPED', font=('Segoe UI', 10, 'bold'), fg='#111111', bg='#eeeeee', padx=16, pady=6); pill.pack(side='right', padx=(10, 18))
ttk.Button(hd, text='Exit', style='Ghost.TButton', command=lambda: request_exit()).pack(side='right')
tread = tk.Canvas(root, height=4, bg=PRIMARY, highlightthickness=0, bd=0); tread.pack(fill='x')      # tyre-tread strip


def _tread(e=None):
    tread.delete('all')
    pass                                                      # plain red line, no zigzag


tread.bind('<Configure>', _tread)

# ---- Status Bar (bottom of every tab): when WH / LB / Salary were last updated + tick / cross
GREEN = '#1a7f37'
sbar_o = tk.Frame(root, bg='#f4f4f4', highlightthickness=1, highlightbackground=LINE); sbar_o.pack(side='bottom', fill='x')
tk.Label(sbar_o, text='Status Bar', bg='#f4f4f4', fg=MUTED, font=('Segoe UI', 9, 'bold')).pack(side='left', padx=(14, 18), pady=6)
stat_w = {}
for _k, _n in sb.STATUS_ROWS:
    _c = tk.Frame(sbar_o, bg='#f4f4f4'); _c.pack(side='left', padx=(0, 26), pady=6)
    tk.Label(_c, text=_n, bg='#f4f4f4', fg=INK, font=('Segoe UI', 9, 'bold')).pack(side='left')
    _t = tk.Label(_c, text='-', bg='#f4f4f4', fg=MUTED, font=('Segoe UI', 9)); _t.pack(side='left', padx=(8, 6))
    _m = tk.Label(_c, text='\u274c', bg='#f4f4f4', fg=PRIMARY, font=('Segoe UI Symbol', 11)); _m.pack(side='left')
    stat_w[_k] = (_t, _m)


def update_status_bar():
    for k, (t, m) in stat_w.items():
        r = sb.STATUS[k]
        t.config(text=r['time'])
        m.config(text='\u2714' if r['ok'] else '\u274c', fg=GREEN if r['ok'] else PRIMARY)


# ---- tabs: Dashboard (optional keypad + slip previews) | Connectivity (machine + printer) | Data (WH / LB reports)
nb = ttk.Notebook(root); nb.pack(fill='both', expand=True, padx=8, pady=(4, 8))
tab_dash, tab_conn, tab_mach, tab_data = (tk.Frame(nb, bg=BG) for _ in range(4))
nb.add(tab_dash, text='Dashboard'); nb.add(tab_conn, text='Connectivity'); nb.add(tab_mach, text='Machines'); nb.add(tab_data, text='Data')

# ---- overview cards
sv = {k: tk.StringVar(value='-') for k in ('mach', 'mach_state', 'prn', 'wh', 'lb', 'emp', 'sal', 'sync', 'watch', 'slips', 'last')}
cardlbl = {}


def card(parent, title, keys, col, wrap=200):
    outer, body = make_card(parent, title, px=16, py=10)
    outer.grid(row=0, column=col, sticky='nsew', padx=6); parent.columnconfigure(col, weight=1, uniform='c')
    for k in keys:
        cardlbl[k] = tk.Label(body, textvariable=sv[k], bg='white', fg=INK, font=('Segoe UI', 10), anchor='w', justify='left', wraplength=wrap)
        cardlbl[k].pack(fill='x', pady=1)


cards_c = tk.Frame(tab_conn, bg=BG); cards_c.pack(fill='x', padx=10, pady=(12, 8))
card(cards_c, 'Machine', ['mach', 'mach_state'], 0); card(cards_c, 'Printer', ['prn'], 1); card(cards_c, 'Activity', ['slips', 'last'], 2)
cards_d = tk.Frame(tab_data, bg=BG); cards_d.pack(fill='x', padx=10, pady=(12, 8))
card(cards_d, 'Reports', ['wh', 'lb', 'emp', 'sal', 'sync'], 0, wrap=900)


# ---- settings
def printers():
    try:
        import win32print
        return sorted(p[2] for p in win32print.EnumPrinters(2 | 4))     # every printer installed in Windows
    except Exception:
        try:
            import subprocess
            out = subprocess.run(['powershell', '-NoProfile', '-Command', '(Get-Printer).Name'], capture_output=True,
                                 text=True, timeout=15, creationflags=0x08000000).stdout
            return sorted(x.strip() for x in out.splitlines() if x.strip())
        except Exception:
            return []


def box(parent, title, col):
    outer, body = make_card(parent, title)
    outer.grid(row=0, column=col, sticky='nsew', padx=6); parent.columnconfigure(col, weight=1, uniform='s'); return body


def field(parent, r, text, var, width=20):
    tk.Label(parent, text=text, bg='white', fg=MUTED, font=FN).grid(row=r, column=0, sticky='w', padx=(0, 16), pady=5)
    ttk.Entry(parent, textvariable=var, width=width).grid(row=r, column=1, sticky='ew', pady=5)


st = tk.Frame(tab_conn, bg=BG); st.pack(fill='x', padx=10)
mf, pf = box(st, 'Slip', 0), box(st, 'Printer connection (tick the one to use)', 1)
for i, (t, k) in enumerate([('Paper (57 / 80 mm roll or A4)', 'paper'), ('Days shown on slip', 'days')]):
    field(mf, i, t, v[k])
mf.columnconfigure(1, weight=1)
pf.columnconfigure(1, weight=1)
ttk.Radiobutton(pf, text='Network printer (IP)', variable=v['mode'], value='network', style='Card.TRadiobutton').grid(row=0, column=0, sticky='w', pady=5)
ttk.Entry(pf, textvariable=v['printer_ip'], width=22).grid(row=0, column=1, sticky='ew', padx=(14, 0), pady=5)
ttk.Radiobutton(pf, text='USB / Windows printer', variable=v['mode'], value='usb', style='Card.TRadiobutton').grid(row=1, column=0, sticky='w', pady=5)
cb = ttk.Combobox(pf, textvariable=v['printer_name'], values=printers(), width=30, state='readonly'); cb.grid(row=1, column=1, sticky='ew', padx=(14, 0), pady=5)
cb.bind('<<ComboboxSelected>>', lambda e: (v['mode'].set('usb'), plain.set(True)))      # a Windows printer prints normal text, attendance + salary slip on one A4 page (untick the box below for a thermal printer)
ttk.Button(pf, text='Refresh', command=lambda: cb.config(values=printers())).grid(row=1, column=2, padx=(8, 0))
def open_queue():
    n = v['printer_name'].get().strip()
    if not n: messagebox.showinfo('Print queue', 'Select a USB / Windows printer first'); return
    import subprocess
    subprocess.Popen(['rundll32', 'printui.dll,PrintUIEntry', '/o', '/n', n])


ttk.Checkbutton(pf, text='Office / laser / inkjet printer (print as normal text)', variable=plain, style='Card.TCheckbutton').grid(row=2, column=0, columnspan=2, sticky='w', pady=(8, 2))
ttk.Button(pf, text='Print queue', command=open_queue).grid(row=2, column=2, padx=(8, 0), pady=(8, 2))
ttk.Checkbutton(pf, text='Start automatically with Windows', variable=auto, style='Card.TCheckbutton').grid(row=3, column=0, columnspan=2, sticky='w', pady=2)


# ---- Data tab: linked folder
af_o, af = make_card(tab_data, 'Reports auto-update')
af_o.pack(fill='x', padx=16, pady=(0, 8))
tk.Label(af, text=f'The app downloads the WH / LB reports from Decibel into the source folder as "{sb.REPORT_NAMES["WH"]}" / "{sb.REPORT_NAMES["LB"]}". Files you copy there yourself (name has WH or LB) are loaded too.', bg='white', fg=MUTED, font=FN, justify='left', wraplength=1100).pack(anchor='w', pady=(0, 4))
for _lbl, _k in (('Work Hour report link', 'WH'), ('Leave Balance report link', 'LB')):
    _r = tk.Frame(af, bg='white'); _r.pack(fill='x', pady=1)
    tk.Label(_r, text=_lbl, bg='white', fg=MUTED, font=FN, width=24, anchor='w').pack(side='left')
    _e = ttk.Entry(_r); _e.insert(0, sb.PORTAL[_k]); _e.config(state='readonly'); _e.pack(side='left', fill='x', expand=True)
tk.Frame(af, bg='white', height=8).pack()
row_f = tk.Frame(af, bg='white'); row_f.pack(fill='x')
tk.Label(row_f, text='Source folder', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(0, 10))
ttk.Entry(row_f, textvariable=v['source_dir']).pack(side='left', fill='x', expand=True)
ttk.Button(row_f, text='Browse', command=lambda: v['source_dir'].set((filedialog.askdirectory(initialdir=v['source_dir'].get() or BASE) or v['source_dir'].get()).replace('/', os.sep))).pack(side='left', padx=(8, 8))


def check_now():
    if safe_apply(): threading.Thread(target=lambda: (sb.sync_reports(), sb.refresh()), daemon=True).start()


ttk.Button(row_f, text='Check now', style='Primary.TButton', command=check_now).pack(side='left', padx=(8, 0))
pc_v = tk.StringVar(value=str(cfg.get('preview_clear', '30')))


def commit_pc(e=None):                                # adjustable clear time of the previewed slip (no password needed)
    s = pc_v.get().strip()
    if s.isdigit(): cfg['preview_clear'] = s; sb.WEB['clear'] = clear_secs()
    else: pc_v.set(str(cfg.get('preview_clear', '30')))
    try: json.dump(cfg, open(CFG, 'w'))
    except OSError: pass


pc_row = tk.Frame(af, bg='white'); pc_row.pack(fill='x', pady=(8, 0)); pc_f = tk.Frame(pc_row, bg='white'); pc_f.pack(side='left')
tk.Label(pc_f, text='Clear previewed slip + keypad after', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(0, 6))
pc_e = ttk.Entry(pc_f, textvariable=pc_v, width=5, justify='center'); pc_e.pack(side='left')
tk.Label(pc_f, text='seconds (0 = never clear)', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(6, 0))
pc_e.bind('<Return>', commit_pc); pc_e.bind('<FocusOut>', commit_pc)


def apply():
    for k in v: cfg[k] = v[k].get().strip()
    cfg['autorun'], cfg['plain'], cfg['lcd_web'] = auto.get(), plain.get(), lcdweb.get()
    cfg['dl_on'] = dl_on.get()
    try:
        if pc_v.get().strip().isdigit(): cfg['preview_clear'] = pc_v.get().strip()
    except NameError: pass
    if not cfg['dl_min'].isdigit() or int(cfg['dl_min']) < 1: raise ValueError('Download repeat time must be a whole number of minutes (1 or more)')
    if not cfg['source_dir']: cfg['source_dir'] = SOURCE
    try: os.makedirs(cfg['source_dir'], exist_ok=True)
    except OSError: pass
    for k in ('device_port', 'comm_key', 'days', 'lcd_port', 'lcd_clear'):
        if not cfg[k].isdigit(): raise ValueError(f'{k.replace("_", " ")} must be a number')
    cfg['paper'] = cfg['paper'].upper()
    if cfg['paper'] not in ('57', '58', '80', 'A4'): raise ValueError('paper must be 57, 80 or A4')
    if cfg['mode'] == 'usb' and not cfg['printer_name']: raise ValueError('Select a USB / Windows printer')
    ms, names = [], set()                                  # the attendance machines table (Machines tab)
    for r in mrows:
        m = {k2: r[k2].get().strip() for k2 in ('name', 'ip', 'port', 'key', 'pip', 'pname')}
        m['on'], m['screen'] = r['on'].get(), screen_num(r['screen'].get())
        m['name'] = m['name'] or f'Machine {len(ms) + 1}'
        m['port'], m['key'] = m['port'] or '4370', m['key'] or '0'
        if m['name'] in names: raise ValueError(f'machine name "{m["name"]}" is used twice')
        names.add(m['name'])
        if m['on'] and not m['ip']: raise ValueError(f'{m["name"]}: machine IP is missing')
        if m['pip'] and m['pname']: raise ValueError(f'{m["name"]}: give either a printer IP or a Windows printer, not both')
        if not (m['port'].isdigit() and m['key'].isdigit()): raise ValueError(f'{m["name"]}: port and comm key must be numbers')
        ms.append(m)
    cfg['machines'] = ms
    sb.MACHINES = machines_from_cfg()
    sb.WEB.update(on=cfg['lcd_web'], key=cfg['lcd_key'], clear=clear_secs())
    sb.PRINTER_IP, sb.PRINTER_NAME = cfg['printer_ip'], (cfg['printer_name'] if cfg['mode'] == 'usb' else '')
    sb.SOURCE_DIR = cfg['source_dir']
    if cfg['dc_site']: sb.PORTAL['site'] = cfg['dc_site']
    sb.DL.update(on=cfg['dl_on'], minutes=int(cfg['dl_min']), user=cfg['dc_user'], pw=cfg['dc_pass'])
    sb.PRINT_MODE = 'text' if (cfg['plain'] or cfg['paper'] == 'A4') else 'raw'       # A4 = office printer: always normal text
    sb.COLS, sb.DAYS = (32 if cfg['paper'] in ('57', '58') else 40), int(cfg['days'])      # the engine widens this to the whole report period by itself
    json.dump(cfg, open(CFG, 'w'))
    sync_lcds(); refresh_links()
    try:
        if cfg['autorun']:
            cmd = f'"{sys.executable}"' if getattr(sys, 'frozen', False) else f'pythonw "{os.path.abspath(__file__)}"'
            open(BAT, 'w').write(f'@echo off\nstart "" {cmd}\n')
        elif os.path.exists(BAT):
            os.remove(BAT)
    except OSError:
        pass


def safe_apply():
    try: apply(); return True
    except ValueError as e: messagebox.showerror('Check settings', str(e)); return False


ths = []


def start():                               # one listener thread per enabled attendance machine; the web LCD server starts too when ticked
    if not safe_apply() or any(t.is_alive() for t in ths): return
    if not sb.MACHINES: messagebox.showinfo('Machines', 'No machine is switched on in the Machines tab'); return
    sb.STOP.clear(); ths.clear()
    for m in sb.MACHINES:
        t = threading.Thread(target=sb.live, args=(m,), daemon=True); t.start(); ths.append(t)
    if cfg['lcd_web']: sb.start_web(int(cfg['lcd_port']), cfg['lcd_key'], clear_secs())
    refresh_links()


def stop(): sb.STOP.set(); sb.STATS['connected'] = False; sb.STATS['machines'].clear()


load_busy, load_btns = [False], []


def load_report(kind):                     # kind: 'WH' (attendance) or 'LB' (leave balance); any name containing WH / LB, .xls (HTML or real), .xlsx or .csv
    f = filedialog.askopenfilename(title=f'Select {kind} report', filetypes=[('Report files', '*.xls *.xlsx *.xlsm *.csv *.txt *.htm *.html'), ('All files', '*.*')])
    if not f: return
    if sb.kind_of(f) != kind:
        messagebox.showerror('Wrong file name', f'The file name must contain  {kind}  (and not the other report name).\nExample:  {kind} 05 Nov 2026.xls\n\nSelected: {os.path.basename(f)}'); return
    cur = sb.report_dates().get(kind)
    if cur and sb.rank(f)[0] < cur:
        messagebox.showerror('Older report', f'The loaded {kind} report is dated {cur:%d %b %Y}.\nThis file ({os.path.basename(f)}) is older and will not be used.'); return
    dst = os.path.join(DATA, os.path.basename(f))
    if load_busy[0]: return
    load_busy[0] = True
    for b in load_btns: b.config(state='disabled')
    load_btns[0 if kind == 'WH' else 1].config(text='Loading, please wait...')

    def done(err=None):                                   # back on the window thread: restore buttons, show any error
        load_busy[0] = False
        for b, t in zip(load_btns, ('Load WH report', 'Load LB report', 'Load salary PDF')): b.config(state='normal', text=t)
        if err: messagebox.showerror('Could not read file', f'Please check the file.\n{err}')

    def work():                                           # heavy part (copy, check, load) runs in the background so the window never hangs
        try:
            if os.path.abspath(f) != os.path.abspath(dst):
                shutil.copyfile(f, dst + '.tmp')
                sb.check_file(dst + '.tmp', kind)
                os.replace(dst + '.tmp', dst)
            else:
                sb.check_file(dst, kind)
            sb._mark_stable(dst)
        except Exception as e:
            try: os.remove(dst + '.tmp')
            except OSError: pass
            root.after(0, lambda e=e: done(e)); return
        root.after(0, done)
        sb.refresh()

    threading.Thread(target=work, daemon=True).start()


def load_salary_pdf():                    # a PDF with salary slips (one page per employee). Normally not needed: just copy the PDFs into the reports folder.
    f = filedialog.askopenfilename(title='Select salary slips PDF', filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')])
    if not f or load_busy[0]: return
    dst = os.path.join(DATA, os.path.basename(f))
    load_busy[0] = True
    for b in load_btns: b.config(state='disabled')
    load_btns[2].config(text='Reading PDF, please wait...')

    def done(err=None):
        load_busy[0] = False
        for b, t in zip(load_btns, ('Load WH report', 'Load LB report', 'Load salary PDF')): b.config(state='normal', text=t)
        if err: messagebox.showerror('Could not read file', f'Please check the file.\n{err}')

    def work():
        try:
            if os.path.abspath(f) != os.path.abspath(dst):
                shutil.copyfile(f, dst + '.tmp'); os.replace(dst + '.tmp', dst)
            sb._mark_stable(dst); sb._sal_bad.pop(os.path.abspath(dst), None)
            sb.refresh_salary()
            err = sb.SAL['error'] if sb.SAL['error'].startswith(os.path.basename(f)) else None
        except Exception as e:
            try: os.remove(dst + '.tmp')
            except OSError: pass
            err = e
        root.after(0, lambda: done(err))
    threading.Thread(target=work, daemon=True).start()


# ---- system tray icon (Open = restore window). Exit stays on the password-protected button below.
tray_q = queue.Queue()
tray_icon = None


def _tray_image():
    return Image.open(io.BytesIO(base64.b64decode(LOGO_B64)))


def _tray_open(icon=None, item=None): tray_q.put('open')
def _tray_exit(icon=None, item=None): tray_q.put('exit')


def start_tray():
    global tray_icon
    menu = pystray.Menu(pystray.MenuItem('Open Slip Kiosk', _tray_open, default=True), pystray.MenuItem('Exit', _tray_exit))
    tray_icon = pystray.Icon('slip_kiosk', _tray_image(), 'Slip Kiosk - Servis Tyres', menu)
    threading.Thread(target=tray_icon.run, daemon=True).start()


# ---- password-protected exit (Exit button below; type the right password to close completely)
def confirm_exit_password(title='Exit Slip Kiosk', prompt='Enter password to exit completely', btn='Exit'):
    win, body = dialog(title)
    pw, msg, result = tk.StringVar(), tk.StringVar(), [False]
    tk.Label(body, text=prompt, font=('Segoe UI', 13, 'bold'), fg=NAVY, bg='white').pack(pady=(0, 12))
    e = ttk.Entry(body, textvariable=pw, show='\u2022', width=28, font=('Segoe UI', 12)); e.pack(fill='x')
    tk.Label(body, textvariable=msg, fg=RED, bg='white', font=FN).pack(pady=(6, 0))

    def submit(event=None):
        if _check_pw(pw.get()): result[0] = True; win.destroy()
        else: msg.set('Wrong password'); pw.set('')

    ttk.Button(body, text=btn, style='Primary.TButton', command=submit).pack(fill='x', pady=(10, 6))
    ttk.Button(body, text='Cancel', command=win.destroy).pack(fill='x')
    win.bind('<Return>', submit)
    center(win)
    e.focus_set()
    root.wait_window(win)
    return result[0]


def request_exit():
    root.deiconify(); root.lift(); root.focus_force()    # show window so the password box isn't hidden
    if confirm_exit_password():
        sb.STOP.set()
        try: tray_icon.stop()
        except Exception: pass
        root.destroy()
        os._exit(0)


# ---- Dashboard tab: touch keypad (left) + slip preview with Print (right)
dash_q = queue.Queue()
entered, pstat, shown = tk.StringVar(), tk.StringVar(), tk.StringVar()
fresh, cur_slip, cur_sal = [False], [None], [None]       # fresh = next digit starts a new ID; cur_slip / cur_sal = attendance / salary slip text on screen
printed = {'att': False, 'sal': False}
tab_dash.columnconfigure(0, weight=0); tab_dash.columnconfigure(1, weight=1, uniform='d'); tab_dash.columnconfigure(2, weight=1, uniform='d'); tab_dash.rowconfigure(1, weight=1)
dbar = tk.Frame(tab_dash, bg=BG); dbar.grid(row=0, column=0, columnspan=3, sticky='ew', padx=14, pady=(8, 0))      # ID of the slips on screen + preview-only switch
dash_id = tk.StringVar()
preview_only = tk.BooleanVar(value=bool(cfg.get('preview_only')))
show_keypad = tk.BooleanVar(value=bool(cfg.get('show_keypad')))


def apply_keypad():                                   # keypad panel on the left of the Dashboard: shown while the box is ticked
    if show_keypad.get(): left_o.grid()
    else: left_o.grid_remove()


def toggle_keypad():                                  # password protected: a wrong / cancelled password puts the box back
    want = show_keypad.get()
    if not confirm_exit_password('Password required', 'Enter password to change this setting', 'OK'):
        show_keypad.set(not want); return
    cfg['show_keypad'] = want; apply_keypad()
    try: json.dump(cfg, open(CFG, 'w'))
    except OSError: pass


ttk.Checkbutton(dbar, text='Show KeyPad', variable=show_keypad, command=toggle_keypad, style='Card.TCheckbutton').pack(side='left', padx=(0, 18))
tk.Label(dbar, textvariable=dash_id, font=('Segoe UI', 13, 'bold'), fg=NAVY, bg=BG).pack(side='left')


def toggle_preview_only():                            # password protected: a wrong / cancelled password puts the box back
    want = preview_only.get()
    if not confirm_exit_password('Password required', 'Enter password to change this setting', 'OK'):
        preview_only.set(not want); return
    cfg['preview_only'] = want; sb.PREVIEW_ONLY = want; sb.WEB['clear'] = clear_secs()
    try: json.dump(cfg, open(CFG, 'w'))
    except OSError: pass


ttk.Checkbutton(dbar, text='Preview only: a punch on the machine shows both slips but prints nothing until Print slip is pressed', variable=preview_only,
                command=toggle_preview_only, style='Card.TCheckbutton').pack(side='right')
left_o, left = make_card(tab_dash, px=12, py=12)
left.config(width=290); left.pack_propagate(False)                     # small keypad panel on the left of the Dashboard
mid_o, mid = make_card(tab_dash, px=14, py=14)
right_o, right = make_card(tab_dash, px=14, py=14)
left_o.grid(row=1, column=0, sticky='ns', padx=(10, 0), pady=8)
mid_o.grid(row=1, column=1, sticky='nsew', padx=(10, 5), pady=8); right_o.grid(row=1, column=2, sticky='nsew', padx=(5, 10), pady=8)
apply_keypad()

top_l = tk.Frame(left, bg='white'); top_l.pack(fill='x', pady=(0, 6))
tk.Label(top_l, text='Employee ID', font=('Segoe UI', 12, 'bold'), fg=NAVY, bg='white').pack(anchor='w')
tk.Label(top_l, text='Type or tap your ID, then Enter', font=('Segoe UI', 9), fg=MUTED, bg='white').pack(anchor='w')
disp_o = tk.Frame(left, bg=LINE); disp_o.pack(fill='x', pady=(0, 10))
disp = tk.Label(disp_o, textvariable=shown, font=('Segoe UI', 22, 'bold'), fg=NAVY, bg='#f6f6f6', anchor='center')
disp.pack(fill='x', padx=1, pady=1, ipady=6)


def _show_id(*_):
    s = entered.get(); shown.set(s or 'ID'); disp.config(fg=NAVY if s else '#bdbdbd'); disp_o.config(bg=PRIMARY if s else LINE)    # grey placeholder when empty; thin red outline while typing


entered.trace_add('write', _show_id); _show_id()
entered.trace_add('write', lambda *_: dash_id.set(('Employee ID: ' + entered.get()) if entered.get() else ''))
pad = tk.Frame(left, bg='white', height=300); pad.pack(fill='x'); pad.grid_propagate(False)      # keypad: a bit smaller, never taller than the panel allows


def fit_pad(e=None):                                  # keypad height = what is left in the panel, capped so the buttons are a little smaller
    free = left.winfo_height() - top_l.winfo_reqheight() - disp_o.winfo_reqheight() - 6 - 10 - 6
    if left.winfo_height() > 100: pad.config(height=max(160, min(380, free)))


left.bind('<Configure>', fit_pad)


btns, _anim = {}, {}


def _mix(c1, c2, t):                                  # colour blend, t = 0 -> c1, t = 1 -> c2
    a, b = [int(c1[i:i + 2], 16) for i in (1, 3, 5)], [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return '#%02x%02x%02x' % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def fade_btn(t, steps=6, ms=20):                      # pressed colour -> normal colour, smoothly
    b, look = btns[t]
    if _anim.get(t): root.after_cancel(_anim[t])
    def step(i=0):
        over = root.winfo_containing(root.winfo_pointerx(), root.winfo_pointery()) is b
        end_bg = look['hv'] if over else look['bg']
        k = (i + 1) / steps
        b.config(bg=_mix(look['pr'], end_bg, k), fg=_mix(look['pf'], look['fg'], k))
        _anim[t] = root.after(ms, step, i + 1) if i + 1 < steps else None
    step()


def flash(t):                                         # used for the physical keyboard: same press effect as a tap
    if t not in btns: return
    b, look = btns[t]
    b.config(bg=look['pr'], fg=look['pf']); root.after(70, lambda: fade_btn(t))


def key(d):
    if fresh[0]: entered.set(''); fresh[0] = False
    if len(entered.get()) < 12: entered.set(entered.get() + d)


PV_PADX, PV_PADY = 14, 10                             # the slip Text widget's own padding (same numbers as in tk.Text(...) below)


def _font_px(px):
    """(line height, width of one character) in pixels for Courier New at `px` - asked straight from Tk, so the 'tkinter.font' module (not packed into the exe) is NOT needed."""
    spec = ('Courier New', -px)
    return int(root.tk.call('font', 'metrics', spec, '-linespace')), int(root.tk.call('font', 'measure', spec, '0'))


class SlipPanel:
    """One slip preview card: title + status, zoom - / + and Print on top, and a sheet of paper whose text is always auto-fitted so the WHOLE slip is visible."""

    def __init__(self, card, title, on_print, idle='Slip Preview'):
        self.shown, self.zoom, self.fit_size, self.fit_ok, self.idle, self.name = None, 0, None, False, idle, title.split('\n')[0]
        self.status = tk.StringVar()
        top = tk.Frame(card, bg='white'); top.pack(fill='x', pady=(0, 8))
        tl = tk.Frame(top, bg='white')                   # packed last: on a narrow window the title gives way, never the buttons
        tk.Label(tl, text=title, font=('Segoe UI', 11, 'bold'), fg=NAVY, bg='white', justify='left').pack(anchor='w')
        self.status_lbl = tk.Label(tl, textvariable=self.status, font=('Segoe UI', 10, 'bold'), bg='white', fg='#555'); self.status_lbl.pack(anchor='w')
        self.btn = tk.Button(top, text='Print slip', command=on_print, font=('Segoe UI', 14, 'bold'), fg='white', bg=PRIMARY, activebackground=PRIMARY_D,
                             activeforeground='white', disabledforeground='#f7a8ae', relief='flat', bd=0, padx=14, pady=6, takefocus=0, cursor='hand2')
        self.btn.pack(side='right'); hover(self.btn, PRIMARY, PRIMARY_D)
        zf = tk.Frame(top, bg='white'); zf.pack(side='right', padx=(6, 10))
        zb = dict(font=('Segoe UI', 16, 'bold'), fg=NAVY, bg='#f2f2f2', activebackground='#e0e0e0', relief='flat', bd=0, width=2, pady=2, takefocus=0, cursor='hand2')
        tk.Button(zf, text='\u2212', command=lambda: self.zoom_by(-2), **zb).pack(side='left')
        self.zlbl = tk.Label(zf, text='100%', font=('Segoe UI', 10, 'bold'), fg=MUTED, bg='white', width=5, cursor='hand2'); self.zlbl.pack(side='left', padx=2)
        tk.Button(zf, text='+', command=lambda: self.zoom_by(2), **zb).pack(side='left')
        self.zlbl.bind('<Button-1>', lambda e: self.zoom_reset())
        tl.pack(side='left')
        self.desk = tk.Frame(card, bg='#f2f2f2'); self.desk.pack(fill='both', expand=True)             # grey desk with the slip as a sheet of paper
        paper = tk.Frame(self.desk, bg=LINE); paper.pack(fill='y', expand=True, pady=8)
        pw_ = tk.Frame(paper, bg='white'); pw_.pack(fill='both', expand=True, padx=1, pady=1)
        self.sc = ttk.Scrollbar(pw_)                              # shown only while the slip is zoomed bigger than the box
        pv = self.pv = tk.Text(pw_, font=('Courier New', 12), width=32, wrap='none', bg='white', fg='#000000', relief='flat', bd=0, padx=PV_PADX, pady=PV_PADY,
                               highlightthickness=0, yscrollcommand=self.sc.set, cursor='arrow', takefocus=0)
        self.xsc = ttk.Scrollbar(pw_, orient='horizontal', command=lambda *a: pv.xview(*a))
        self.sc.config(command=pv.yview); pv.config(xscrollcommand=self.xsc.set)
        pv.pack(side='left', fill='both', expand=True); pv.config(spacing1=0, spacing2=0, spacing3=0)
        pv.bind('<MouseWheel>', lambda e: 'break' if self.fit_ok else None)         # slip is fitted: no scrolling needed, so the wheel is locked
        self.desk.bind('<Configure>', lambda e: root.after_idle(self.fit))          # window resized / different screen -> refit
        pv.tag_configure('hint', font=('Segoe UI', 16, 'bold'), foreground='#9a9a9a', justify='center')
        pv.tag_configure('slip', justify='left', foreground='#000000'); pv.tag_configure('red', foreground='#d00000')
        pv.tag_configure('bold', font=('Courier New', -10, 'bold')); pv.tag_configure('big', font=('Courier New', -20, 'bold'))     # sizes are set again by setfont()

    def setfont(self, px, width):                       # slip font + the bold / double size variants used for the totals and NET SALARY
        self.pv.config(font=('Courier New', -px), width=width)
        self.pv.tag_configure('bold', font=('Courier New', -px, 'bold')); self.pv.tag_configure('big', font=('Courier New', -2 * px, 'bold'))

    def scrollbars(self, on):
        try:
            if on:
                if not self.sc.winfo_ismapped(): self.sc.pack(side='right', fill='y', before=self.pv)
                if not self.xsc.winfo_ismapped(): self.xsc.pack(side='bottom', fill='x', before=self.pv)
            else:
                self.sc.pack_forget(); self.xsc.pack_forget()
        except Exception: pass

    def zoom_by(self, d):
        if not self.shown: return
        self.zoom = max(-8, min(30, self.zoom + d)); self.fit()

    def zoom_reset(self):
        self.zoom = 0
        if self.shown: self.fit()

    def fit(self):                                    # size the slip font so the WHOLE slip is visible: arithmetic on real font metrics + a check on the real widget, no scrolling
        text, pv = self.shown, self.pv
        if not text: return
        try:
            root.update_idletasks()
            lines = text.rstrip('\n').split('\n')
            n, longest = len(lines) + sum(1 for l in lines if sb.is_big(l)), max(len(l) * (2 if sb.is_big(l) else 1) for l in lines)       # the double size line counts as 2 lines / 2x width
            pv.config(height=1); root.update_idletasks()
            H, W = int(pv.winfo_height()), int(self.desk.winfo_width()) - 12
            if H < 80 or W < 80: return                                  # window not laid out yet; the <Configure> event calls us again
            px = 4
            for cand in range(36, 3, -1):                                # biggest font whose height AND width fit (+1 line: Text keeps an empty last line)
                ls, cw = _font_px(cand)
                if ls * (n + 1) + 2 * PV_PADY + 10 <= H and cw * (longest + 1) + 2 * PV_PADX + 10 <= W: px = cand; break
            for _ in range(12):                                          # then CHECK the real widget: last line must be fully inside, otherwise one size smaller
                self.setfont(px, longest + 1); root.update_idletasks()
                bb = pv.bbox(f'{len(lines)}.0')
                try: ok = bb is not None and int(bb[1]) + int(bb[3]) <= int(pv.winfo_height()) - 2
                except Exception: ok = True
                if ok or px <= 4: break
                px -= 1
            fit_px = px
            px = max(4, min(60, fit_px + self.zoom))                     # zoom + / - : offset from the auto-fit size
            zoomed_in = px > fit_px
            if self.zoom != 0:
                ls, cw = _font_px(px)
                chars = max(10, (W - 2 * PV_PADX - 12) // max(1, cw))
                self.setfont(px, min(longest + 1, chars)); root.update_idletasks()
            self.scrollbars(zoomed_in)
            self.zlbl.config(text=f'{int(round(px * 100.0 / max(1, fit_px)))}%')
            self.fit_size, self.fit_ok = px, (not zoomed_in)
            if int(pv.winfo_height()) != H: root.after(50, self.fit)     # layout moved while we measured -> measure once more
        except Exception as e:
            print(f'fit error ({self.name}): {e!r}')
            try: pv.config(font=('Courier New', -max(5, min(14, int(self.desk.winfo_height()) // (len(text.split(chr(10))) + 3)))))       # emergency fallback: rough estimate
            except Exception: pass

    def show(self, text, hint=False):                 # hint=True: big centred grey message instead of a slip
        pv = self.pv
        self.shown = None if hint else text
        self.zoom = 0                                 # every new slip (Enter or machine punch) always opens at 100%
        pv.config(state='normal'); pv.delete('1.0', 'end')
        if hint: self.fit_size = None; self.scrollbars(False); self.zlbl.config(text='100%'); pv.config(font=('Courier New', 12), width=32, height=24)
        pv.insert('end', text, 'hint' if hint else 'slip')
        if not hint:
            for i, ln in enumerate(text.split('\n'), 1):
                if sb.is_red(ln): pv.tag_add('red', f'{i}.0', f'{i}.end')            # Absent / Fullday Ded. lines in red
                if sb.is_big(ln): pv.tag_add('big', f'{i}.0', f'{i}.end')            # NET SALARY: double size, bold
                elif sb.is_bold(ln): pv.tag_add('bold', f'{i}.0', f'{i}.end')        # totals: bold
        pv.config(state='disabled'); pv.yview_moveto(0)
        if not hint: self.fit(); root.after(200, self.fit)      # second pass once the window has finished laying out

    def clear(self):
        self.zoom = 0; self.status.set(''); self.show('\n\n\n\n' + self.idle, True)


def _stamp(*_): shown_at[0] = time.time() if entered.get() else 0       # the clear timer restarts with every key / punch and stops when the ID is empty
entered.trace_add('write', _stamp)
shown_at, cur_mach = [0], ['']                         # when the last machine punch was shown (for the auto-clear) and which machine it came from (its own printer prints it)


def clear():
    shown_at[0] = 0; cur_mach[0] = ''
    entered.set(''); fresh[0] = False; cur_slip[0] = cur_sal[0] = None
    printed.update(att=False, sal=False); att.clear(); sal.clear()


def show_att(uid):                                    # attendance + leave balance slip of the ID
    loading = sb.STATS['loading']
    if loading and not len(sb.DATA): cur_slip[0] = None; att.show('\n\n\nReports are loading,\ntry again in a moment', True); return
    text = sb.slip(uid, datetime.now())               # while a new report loads, the previous data keeps serving slips
    if not text:
        cur_slip[0] = None
        if loading: att.show('\n\n\nReports are loading,\ntry again in a moment', True)
        else: att.show(f'\n\n\nID {uid}\nwas not found in the reports', True)
        return
    cur_slip[0] = text; att.show(text)


def sal_hint(uid):
    S = sb.SAL
    if S['error'] and not sb.SALARY: return f'\n\n\nSalary slips unavailable\n{S["error"]}'
    if S['loading'] and not sb.SALARY: return '\n\n\nSalary slips are loading,\ntry again in a moment'
    if not sb.SALARY: return '\n\n\nNo salary slips loaded.\nCopy the salary PDF files\ninto the reports folder.'
    return f'\n\n\nNo salary slip\nfor ID {uid}'


def show_sal(uid):                                    # salary slip of the ID (from the PDFs)
    text = sb.salary_text(uid)
    cur_sal[0] = text
    if text: sal.show(text)
    else: sal.show(sal_hint(uid), True)


def enter():
    uid = entered.get().strip()
    if not uid: return
    cur_mach[0] = ''                                   # typed on the keypad: default printer
    fresh[0] = True; att.status.set(''); sal.status.set(''); printed.update(att=False, sal=False)
    if not safe_apply(): return
    show_att(uid); show_sal(uid); shown_at[0] = time.time()      # the auto-clear time starts now


def done_print(which):                                # a slip was printed: when every slip on screen is printed, the ID and previews are cleared for the next employee
    p = att if which == 'att' else sal
    printed[which] = True
    if (not cur_slip[0] or printed['att']) and (not cur_sal[0] or printed['sal']): clear(); printed[which] = False
    p.status_lbl.config(fg=INK); p.status.set('Printed'); root.after(3000, lambda: p.status.get() == 'Printed' and p.status.set(''))


def print_now():                                      # attendance + leave balance slip
    t, mach = cur_slip[0], cur_mach[0]
    if not t: messagebox.showinfo('Print', 'Enter an ID and press Enter to see the slip first'); return
    att.btn.config(state='disabled'); att.status.set('Printing...'); att.status_lbl.config(fg='#555')

    ta, ts_ = cur_slip[0], cur_sal[0]

    def work():
        try:
            pr = printer_of(mach)
            if sb.pair_mode(*pr):                          # Windows printer / A4: attendance + salary slip on ONE page
                sb.print_pair(ta, ts_, *pr); dash_q.put(('ok', t))
                if ts_: dash_q.put(('sok', ts_))
            else: sb.print_slip(t, *pr); dash_q.put(('ok', t))
        except Exception as e: dash_q.put(('err', str(e)))
    threading.Thread(target=work, daemon=True).start()


def print_salary():                                   # salary slip
    t, mach = cur_sal[0], cur_mach[0]
    if not t: messagebox.showinfo('Print', 'There is no salary slip on screen'); return
    sal.btn.config(state='disabled'); sal.status.set('Printing...'); sal.status_lbl.config(fg='#555')

    ta = cur_slip[0]

    def work():
        try:
            pr = printer_of(mach)
            if sb.pair_mode(*pr):                          # Windows printer / A4: attendance + salary slip on ONE page
                sb.print_pair(ta, t, *pr); dash_q.put(('sok', t))
                if ta: dash_q.put(('ok', ta))
            else: sb.print_slip(t, *pr); dash_q.put(('sok', t))
        except Exception as e: dash_q.put(('serr', str(e)))
    threading.Thread(target=work, daemon=True).start()


KP = [['7', '8', '9'], ['4', '5', '6'], ['1', '2', '3'], ['Clear', '0', 'Enter']]
for r, row in enumerate(KP):
    pad.rowconfigure(r, weight=1, uniform='k')
    for c, t in enumerate(row):
        pad.columnconfigure(c, weight=1, uniform='k')
        cmd = clear if t == 'Clear' else enter if t == 'Enter' else (lambda d=t: key(d))
        look = {'Clear': dict(fg=INK, bg='#e8e8e8', hv='#dcdcdc', pr='#c9c9c9', pf=INK),
                'Enter': dict(fg='white', bg=PRIMARY, hv=PRIMARY_D, pr='#82030a', pf='white')}.get(t, dict(fg=NAVY, bg='#f4f4f4', hv='#ebebeb', pr='#d2d2d2', pf=NAVY))
        b = tk.Button(pad, text=t, command=cmd, font=('Segoe UI', 16 if len(t) == 1 else 11, 'bold'), relief='flat', bd=0, fg=look['fg'],
                      bg=look['bg'], activebackground=look['pr'], activeforeground=look['pf'], takefocus=0, cursor='hand2')
        b.grid(row=r, column=c, sticky='nsew', padx=3, pady=3)
        hover(b, look['bg'], look['hv'])
        btns[t] = (b, look)
        b.bind('<ButtonRelease-1>', lambda e, t=t: root.after(40, fade_btn, t), add='+')     # tap: red flash that fades out

att = SlipPanel(mid, 'Attendance &\nLeave balance', print_now)
sal = SlipPanel(right, 'Salary Slip', print_salary)
clear()


def on_key(e):                                         # physical keyboard also works while the keypad is shown on the Dashboard
    try:
        if not show_keypad.get() or nb.select() != str(tab_dash) or isinstance(root.focus_get(), (tk.Entry, ttk.Entry, ttk.Combobox)): return
    except Exception:
        return
    if e.char and e.char.isdigit(): flash(e.char); key(e.char)
    elif e.keysym == 'BackSpace': entered.set(entered.get()[:-1]); fresh[0] = False
    elif e.keysym in ('Return', 'KP_Enter'): flash('Enter'); enter()
    elif e.keysym == 'Escape': flash('Clear'); clear()


root.bind('<Key>', on_key)

sb.SLIP_HOOK = lambda uid, text, state, err, machine='': dash_q.put(('punch', (uid, text, state, err, machine)))

# ---- Connectivity and Data tabs open only after typing the login password (Dashboard is always open)
unlocked, keep_unlocked = [False], [0.0]


def tab_click(e):
    try: i = nb.index(f'@{e.x},{e.y}')
    except tk.TclError: return
    if nb.tabs()[i] == str(tab_dash): return                    # Dashboard: normal
    if not unlocked[0]:
        if not confirm_exit_password('Password required', 'Enter password to open this tab', 'Open'): return 'break'
        unlocked[0] = True
    nb.select(i); return 'break'


def tab_changed(e=None):
    if nb.select() == str(tab_dash):
        if time.time() > keep_unlocked[0]: unlocked[0] = False    # clicking the Dashboard tab locks Connectivity / Data again
    elif not unlocked[0]: nb.select(tab_dash)                   # e.g. Ctrl+Tab without the password


nb.bind('<Button-1>', tab_click); nb.bind('<<NotebookTabChanged>>', tab_changed)

# ---- Connectivity tab buttons + live log | Data tab buttons
bar = tk.Frame(tab_conn, bg=BG); bar.pack(fill='x', padx=16, pady=12)
ttk.Button(bar, text='Save + Start', style='Primary.TButton', command=start).pack(side='left')
ttk.Button(bar, text='Stop', command=stop).pack(side='left', padx=8)
log_o, log_b = make_card(tab_conn, 'Live log', py=10)
log_o.pack(fill='both', expand=True, padx=16, pady=(0, 14))
lsc = ttk.Scrollbar(log_b); lsc.pack(side='right', fill='y')
log = tk.Text(log_b, height=8, font=('Consolas', 9), bg='#111111', fg='#f2f2f2', insertbackground=PRIMARY, relief='flat', bd=0, padx=10, pady=8,
              highlightthickness=0, wrap='word', yscrollcommand=lsc.set)
log.pack(side='left', fill='both', expand=True); lsc.config(command=log.yview)
dc_o, dc = make_card(tab_data, 'Decibel login (editable)')
dc_o.pack(fill='x', padx=16, pady=(0, 8))
dc_row = tk.Frame(dc, bg='white'); dc_row.pack(fill='x')
tk.Label(dc_row, text='Decibel', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(0, 8))
ttk.Entry(dc_row, textvariable=v['dc_site'], width=30).pack(side='left', padx=(0, 16))
tk.Label(dc_row, text='User ID', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(0, 8))
ttk.Entry(dc_row, textvariable=v['dc_user'], width=16).pack(side='left', padx=(0, 16))
tk.Label(dc_row, text='Password', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(0, 8))
dc_pw = ttk.Entry(dc_row, textvariable=v['dc_pass'], width=18, show='*'); dc_pw.pack(side='left', padx=(0, 6))
dc_show = tk.BooleanVar(value=False)
ttk.Checkbutton(dc_row, text='Show', variable=dc_show, style='Card.TCheckbutton', command=lambda: dc_pw.config(show='' if dc_show.get() else '*')).pack(side='left', padx=(0, 12))
dc_msg = tk.Label(dc_row, text='', bg='white', fg=MUTED, font=FN)


def save_login():
    if safe_apply():
        dc_msg.config(text='Saved'); dc_msg.after(3000, lambda: dc_msg.config(text=''))


ttk.Button(dc_row, text='Save', style='Primary.TButton', command=save_login).pack(side='left'); dc_msg.pack(side='left', padx=(10, 0))
dl_row = tk.Frame(dc, bg='white'); dl_row.pack(fill='x', pady=(10, 0))
ttk.Checkbutton(dl_row, text='Auto-download reports', variable=dl_on, style='Card.TCheckbutton', command=lambda: safe_apply()).pack(side='left', padx=(0, 16))
tk.Label(dl_row, text='Repeat every', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(0, 8))
ttk.Entry(dl_row, textvariable=v['dl_min'], width=5, justify='center').pack(side='left')
tk.Label(dl_row, text='minutes', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(8, 16))


def download_now():
    if safe_apply():
        sb.DL['force'] = True; sb.DL['result'] = 'Downloading...'


ttk.Button(dl_row, text='Download now', style='Primary.TButton', command=download_now).pack(side='left')
dl_msg = tk.Label(dl_row, text='', bg='white', fg=MUTED, font=FN, anchor='w', justify='left', wraplength=520); dl_msg.pack(side='left', padx=(14, 0))
bd_o, bar_d = make_card(tab_data, 'Manual upload')
bd_o.pack(fill='x', padx=16, pady=(0, 8))
tk.Label(bar_d, text='Normally not needed: just copy the WH / LB files into the reports folder. Or pick a file here (.xls, .xlsx or .csv, name must contain WH or LB; salary slips: any PDF).', bg='white', fg=MUTED, font=FN).pack(anchor='w', pady=(0, 10))
btn_row = tk.Frame(bar_d, bg='white'); btn_row.pack(anchor='w')
for t, c in [('Load WH report', lambda: load_report('WH')), ('Load LB report', lambda: load_report('LB')), ('Load salary PDF', lambda: load_salary_pdf())]:
    load_btns.append(ttk.Button(btn_row, text=t, style='Primary.TButton', command=c)); load_btns[-1].pack(side='left', padx=(0, 8))



# ---- Machines tab: every attendance machine with its own LCD screen (extra monitor on this PC) and/or its own web LCD page
def clear_secs():                                      # seconds a slip stays on the screens before it is cleared (0 = never): previewed slips use the Dashboard box, otherwise the web LCD value
    try: return max(0, int(cfg.get('preview_clear' if sb.PREVIEW_ONLY else 'lcd_clear') or 0))
    except (TypeError, ValueError): return 0


def preview_secs():                                    # seconds a slip stays on the Dashboard before it is cleared (Data tab box, 0 = never)
    try: return max(0, int(cfg.get('preview_clear') or 0))
    except (TypeError, ValueError): return 0


def printer_of(mach):                                  # (network printer IP, Windows printer name) allotted to a machine: its slips print ONLY there
    m = next((m for m in cfg.get('machines', []) if m.get('name') == mach), None) or {}
    return (m.get('pip') or None), (m.get('pname') or None)


def machines_from_cfg():
    out = []
    for m in cfg.get('machines', []):
        if m.get('on'):
            try: out.append({'name': m['name'], 'ip': m['ip'], 'port': int(m.get('port') or 4370), 'key': int(m.get('key') or 0), 'printer_ip': m.get('pip', ''), 'printer_name': m.get('pname', '')})
            except (ValueError, KeyError): pass
    return out


def monitors():                                        # [(left, top, width, height)] of every screen connected to this PC, left to right
    try:
        from ctypes import wintypes
        res = []
        proc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(wintypes.RECT), ctypes.c_ssize_t)

        def cb(h, dc, rc, d):
            r = rc.contents; res.append((r.left, r.top, r.right - r.left, r.bottom - r.top)); return 1
        ctypes.windll.user32.EnumDisplayMonitors(None, None, proc(cb), 0)
        if res: return sorted(res, key=lambda m: (m[0], m[1]))
    except Exception:
        pass
    return [(0, 0, root.winfo_screenwidth(), root.winfo_screenheight())]


def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM); s.connect(('10.255.255.255', 1)); ip = s.getsockname()[0]; s.close(); return ip
    except Exception:
        return '127.0.0.1'


MONS = monitors()
SCREENS = ['No LCD screen'] + [f'Monitor {i} ({w}x{h})' for i, (l, t, w, h) in enumerate(MONS, 1)]


def screen_text(n):
    n = int(n or 0)
    return SCREENS[n] if 0 < n < len(SCREENS) else SCREENS[0]


def screen_num(s):
    m = re.match(r'Monitor (\d+)', s or '')
    return int(m.group(1)) if m else 0


mach_o, mach_b = make_card(tab_mach, 'Attendance machines and their LCD screens')
mach_o.pack(fill='x', padx=16, pady=(12, 8))
mrows = []
for c, (t, _) in enumerate([('Use', 0), ('Name', 0), ('Machine IP', 0), ('Port', 0), ('Comm Key', 0), ('Own printer IP (optional)', 0), ('Own Windows / USB printer (optional)', 0), ('LCD screen on this PC', 0)]):
    tk.Label(mach_b, text=t, bg='white', fg=MUTED, font=FN).grid(row=0, column=c, sticky='w', padx=(0, 8), pady=(0, 4))


def add_mrow(m=None):
    r = len(mrows) + 1
    m = m or {'on': True, 'name': f'Machine {r}', 'ip': f'192.168.1.{100 + r}', 'port': '4370', 'key': '0', 'pip': '', 'screen': 0}
    row = {'on': tk.BooleanVar(value=bool(m.get('on', True))), 'screen': tk.StringVar(value=screen_text(m.get('screen')))}
    for k2 in ('name', 'ip', 'port', 'key', 'pip', 'pname'): row[k2] = tk.StringVar(value=str(m.get(k2, '')))
    row['w'] = [ttk.Checkbutton(mach_b, variable=row['on'], style='Card.TCheckbutton')]
    for c, (k2, wd) in enumerate((('name', 14), ('ip', 16), ('port', 6), ('key', 6), ('pip', 16)), 1): row['w'].append(ttk.Entry(mach_b, textvariable=row[k2], width=wd))
    pcb = ttk.Combobox(mach_b, textvariable=row['pname'], values=[''] + list(cb['values']), width=26, state='readonly')
    pcb.bind('<<ComboboxSelected>>', lambda e, r_=row: r_['pname'].get() and plain.set(True))      # Windows printer = normal text, both slips on one A4 page
    row['w'].append(pcb)
    row['w'].append(ttk.Combobox(mach_b, textvariable=row['screen'], values=SCREENS, width=22, state='readonly'))
    for c, w in enumerate(row['w']): w.grid(row=r, column=c, sticky='w', padx=(0, 8), pady=3)
    mrows.append(row)


def remove_mrow():
    if len(mrows) > 1:
        for w in mrows.pop()['w']: w.destroy()


for _m in cfg['machines']: add_mrow(_m)
mbt = tk.Frame(mach_o, bg='white'); mbt.pack(fill='x', padx=16, pady=(0, 12))
ttk.Button(mbt, text='Save + Start', style='Primary.TButton', command=start).pack(side='left')
ttk.Button(mbt, text='Save', command=lambda: safe_apply()).pack(side='left', padx=8)
ttk.Button(mbt, text='Add machine', command=add_mrow).pack(side='left'); ttk.Button(mbt, text='Remove last', command=remove_mrow).pack(side='left', padx=8)
tk.Label(mach_b, text='LCD screen: a window that covers that whole monitor and shows ONLY this machine\'s punches (Monitor 1 = the leftmost screen). Own printer (IP or Windows / USB printer): the slips of this machine print ONLY on it; leave both empty to use the printer from the Connectivity tab.',
         bg='white', fg=MUTED, font=FN, justify='left', wraplength=900).grid(row=99, column=0, columnspan=8, sticky='w', pady=(8, 0))

web_o, web_b = make_card(tab_mach, 'LCD over the network (any PC, TV or tablet with a browser, no cable to this PC needed)')
web_o.pack(fill='x', padx=16, pady=(0, 8))
wr = tk.Frame(web_b, bg='white'); wr.pack(fill='x')
ttk.Checkbutton(wr, text='Enable web LCD', variable=lcdweb, style='Card.TCheckbutton').pack(side='left')
tk.Label(wr, text='Port', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(20, 6)); ttk.Entry(wr, textvariable=v['lcd_port'], width=7).pack(side='left')
tk.Label(wr, text='Clear slip after (seconds, 0 = never)', bg='white', fg=MUTED, font=FN).pack(side='left', padx=(20, 6)); ttk.Entry(wr, textvariable=v['lcd_clear'], width=5).pack(side='left')
tk.Label(web_b, text='Open the machine\'s link in full screen (F11) on the LCD at that machine. The link holds a secret key; the first time Windows may ask to allow the program through the firewall.',
         bg='white', fg=MUTED, font=FN, justify='left', wraplength=900).pack(anchor='w', pady=(8, 4))
links = tk.Text(web_b, height=6, font=('Consolas', 10), bg='#f6f6f6', fg=INK, relief='flat', bd=0, padx=8, pady=6, highlightthickness=0, wrap='none', state='disabled')
links.pack(fill='x')


def refresh_links():
    ip, on = lan_ip(), [m for m in cfg.get('machines', []) if m.get('on')]
    if cfg.get('lcd_web'): lines = [f'{m["name"]}:  http://{ip}:{cfg.get("lcd_port", "8080")}/lcd/{i}?k={cfg["lcd_key"]}' for i, m in enumerate(on, 1)]
    else: lines = ['Tick "Enable web LCD" and press Save + Start to get the links.']
    if sb.WEB.get('error'): lines.append(sb.WEB['error'])
    links.config(state='normal'); links.delete('1.0', 'end'); links.insert('end', '\n'.join(lines)); links.config(state='disabled')


class LcdWindow:
    """Full-screen preview of ONE machine on its own monitor: the attendance and salary slips of whoever punched on that machine, cleared again after a while."""

    def __init__(self, name, mon):
        self.name, self.mon, self.stamp = name, mon, 0
        l, t, w_, h_ = MONS[mon - 1]
        w = self.win = tk.Toplevel(root); w.configure(bg=BG); w.title(f'LCD - {name}'); w.overrideredirect(True); w.geometry(f'{w_}x{h_}+{l}+{t}')
        top = tk.Frame(w, bg=HEAD); top.pack(fill='x')
        tk.Label(top, text=name, font=('Segoe UI', 20, 'bold'), fg='#111111', bg=HEAD).pack(side='left', padx=20, pady=10)
        self.idv = tk.StringVar()
        tk.Label(top, textvariable=self.idv, font=('Segoe UI', 20, 'bold'), fg=PRIMARY, bg=HEAD).pack(side='right', padx=20)
        tk.Frame(w, bg=PRIMARY, height=4).pack(fill='x')
        body = tk.Frame(w, bg=BG); body.pack(fill='both', expand=True, padx=8, pady=8)
        body.columnconfigure(0, weight=1, uniform='l'); body.columnconfigure(1, weight=1, uniform='l'); body.rowconfigure(0, weight=1)
        lo, lc = make_card(body, px=14, py=14); ro, rc = make_card(body, px=14, py=14)
        lo.grid(row=0, column=0, sticky='nsew', padx=(0, 5)); ro.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        self.att = SlipPanel(lc, 'Attendance &\nLeave balance', lambda: self.print_('att'))
        self.sal = SlipPanel(rc, 'Salary Slip', lambda: self.print_('sal'))
        self.clear()
        w.lift(); w.after(1000, self.tick)

    def close(self):
        try: self.win.destroy()
        except Exception: pass

    def clear(self):
        self.cur = {'att': None, 'sal': None}; self.stamp = 0; self.idv.set('')
        self.att.clear(); self.sal.clear()

    def tick(self):
        try:
            s = clear_secs()
            if self.stamp and s and time.time() - self.stamp > s: self.clear()
        finally:
            self.win.after(1000, self.tick)

    def print_(self, which):
        t, p = self.cur[which], self.att if which == 'att' else self.sal
        if not t: return
        p.btn.config(state='disabled'); p.status.set('Printing...'); p.status_lbl.config(fg='#555')
        pip, pname = printer_of(self.name)

        def work():
            try:
                if sb.pair_mode(pip, pname): sb.print_pair(self.cur['att'], self.cur['sal'], pip, pname)
                else: sb.print_slip(t, pip, pname)
                ok, err = True, ''
            except Exception as e: ok, err = False, str(e)
            root.after(0, lambda: self.printed(p, ok, err))
        threading.Thread(target=work, daemon=True).start()

    def printed(self, p, ok, err):
        for q_ in ((self.att, self.sal) if ok and sb.pair_mode(*printer_of(self.name)) and self.cur['att'] and self.cur['sal'] else (p,)):
            q_.btn.config(state='normal'); q_.status_lbl.config(fg=INK if ok else RED); q_.status.set('Printed' if ok else 'Print failed')

    def handle(self, uid, text, state, err):               # same states as the Dashboard, but only for this machine's punches
        a, s = self.att, self.sal
        self.stamp = time.time(); self.idv.set(f'Employee ID: {uid}')
        if state in ('show', 'notfound'):
            self.cur['sal'] = None; s.clear()
            if state == 'notfound': self.cur['att'] = None; a.status.set(''); a.show(f'\n\n\nID {uid}\nwas not found in the reports', True)
            else: self.cur['att'] = text; a.show(text); a.status_lbl.config(fg='#555'); a.status.set('Printing...')
        elif state == 'ok': a.status_lbl.config(fg=INK); a.status.set('Printed')
        elif state == 'err': a.status_lbl.config(fg=RED); a.status.set('Print failed')
        elif state == 'pshow':
            a_txt, s_txt = text
            self.cur['att'], self.cur['sal'] = a_txt, s_txt
            if a_txt: a.show(a_txt); a.status_lbl.config(fg='#555'); a.status.set('Preview only')
            else: a.status.set(''); a.show(f'\n\n\nID {uid}\nwas not found in the reports', True)
            if s_txt: s.show(s_txt); s.status_lbl.config(fg='#555'); s.status.set('Preview only')
            else: s.status.set(''); s.show(sal_hint(uid), True)
        elif state in ('sshow', 'snotfound'):
            self.cur['att'] = None; a.clear()
            if state == 'snotfound': self.cur['sal'] = None; s.status.set(''); s.show(sal_hint(uid), True)
            else: self.cur['sal'] = text; s.show(text); s.status_lbl.config(fg='#555'); s.status.set('Printing...')
        elif state == 'sok': s.status_lbl.config(fg=INK); s.status.set('Printed')
        elif state == 'serr': s.status_lbl.config(fg=RED); s.status.set('Print failed')


lcds = {}


def sync_lcds():                                       # open / close / move the LCD windows to match the Machines tab (does nothing when nothing changed)
    want = {m['name']: int(m.get('screen') or 0) for m in cfg.get('machines', []) if m.get('on') and 0 < int(m.get('screen') or 0) <= len(MONS)}
    for n in list(lcds):
        if want.get(n) != lcds[n].mon: lcds.pop(n).close()
    for n, mon in want.items():
        if n not in lcds:
            try: lcds[n] = LcdWindow(n, mon)
            except Exception as e: print(f'LCD window for {n} failed: {e!r}')


def poll():
    if shown_at[0] and preview_secs() and time.time() - shown_at[0] > preview_secs(): clear()      # previewed slip: cleared after the time set in the Data tab
    while not dash_q.empty():
        kind, info = dash_q.get()
        if kind in ('ok', 'err', 'punch'): att.btn.config(state='normal')
        if kind in ('sok', 'serr', 'punch'): sal.btn.config(state='normal')
        if kind == 'ok':
            sb.STATS['printer'] = 'OK (manual print)'; done_print('att')
        elif kind == 'sok':
            sb.STATS['printer'] = 'OK (salary slip)'; done_print('sal')
        elif kind == 'punch':                                   # employee punched on the machine: punch state 4 = salary slip, every other state = attendance + leave balance slip (it prints by itself)
            uid, text, state, err, mach = info
            cur_mach[0], shown_at[0] = mach, time.time()
            if mach in lcds: lcds[mach].handle(uid, text, state, err)      # this machine's own LCD screen
            entered.set(uid); fresh[0] = True
            if state in ('show', 'notfound'):                   # attendance punch: only the attendance slip is shown / printed
                cur_sal[0] = None; sal.clear(); printed.update(att=False, sal=False)
                if state == 'notfound': cur_slip[0] = None; att.status.set(''); att.show(f'\n\n\nID {uid}\nwas not found in the reports', True)
                else: cur_slip[0] = text; att.show(text); att.status_lbl.config(fg='#555'); att.status.set('Printing...')
            elif state == 'ok': att.status_lbl.config(fg=INK); att.status.set('Printed')
            elif state == 'err': att.status_lbl.config(fg=RED); att.status.set('Print failed')
            elif state == 'pshow':                              # preview-only mode: both slips are only shown, Print slip prints them
                a_txt, s_txt = text
                printed.update(att=False, sal=False)
                cur_slip[0], cur_sal[0] = a_txt, s_txt
                if a_txt: att.show(a_txt); att.status_lbl.config(fg='#555'); att.status.set('Preview only')
                else: att.status.set(''); att.show(f'\n\n\nID {uid}\nwas not found in the reports', True)
                if s_txt: sal.show(s_txt); sal.status_lbl.config(fg='#555'); sal.status.set('Preview only')
                else: sal.status.set(''); sal.show(sal_hint(uid), True)
            elif state in ('sshow', 'snotfound'):               # salary punch: only the salary slip is shown / printed
                cur_slip[0] = None; att.clear(); printed.update(att=False, sal=False)
                if state == 'snotfound': cur_sal[0] = None; sal.status.set(''); sal.show(sal_hint(uid), True)
                else: cur_sal[0] = text; sal.show(text); sal.status_lbl.config(fg='#555'); sal.status.set('Printing...')
            elif state == 'sok': sal.status_lbl.config(fg=INK); sal.status.set('Printed')
            elif state == 'serr': sal.status_lbl.config(fg=RED); sal.status.set('Print failed')
        elif kind == 'serr':
            sb.STATS['printer'] = f'Error: {info}'; sal.status_lbl.config(fg=RED); sal.status.set('Print failed')
            messagebox.showerror('Printer', f'Could not print the salary slip:\n{info}')
        else:
            sb.STATS['printer'] = f'Error: {info}'; att.status_lbl.config(fg=RED); att.status.set('Print failed')
            messagebox.showerror('Printer', f'Could not print:\n{info}')
    while not q.empty():
        log.insert('end', q.get()); log.see('end')
    while not tray_q.empty():
        action = tray_q.get()
        if action == 'open': root.deiconify(); root.lift(); root.focus_force()
        elif action == 'exit': request_exit()
    run, S, dts = bool(any(t.is_alive() for t in ths) and not sb.STOP.is_set()), sb.STATS, sb.report_dates()
    pill.config(text='\u25cf  RUNNING' if run else '\u25cf  STOPPED', bg=PRIMARY if run else '#eeeeee', fg='white' if run else '#111111')
    ms_ = S.get('machines', {}); n_ok = sum(1 for m in sb.MACHINES if ms_.get(m['name']))
    dot = {True: '\u25cf', False: '\u25cb'}                 # filled = connected
    sv['mach'].set('\n'.join(dot[bool(ms_.get(m['name']))] + ' ' + m['name'] + '  ' + m['ip'] for m in sb.MACHINES) or 'No machine switched on')
    sv['mach_state'].set(f'Connected: {n_ok} of {len(sb.MACHINES)}')
    cardlbl['mach_state'].config(fg=INK if n_ok else PRIMARY, font=('Segoe UI', 10, 'bold'))
    sv['prn'].set((f'Network {cfg["printer_ip"]}' if cfg['mode'] == 'network' else f'USB / Windows\n{cfg["printer_name"] or "(none selected)"}') + f'\nLast print: {S["printer"]}')
    f = lambda k: f'{k} report: ' + (f'{dts[k]:%d %b %Y}' if k in dts else 'not loaded')
    sv['sync'].set(f'Auto-check: every few seconds - {sb.SYNC["result"]}')
    sv['watch'].set(f'Auto-load: put WH / LB files in the reports folder (or the linked folder)\nLast check of the linked folder: {sb.SYNC["last"]}')
    A = sb.SAL
    sv['sal'].set('Salary slips: ' + (f'{A["n"]} employees ({A["month"]}) from {A["files"]} PDF file(s)' if A['n'] else 'none loaded yet - copy the salary PDF files into the reports folder')
                  + (f'   [reading PDF: page {A["done"]} of {A["total"]}]' if A['loading'] else '') + (f'\nSalary PDF: {A["error"]}' if A['error'] else ''))
    sv['wh'].set(f('WH')); sv['lb'].set(f('LB')); sv['emp'].set('Loading new reports...' + (f'  ({len(sb.DATA)} employees in use meanwhile)' if sb.DATA else '') if S['loading'] else f'{len(sb.DATA)} employees loaded' + (f'\nError: {S["error"]}' if S['error'] else ''))
    sv['slips'].set(f'Slips printed: {S["slips"]}'); sv['last'].set(f'Last punch:\n{S["last"]}')
    update_status_bar()
    dl_msg.config(text=(f'{sb.DL["result"]}  (last run {sb.DL["last"]}' + (f', next {sb.DL["next"]})' if cfg['dl_on'] and sb.DL['next'] != '-' else ')')) if sb.DL['last'] != '-' else sb.DL['result'])
    root.after(500, poll)


sb.SOURCE_DIR = cfg['source_dir']
try: os.makedirs(cfg['source_dir'], exist_ok=True)
except OSError: pass
sb.DL.update(on=bool(cfg['dl_on']), minutes=int(cfg['dl_min']) if str(cfg['dl_min']).isdigit() and int(cfg['dl_min']) > 0 else 5, user=cfg['dc_user'], pw=cfg['dc_pass'])
sb.PREVIEW_ONLY = bool(cfg.get('preview_only'))
sb.MACHINES = machines_from_cfg()
sb.start_watchers()                         # checks the reports folder and the linked folder every few seconds; a new file is live in about 15-30 s
start_tray()
root.after(800, sync_lcds)                 # LCD screens open with the program, even before Start is pressed
refresh_links()
poll()
if cfg['autorun']: root.after(1500, start)
root.mainloop()
