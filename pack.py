"""Packs the app code into site/data.txt (compressed + base64).
If a code file is missing (deleted), data.txt is published with a DISABLED marker, so every exe stops working.
password.json (optional: {"salt":..,"hash":..}) is injected into the app = one shared login password for all PCs.
A syntax error in a code file stops the deploy (the previous version stays live)."""
import base64, json, os, zlib
FILES = ['slip_bridge.py', 'slip_kiosk_app.py']
code, missing = {}, []
for f in FILES:
    if not os.path.isfile(f):
        missing.append(f)
        continue
    src = open(f, encoding='utf-8').read()
    if f == 'slip_kiosk_app.py' and os.path.isfile('password.json'):      # shared password for every PC
        pw = json.load(open('password.json', encoding='utf-8'))
        assert set(pw) >= {'salt', 'hash'}, 'password.json must contain salt and hash'
        marker = 'SHARED_PW = None  # PACK_PW'
        assert marker in src, 'PACK_PW marker missing in slip_kiosk_app.py'
        src = src.replace(marker, 'SHARED_PW = ' + repr({'salt': pw['salt'], 'hash': pw['hash']}) + '  # from password.json', 1)
        print('shared password injected')
    compile(src, f, 'exec')
    code[f] = src
if missing:
    print('MISSING FILES:', missing, '-> publishing DISABLED marker')
    code = {'__disabled__': 'missing ' + ', '.join(missing)}
blob = zlib.compress(json.dumps(code).encode('utf-8'), 9)
os.makedirs('site', exist_ok=True)
open('site/data.txt', 'w').write(base64.b64encode(blob).decode('ascii'))
open('site/index.html', 'w').write('<!DOCTYPE html><meta charset="utf-8"><title>OK</title>OK')
open('site/.nojekyll', 'w').close()
print('packed', os.path.getsize('site/data.txt'), 'bytes')
