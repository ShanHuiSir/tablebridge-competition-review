# Run this specification on Windows to create a Windows executable.
from PyInstaller.utils.hooks import collect_all

wv_data, wv_bins, wv_hidden = collect_all('webview')
a = Analysis(['desktop.py'], pathex=['.'], binaries=wv_bins,
             datas=[('static', 'static'), ('samples', 'samples')] + wv_data,
             hiddenimports=wv_hidden + ['xlrd', 'xlwt', 'openpyxl', 'sandbox'],
             hookspath=[], runtime_hooks=[], excludes=[])
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='TableBridge',
          debug=False, bootloader_ignore_signals=False, strip=False,
          upx=False, console=True)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='TableBridge')
