# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['klusterbox.py'],
             pathex=['/Users/thomas/klusterbox/kb_install'],
             binaries=[],
             datas=[
		('/Users/thomas/klusterbox/kb_sub/kb_images/kb_about.jpg','.'),
		('/Users/thomas/klusterbox/kb_sub/kb_images/kb_icon2.gif','.'),
		('/Users/thomas/klusterbox/kb_sub/kb_images/kb_icon2.ico','.'),
		('/Users/thomas/klusterbox/kb_sub/kb_images/kb_icon2.jpg','.'),
		('/Users/thomas/klusterbox/history.txt','.'),
		('/Users/thomas/klusterbox/readme.txt','.'),
		('/Users/thomas/klusterbox/LICENSE.txt','.'),
		('/Users/thomas/klusterbox/cheatsheet.txt','.'),
		('/Users/thomas/klusterbox/speedsheet_instructions.txt','.'),
		('/Users/thomas/klusterbox/klusterbox.py','.'),
		('/Users/thomas/klusterbox/projvar.py','.'),
		('/Users/thomas/klusterbox/kbtoolbox.py','.'),
		('/Users/thomas/klusterbox/kbdatabase.py','.'),
		('/Users/thomas/klusterbox/kbreports.py','.'),
		('/Users/thomas/klusterbox/kbspreadsheets.py','.'),
		('/Users/thomas/klusterbox/kbspeedsheets.py','.'),							('/Users/thomas/klusterbox/kbequitability.py','.'),
		('/Users/thomas/klusterbox/requirements.txt','.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='klusterbox',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='kb_sub/kb_images/kb_icon2.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='klusterbox')
app = BUNDLE(coll,
             name='klusterbox.app',
             icon='kb_sub/kb_images/kb_icon2.ico',
             bundle_identifier=None)