# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['klusterbox.py'],
             pathex=['/Users/toman/klusterbox/kb_install'],
             binaries=[],
             datas=[
		('/Users/toman/klusterbox/kb_install/kb_sub/kb_images/kb_about.jpg','.'),
		('/Users/toman/klusterbox/kb_install/kb_sub/kb_images/kb_icon2.gif','.'),
		('/Users/toman/klusterbox/kb_install/kb_sub/kb_images/kb_icon2.ico','.'),
		('/Users/toman/klusterbox/kb_install/kb_sub/kb_images/kb_icon2.jpg','.'),
		('/Users/toman/klusterbox/history.txt','.'),
		('/Users/toman/klusterbox/readme.txt','.'),
		('/Users/toman/klusterbox/LICENSE.txt','.'),
		('/Users/toman/klusterbox/klusterbox.py','.'),
		('/Users/toman/klusterbox/requirements.txt','.')],
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