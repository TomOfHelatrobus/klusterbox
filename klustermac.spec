# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['klusterbox.py'],
             pathex=['/Users/thomasweeks/klusterbox/kb_install'],
             binaries=[],
             datas=[
		('/Users/thomasweeks/klusterbox/kb_install/kb_sub/kb_images/kb_about.jpg','.'),
		('/Users/thomasweeks/klusterbox/kb_install/kb_sub/kb_images/kb_icon1.icns','.'),
		('/Users/thomasweeks/klusterbox/kb_install/kb_sub/kb_images/kb_icon2.gif','.'),
		('/Users/thomasweeks/klusterbox/kb_install/kb_sub/kb_images/kb_icon2.ico','.'),
		('/Users/thomasweeks/klusterbox/kb_install/kb_sub/kb_images/kb_icon2.jpg','.')],
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
          console=False , icon='kb_sub/kb_images/kb_icon1.icns')
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
             icon='kb_sub/kb_images/kb_icon1.icns',
             bundle_identifier=None)