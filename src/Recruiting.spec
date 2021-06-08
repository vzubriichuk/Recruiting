# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['Recruiting.py'],
             pathex=['.'],
             binaries=[],
             datas=[( '.\\resources\\*.ico', 'resources' ),
                    ( '.\\resources\\*.pdf', 'resources' ),
                    ( '.\\resources\\*.png', 'resources' ),
                    ( '.\\resources\\*.docx', 'resources' ),
                    ( '.\\resources\\*.txt', 'resources' )],
             hiddenimports=['babel.numbers'],
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
          name='Recruiting',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='resources\\file.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='Recruiting')
