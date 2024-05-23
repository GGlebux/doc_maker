block_cipher = None

a = Analysis(['main.py'],
             pathex=[],
             binaries=[],
             datas=[('C:/Users/Gecko/Desktop/doc_maker/template.docx', '.')],
             hiddenimports=['tkinter', 'tkinter.filedialog', 'docxtpl', 'docx', 'docxcompose', 'openpyxl'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,a.scripts,a.binaries,a.zipfiles,a.datas,
          name='Pensil',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
                a.binaries,
                a.zipfiles,
                a.datas,
                strip=False,
                upx=True,
                upx_exclude=[],
                name='dist')
