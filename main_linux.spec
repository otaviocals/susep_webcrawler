# -*- mode: python -*-

block_cipher = None

a = Analysis(['main.py'],
             pathex=["/home/cals/Projects/susep_webcrawler","/usr/lib/python3.6/site-packages/kivy"],
             binaries=None,
             datas=[
("libs","libs"),
("logo","logo"),
("kivy","kivy"),
("rlibs","rlibs"),
("rscripts","rscripts"),
("webscraper","webscraper"),
("phantomjs/phantomjs","phantom")],
             hiddenimports=["six","packaging","packaging.version","packaging.specifiers","packaging.requirements","appdirs","kivy","kivy.garden","kivy.app","kivy.uix","cython","pygame","pyenchant","opencv-contrib-python","selenium"],
             hookspath=["."],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='SUSEP_Webscraper',
          debug=False,
          strip=False,
          upx=True,
          icon='logo.ico',
          console=False )
