# -*- mode: python -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['/home/cals/Projects/susep_webcrawler'],
             binaries=None,
             datas=[
("libs","libs"),
("logo","logo"),
("kivy","kivy"),
("rlibs","rlibs"),
("rscripts","rscripts"),
("webscraper","webscraper"),
("phantomjs/phantomjs.exe","phantom")],
             hiddenimports=["six","packaging","packaging.version","packaging.specifiers","packaging.requirements","appdirs"],
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
          name='Webscraper de Taxa de Juros',
          debug=False,
          strip=False,
          upx=True,
	  icon='logo.ico',
          console=True )
