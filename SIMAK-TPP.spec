# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('files', 'files'),
    ],
    hiddenimports=['secrets', 'hmac', '_hashlib', '_blake2'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Stdlib berat yang tidak dipakai
        'unittest', 'test', 'distutils', 'setuptools', 'pkg_resources',
        'xmlrpc', 'ftplib', 'imaplib',
        'poplib', 'smtplib', 'telnetlib', 'nntplib',
        'curses',
        'doctest', 'pdb', 'profile', 'pstats', 'cProfile',
        'lib2to3', 'tkinter.test', 'idlelib',
        # Submodul numpy berat yang tidak dipakai
        'numpy.core._multiarray_tests',
        'numpy.random._examples',
        # Submodul pandas yang tidak dipakai di GUI
        'pandas.tests', 'pandas.io.formats.style',
        # Modul Pillow yang tidak dipakai
        'PIL.ImageFilter', 'PIL.ImageDraw2',
        # Selenium internal yang jarang dipakai
        'selenium.webdriver.firefox',
        'selenium.webdriver.safari',
        'selenium.webdriver.ie',
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SIMAK-TPP',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll', 'msvcp140.dll', 'tcl86t.dll', 'tk86t.dll', '_tkinter.pyd', 'python3*.dll'],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo-pemprov.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=['tcl86t.dll', 'tk86t.dll', '_tkinter.pyd'],
    name='SIMAK-TPP',
)
