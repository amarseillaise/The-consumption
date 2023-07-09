from cx_Freeze import setup, Executable

executables = [Executable('The Consumption.py',
                          targetName='The Consumption.exe',
                          base='Win32GUI',
                          icon='Icon.ico',
                          shortcutName='The Consumption',
                          shortcutDir='ProgramMenuFolder')]

excludes = ['logging', 'unittest', 'email', 'html', 'http',
            'unicodedata', 'bz2', 'select']

zip_include_packages = ['collections', 'encodings', 'importlib', 'wx']

options = {
    'build_exe': {
        'include_msvcr': True,
        'excludes': excludes,
        'zip_include_packages': zip_include_packages,
    }
}

setup(name='The Consumption',
      version='1.0',
      description='Считает расход',
      executables=executables,
      options=options)