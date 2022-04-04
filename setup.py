from cx_Freeze import setup, Executable

executables = [Executable('main.py', targetName='test.exe')]
include_files = ['firm.db', 'imap.py', 'smtp.py', 'sql.py', 'config.yaml']
excludes = ['http', 'pickle', 'tkinter', 'multiprocessing']

options = {
    'build_exe': {
        'include_msvcr': True,
        'excludes': excludes,
        'include_files': include_files
    }
}

setup(name='test_crm',
      version='0.0.2',
      description='Testing some functions!',
      executables=executables,
      options=options)
