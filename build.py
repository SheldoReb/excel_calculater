import subprocess

py_file = 'excel_calculater.py'
exe_name = 'ExcelCalculater'

# Define the PyInstaller command
pyinstaller_cmd = f'pyinstaller --name={exe_name} --onefile --windowed --clean --strip --noupx --noconsole {py_file}'

# Run the command
subprocess.run(pyinstaller_cmd, shell=True)
