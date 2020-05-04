from subprocess import call
call('CALL conda.bat activate hazus_env & start /min python src\gui.py', shell=True)
