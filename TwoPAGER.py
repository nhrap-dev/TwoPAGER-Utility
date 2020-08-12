from subprocess import call
try:
    try:
        call('CALL conda.bat activate hazus_env & start /min python src\gui.py', shell=True)
    except:
        call('start /min python src\gui.py', shell=True)

except:
    import ctypes
    import sys
    messageBox = ctypes.windll.user32.MessageBoxW
    error = sys.exc_info()[0]
    messageBox(0, u"Unexpected error: {er} | If this problem persists, and you've ensured it's not user error, ask your developers to write better code.".format(
        er=error), u"HazPy", 0x1000)
