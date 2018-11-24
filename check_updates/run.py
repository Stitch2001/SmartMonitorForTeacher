from PyInstaller.__main__ import run
if __name__ == '__main__':
    opts = [r'E:\SmartMonitor for Windows\check_updates\check_updates.py',\
            '-F','-w',r'--distpath=E:\SmartMonitor for Windows\check_updates',\
            r'--workpath=E:\SmartMonitor for Windows\check_updates',\
            r'--specpath=E:\SmartMonitor for Windows\check_updates',\
            r'--upx-dir','upx393w']
    run(opts)