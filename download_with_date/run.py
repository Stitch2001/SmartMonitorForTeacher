from PyInstaller.__main__ import run
if __name__ == '__main__':
    opts = [r'E:\SmartMonitor for Windows\download_with_date\download_with_date.py',\
            '-F',r'--distpath=E:\SmartMonitor for Windows\download_with_date',\
            r'--workpath=E:\SmartMonitor for Windows\download_with_date',\
            r'--specpath=E:\SmartMonitor for Windows\download_with_date',\
            r'--upx-dir','upx393w']
    run(opts)