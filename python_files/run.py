from PyInstaller.__main__ import run
if __name__ == '__main__':
    opts = [r'E:\SmartMonitor for Windows\python_files\daily_download.py',\
            '-F','-w',r'--distpath=E:\SmartMonitor for Windows\python_files',\
            r'--workpath=E:\SmartMonitor for Windows\python_files',\
            r'--specpath=E:\SmartMonitor for Windows\python_files',\
            r'--upx-dir','upx393w']
    run(opts)