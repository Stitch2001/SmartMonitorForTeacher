from PyInstaller.__main__ import run
if __name__ == '__main__':
    opts = [r'E:\SmartMonitor for Windows\daily_download\daily_download.py',\
            '-F',r'--distpath=E:\SmartMonitor for Windows\daily_download',\
            r'--workpath=E:\SmartMonitor for Windows\daily_download',\
            r'--specpath=E:\SmartMonitor for Windows\daily_download',\
            r'--upx-dir','upx393w']
    run(opts)