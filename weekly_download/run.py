from PyInstaller.__main__ import run
if __name__ == '__main__':
    opts = [r'E:\SmartMonitor for Windows\weekly_download\weekly_download.py',\
            '-F',r'--distpath=E:\SmartMonitor for Windows\weekly_download',\
            r'--workpath=E:\SmartMonitor for Windows\weekly_download',\
            r'--specpath=E:\SmartMonitor for Windows\weekly_download',\
            r'--upx-dir','upx393w']
    run(opts)