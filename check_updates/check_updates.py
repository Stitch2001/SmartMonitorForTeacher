# 初始化LeanCloud
import leancloud
def initLeanCloud():
    leancloud.init("kJ4C4D7mWjjAD2X5G3JpPe81-gzGzoHsz", "MwsllyERC65LKHtrq2qE2ifL")
    import logging
    #logging.basicConfig(level=logging.DEBUG)

CURRENT_VISION = 1; #版本号配置！！

#主程序
print("检查更新...")
initLeanCloud()
updatesForWindows = leancloud.Object.extend("UpdatesW3")
query = updatesForWindows.query
query.greater_than("vision",CURRENT_VISION)
updatesList = query.find()

if updatesList != []:
    print("发现新版本，正在下载...")
    file = updatesList[0].get("file")
    url = file.url
    import requests
    r = requests.get(url)
    with open("updated.exe", "wb") as code:
        code.write(r.content)

    import os
    os.system("updated.exe")
else:
    print("未发现新版本")