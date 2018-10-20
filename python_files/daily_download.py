# LeanCloud初始化
import leancloud
leancloud.init("kJ4C4D7mWjjAD2X5G3JpPe81-gzGzoHsz", "MwsllyERC65LKHtrq2qE2ifL")
import logging
logging.basicConfig(level=logging.DEBUG)

# （伪）常量定义
SENIOR_1 = 0
SENIOR_2 = 1
SENIOR_3 = 2

# 获取需要抓取的年级
import argparse
ap = argparse.ArgumentParser()
ap.add_argument("-int","--Grade",required=True,help = "Grade")
args = vars(ap.parse_args())
grade = args["Grade"]

# 抓取应到实到数据
ClassData = leancloud.Object.extend('ClassData')
gradeQuery = ClassData.query
isDownloadedQuery = ClassData.query
gradeQuery.equal_to('grade',int(grade))
query_list = gradeQuery.find()

# 抓取扣分数据
SituationData = leancloud.Object.extend('SituationData')
situationQuery = SituationData.query
situationQuery.equal_to('grade',int(grade))
query_list2 = situationQuery.find()

# 处理数据·处理应到实到数据
ought = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
fact = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
leave = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
temporary = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
absent = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
checker = query_list[0].get('checker')
for result in query_list:
    if result.get('isDownloaded') == 1:
        classroom = result.get('classroom')
        ought[classroom] = result.get('ought')
        fact[classroom] = result.get('fact')
        leave[classroom] = result.get('leave')
        temporary[classroom] = result.get('temporary')
        absent[classroom] = result.get('absent')
        #result.set('isDownloaded',2)
        #result.save();

# 处理数据·处理扣分数据
location = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
event = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
score = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
date = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
for result in query_list2:
    if result.get('isDownloaded') == 1:
        classroom = result.get('classroom')
        location[classroom] = result.get('location')
        event[classroom] = result.get('event')
        score[classroom] = result.get('score')
        date[classroom] = result.get('date')
        #result.set('isDownloaded',2)
        #result.save();

# 生成表格
import xlwt