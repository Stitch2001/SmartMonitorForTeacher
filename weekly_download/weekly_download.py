# （伪）常量定义
SENIOR_1 = 0
SENIOR_2 = 1
SENIOR_3 = 2
JUNIOR_1 = 3
JUNIOR_2 = 4
JUNIOR_3 = 5

PATTERN_NOON = 0
PATTERN_NIGHT = 1

# 初始化LeanCloud
import leancloud
def initLeanCloud():
    leancloud.init("kJ4C4D7mWjjAD2X5G3JpPe81-gzGzoHsz", "MwsllyERC65LKHtrq2qE2ifL")
    import logging
    #logging.basicConfig(level=logging.DEBUG)

# 生成表格
import xlwt
def MakeExcel(grade,absent,score,beginDate,endDate):

    if grade == SENIOR_1 :
        gradeString = '高一'
    elif grade == SENIOR_2:
        gradeString = '高二'
    else :
        gradeString = '高三'

    excel = xlwt.Workbook(encoding='Utf-8')
    sheet = excel.add_sheet('sheet1')

    #设置列宽
    sheet.col(0).width = 3116 # 1088（列宽）*2.8
    sheet.col(1).width = 2346
    sheet.col(2).width = 2346
    sheet.col(3).width = 2346
    sheet.col(4).width = 2346
    sheet.col(5).width = 2346
    sheet.col(6).width = 2346
    sheet.col(7).width = 2346
    #设置行高
    height_style = xlwt.easyxf('font:height 520;') #26（字号）*20
    sheet.row(0).set_style(height_style)
    i = 1
    height_style = xlwt.easyxf('font:height 400;')
    while i <= 22:
        sheet.row(i).set_style(height_style)
        i += 1

    #标题部分
    style = xlwt.XFStyle()
        #设置字体
    titleFont = xlwt.Font()
    titleFont.bold = True
    titleFont.height = 440 # 22（字号）*20=440
        #设置居中
    alignment = xlwt.Alignment()
    alignment.horz = alignment.HORZ_CENTER
    alignment.vert = alignment.VERT_CENTER
    style.font = titleFont
    style.alignment = alignment
    sheet.write_merge(0, 0, 0, 7, gradeString+'午休晚修扣分情况汇总表', style)

    #表头部分
    font = xlwt.Font()
    font.height = 220
    font.bold = False
    font.name = '宋体'
    style.font = font
    sheet.write_merge(1, 1, 0, 3, '时间：'+beginDate[5:10]+' ~ '+endDate[5:10], style)

    #表头2部分
    border = xlwt.Borders()
    border.left = xlwt.Borders.THIN
    border.left_colour = 0x000000
    border.right = xlwt.Borders.THIN
    border.right_colour = 0x000000
    border.top = xlwt.Borders.THIN
    border.top_colour = 0x000000
    border.bottom = xlwt.Borders.THIN
    border.bottom_colour = 0x000000
    style2 = xlwt.XFStyle()
    style2.font = font
    style2.alignment = alignment
    style2.borders = border
    sheet.write_merge(2, 2, 0, 0, '班级', style2)
    sheet.write_merge(2, 2, 1, 1, '星期一', style2)
    sheet.write_merge(2, 2, 2, 2, '星期二', style2)
    sheet.write_merge(2, 2, 3, 3, '星期三', style2)
    sheet.write_merge(2, 2, 4, 4, '星期四', style2)
    sheet.write_merge(2, 2, 5, 5, '星期五', style2)
    sheet.write_merge(2, 2, 6, 6, '星期六', style2)
    sheet.write_merge(2, 2, 7, 7, '星期日', style2)

    #班级数据部分
    i = 3
    j = 1
    while j <= 18 :
        sheet.write_merge(i, i, 0, 0, gradeString+'（'+str(j)+'）', style2)
        if (absent[1][j]+score[1][j] != 0) :
            sheet.write_merge(i, i, 1, 1, '-'+str(absent[1][j]+score[1][j]), style2)
        else : sheet.write_merge(i, i, 1, 1, '', style2)
        if (absent[2][j]+score[2][j] != 0) :
            sheet.write_merge(i, i, 2, 2, '-'+str(absent[2][j]+score[2][j]), style2)
        else : sheet.write_merge(i, i, 2, 2, '', style2)
        if (absent[3][j]+score[3][j] != 0) :
            sheet.write_merge(i, i, 3, 3, '-'+str(absent[3][j]+score[3][j]), style2)
        else : sheet.write_merge(i, i, 3, 3, '', style2)
        if (absent[4][j]+score[4][j] != 0) :
            sheet.write_merge(i, i, 4, 4, '-'+str(absent[4][j]+score[4][j]), style2)
        else : sheet.write_merge(i, i, 4, 4, '', style2)
        if (absent[5][j]+score[5][j] != 0) :
            sheet.write_merge(i, i, 5, 5, '-'+str(absent[5][j]+score[5][j]), style2)
        else : sheet.write_merge(i, i, 5, 5, '', style2)
        if (absent[6][j]+score[6][j] != 0) :
            sheet.write_merge(i, i, 6, 6, '-'+str(absent[6][j]+score[6][j]), style2)
        else : sheet.write_merge(i, i, 6, 6, '', style2)
        if (absent[7][j]+score[7][j] != 0) :
            sheet.write_merge(i, i, 7, 7, '-'+str(absent[7][j]+score[7][j]), style2)
        else : sheet.write_merge(i, i, 7, 7, '', style2)
        i += 1
        j += 1

    #对三个年级的19/20号班级分类讨论
    if grade == SENIOR_1 :
        i = 0
        while i <= 7 :
            sheet.write_merge(21, 21, i, i, '', style2)
            sheet.write_merge(22, 22, i, i, '', style2)
            i += 1
    elif grade == SENIOR_2:
        # 抓取初二应到实到数据
        ClassData = leancloud.Object.extend('ClassData')
        gradeQuery = ClassData.query
        gradeQuery.equal_to('grade', JUNIOR_2)
        query_list = gradeQuery.find()
        # 处理应到实到数据
        absent1 = [[0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0]]
        isFirstResult = True
        lastDayOfWeek = 1
        for result in query_list:
            if result.get('isWeeklyDownloaded') == False:
                dayOfWeek = result.get('dayOfWeek')
                if ((dayOfWeek - lastDayOfWeek) > 1):  # 如果时间间隔大于1天，则说明已经是下一周
                    if (not isFirstResult):
                        break
                elif ((dayOfWeek - lastDayOfWeek) < 0):  # 如果有时间间隔小于0的情况，则必定是从周日到周一，否则就是下一周了
                    if (dayOfWeek != 1 or lastDayOfWeek != 7):
                        break
                classroom = result.get('classroom')
                absent1[dayOfWeek][classroom] = result.get('absent')
                lastDayOfWeek = dayOfWeek
                endDate = str(result.get('createdAt'))[0:10]
                if isFirstResult:
                    isFirstResult = False
                    beginDate = str(result.get('createdAt'))
                # result.set('isWeeklyDownloaded',True)
                # result.save();
        # 抓取扣分数据
        SituationData = leancloud.Object.extend('SituationData')
        situationQuery = SituationData.query
        situationQuery.equal_to('grade', JUNIOR_2)
        query_list2 = situationQuery.find()
        # 处理扣分数据
        score1 = [[0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0]]
        isFirstResult1 = True
        lastDayOfWeek = 1
        for result in query_list2:
            if result.get('isWeeklyDownloaded') == False:
                dayOfWeek = result.get('dayOfWeek')
                if ((dayOfWeek - lastDayOfWeek) > 1):  # 如果时间间隔大于1天，则说明已经是下一周
                    if (not isFirstResult1):
                        break
                elif ((dayOfWeek - lastDayOfWeek) < 0):  # 如果有时间间隔小于0的情况，则必定是从周日到周一，否则就是下一周了
                    if (dayOfWeek != 1 or lastDayOfWeek != 7):
                        break
                classroom = result.get('classroom')
                score1[dayOfWeek][classroom] += result.get('score')
                print(score1[dayOfWeek][classroom])
                lastDayOfWeek = dayOfWeek
                isFirstResult1 = False
                # result.set('isWeeklyDownloaded',True)
                # result.save();
        #写入表格
        sheet.write_merge(21, 21, 0, 0, '初二（1）', style2)
        if (absent1[1][1] + score1[1][1] != 0):
            sheet.write_merge(21, 21, 1, 1, '-' + str(absent1[1][1] + score1[1][1]), style2)
        else:
            sheet.write_merge(21, 21, 1, 1, '', style2)
        if (absent1[2][1] + score1[2][1] != 0):
            sheet.write_merge(21, 21, 2, 2, '-' + str(absent1[2][1] + score1[2][1]), style2)
        else:
            sheet.write_merge(21, 21, 2, 2, '', style2)
        if (absent1[3][1] + score1[3][1] != 0):
            sheet.write_merge(21, 21, 3, 3, '-' + str(absent1[3][1] + score1[3][1]), style2)
        else:
            sheet.write_merge(21, 21, 3, 3, '', style2)
        if (absent1[4][1] + score1[4][1] != 0):
            sheet.write_merge(21, 21, 4, 4, '-' + str(absent1[4][1] + score1[4][1]), style2)
        else:
            sheet.write_merge(21, 21, 4, 4, '', style2)
        if (absent1[5][1] + score1[5][1] != 0):
            sheet.write_merge(21, 21, 5, 5, '-' + str(absent1[5][1] + score1[5][1]), style2)
        else:
            sheet.write_merge(21, 21, 5, 5, '', style2)
        if (absent1[6][1] + score1[6][1] != 0):
            sheet.write_merge(21, 21, 6, 6, '-' + str(absent1[6][1] + score1[6][1]), style2)
        else:
            sheet.write_merge(21, 21, 6, 6, '', style2)
        if (absent1[7][1] + score1[7][1] != 0):
            sheet.write_merge(21, 21, 7, 7, '-' + str(absent1[7][1] + score1[7][1]), style2)
        else:
            sheet.write_merge(21, 21, 7, 7, '', style2)

        sheet.write_merge(22, 22, 0, 0, '初二（2）', style2)
        if (absent1[1][2] + score1[1][2] != 0):
            sheet.write_merge(22, 22, 1, 1, '-' + str(absent1[1][2] + score1[1][2]), style2)
        else:
            sheet.write_merge(22, 22, 1, 1, '', style2)
        if (absent1[2][2] + score1[2][2] != 0):
            sheet.write_merge(22, 22, 2, 2, '-' + str(absent1[2][2] + score1[2][2]), style2)
        else:
            sheet.write_merge(22, 22, 2, 2, '', style2)
        if (absent1[3][2] + score1[3][2] != 0):
            sheet.write_merge(22, 22, 3, 3, '-' + str(absent1[3][2] + score1[3][2]), style2)
        else:
            sheet.write_merge(22, 22, 3, 3, '', style2)
        if (absent1[4][2] + score1[4][2] != 0):
            sheet.write_merge(22, 22, 4, 4, '-' + str(absent1[4][2] + score1[4][2]), style2)
        else:
            sheet.write_merge(22, 22, 4, 4, '', style2)
        if (absent1[5][2] + score1[5][2] != 0):
            sheet.write_merge(22, 22, 5, 5, '-' + str(absent1[5][2] + score1[5][2]), style2)
        else:
            sheet.write_merge(22, 22, 5, 5, '', style2)
        if (absent1[6][2] + score1[6][2] != 0):
            sheet.write_merge(22, 22, 6, 6, '-' + str(absent1[6][2] + score1[6][2]), style2)
        else:
            sheet.write_merge(22, 22, 6, 6, '', style2)
        if (absent1[7][2] + score1[7][2] != 0):
            sheet.write_merge(22, 22, 7, 7, '-' + str(absent1[7][2] + score1[7][2]), style2)
        else:
            sheet.write_merge(22, 22, 7, 7, '', style2)

    elif grade == SENIOR_3:
        # 抓取初三应到实到数据
        ClassData = leancloud.Object.extend('ClassData')
        gradeQuery = ClassData.query
        gradeQuery.equal_to('grade', JUNIOR_3)
        query_list = gradeQuery.find()
        # 处理应到实到数据
        absent1 = [[0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0]]
        isFirstResult = True
        lastDayOfWeek = 1
        for result in query_list:
            if result.get('isWeeklyDownloaded') == False:
                dayOfWeek = result.get('dayOfWeek')
                if ((dayOfWeek - lastDayOfWeek) > 1):  # 如果时间间隔大于1天，则说明已经是下一周
                    if (not isFirstResult):
                        break
                elif ((dayOfWeek - lastDayOfWeek) < 0):  # 如果有时间间隔小于0的情况，则必定是从周日到周一，否则就是下一周了
                    if (dayOfWeek != 1 or lastDayOfWeek != 7):
                        break
                classroom = result.get('classroom')
                absent1[dayOfWeek][classroom] = result.get('absent')
                lastDayOfWeek = dayOfWeek
                endDate = str(result.get('createdAt'))[0:10]
                if isFirstResult:
                    isFirstResult = False
                    beginDate = str(result.get('createdAt'))
                # result.set('isWeeklyDownloaded',True)
                # result.save();
        # 抓取扣分数据
        SituationData = leancloud.Object.extend('SituationData')
        situationQuery = SituationData.query
        situationQuery.equal_to('grade', JUNIOR_3)
        query_list2 = situationQuery.find()
        # 处理扣分数据
        score1 = [[0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0]]
        isFirstResult1 = True
        lastDayOfWeek = 1
        for result in query_list2:
            if result.get('isWeeklyDownloaded') == False:
                dayOfWeek = result.get('dayOfWeek')
                if ((dayOfWeek - lastDayOfWeek) > 1):  # 如果时间间隔大于1天，则说明已经是下一周
                    if (not isFirstResult1):
                        break
                elif ((dayOfWeek - lastDayOfWeek) < 0):  # 如果有时间间隔小于0的情况，则必定是从周日到周一，否则就是下一周了
                    if (dayOfWeek != 1 or lastDayOfWeek != 7):
                        break
                classroom = result.get('classroom')
                score1[dayOfWeek][classroom] += result.get('score')
                print(score1[dayOfWeek][classroom])
                lastDayOfWeek = dayOfWeek
                isFirstResult1 = False
                # result.set('isWeeklyDownloaded',True)
                # result.save();
        #写入表格
        sheet.write_merge(21, 21, 0, 0, '初三（1）', style2)
        if (absent1[1][1] + score1[1][1] != 0):
            sheet.write_merge(21, 21, 1, 1, '-' + str(absent1[1][1] + score1[1][1]), style2)
        else:
            sheet.write_merge(21, 21, 1, 1, '', style2)
        if (absent1[2][1] + score1[2][1] != 0):
            sheet.write_merge(21, 21, 2, 2, '-' + str(absent1[2][1] + score1[2][1]), style2)
        else:
            sheet.write_merge(21, 21, 2, 2, '', style2)
        if (absent1[3][1] + score1[3][1] != 0):
            sheet.write_merge(21, 21, 3, 3, '-' + str(absent1[3][1] + score1[3][1]), style2)
        else:
            sheet.write_merge(21, 21, 3, 3, '', style2)
        if (absent1[4][1] + score1[4][1] != 0):
            sheet.write_merge(21, 21, 4, 4, '-' + str(absent1[4][1] + score1[4][1]), style2)
        else:
            sheet.write_merge(21, 21, 4, 4, '', style2)
        if (absent1[5][1] + score1[5][1] != 0):
            sheet.write_merge(21, 21, 5, 5, '-' + str(absent1[5][1] + score1[5][1]), style2)
        else:
            sheet.write_merge(21, 21, 5, 5, '', style2)
        if (absent1[6][1] + score1[6][1] != 0):
            sheet.write_merge(21, 21, 6, 6, '-' + str(absent1[6][1] + score1[6][1]), style2)
        else:
            sheet.write_merge(21, 21, 6, 6, '', style2)
        if (absent1[7][1] + score1[7][1] != 0):
            sheet.write_merge(21, 21, 7, 7, '-' + str(absent1[7][1] + score1[7][1]), style2)
        else:
            sheet.write_merge(21, 21, 7, 7, '', style2)

        sheet.write_merge(22, 22, 0, 0, '初三（2）', style2)
        if (absent1[1][2] + score1[1][2] != 0):
            sheet.write_merge(22, 22, 1, 1, '-' + str(absent1[1][2] + score1[1][2]), style2)
        else:
            sheet.write_merge(22, 22, 1, 1, '', style2)
        if (absent1[2][2] + score1[2][2] != 0):
            sheet.write_merge(22, 22, 2, 2, '-' + str(absent1[2][2] + score1[2][2]), style2)
        else:
            sheet.write_merge(22, 22, 2, 2, '', style2)
        if (absent1[3][2] + score1[3][2] != 0):
            sheet.write_merge(22, 22, 3, 3, '-' + str(absent1[3][2] + score1[3][2]), style2)
        else:
            sheet.write_merge(22, 22, 3, 3, '', style2)
        if (absent1[4][2] + score1[4][2] != 0):
            sheet.write_merge(22, 22, 4, 4, '-' + str(absent1[4][2] + score1[4][2]), style2)
        else:
            sheet.write_merge(22, 22, 4, 4, '', style2)
        if (absent1[5][2] + score1[5][2] != 0):
            sheet.write_merge(22, 22, 5, 5, '-' + str(absent1[5][2] + score1[5][2]), style2)
        else:
            sheet.write_merge(22, 22, 5, 5, '', style2)
        if (absent1[6][2] + score1[6][2] != 0):
            sheet.write_merge(22, 22, 6, 6, '-' + str(absent1[6][2] + score1[6][2]), style2)
        else:
            sheet.write_merge(22, 22, 6, 6, '', style2)
        if (absent1[7][2] + score1[7][2] != 0):
            sheet.write_merge(22, 22, 7, 7, '-' + str(absent1[7][2] + score1[7][2]), style2)
        else:
            sheet.write_merge(22, 22, 7, 7, '', style2)

    path = GetDesktopPath()
    excel.save(path+"\\"+beginDate[5:10]+' ~ '+endDate[5:10]+gradeString+"午休晚修情况汇总表.xls")

# 获取桌面路径
import os
def GetDesktopPath():
    return os.path.join(os.path.expanduser("~"), 'Desktop')

# 获取需要抓取的年级
def getGrade():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("-intA", "--Grade", required=True, help="Grade")
    args = vars(ap.parse_args())
    grade = args["Grade"]
    return grade

#主程序
print("获取周汇总表...")
initLeanCloud()
grade = getGrade()

# 抓取应到实到数据
ClassData = leancloud.Object.extend('ClassData')
gradeQuery = ClassData.query
gradeQuery.equal_to('grade',int(grade))
isDownloadedQuery = ClassData.query
isDownloadedQuery.does_not_exist("isWeeklyDownloaded")
query = leancloud.Query.and_(gradeQuery,isDownloadedQuery)
query_list = query.find()

# 抓取扣分数据
SituationData = leancloud.Object.extend('SituationData')
situationQuery = SituationData.query
situationQuery.equal_to('grade',int(grade))
isDownloadedQuery = SituationData.query
isDownloadedQuery.does_not_exist("isWeeklyDownloaded")
query = leancloud.Query.and_(situationQuery,isDownloadedQuery)
query_list2 = query.find()

# 处理数据·处理应到实到数据
absent = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]
isFirstResult = True
lastDayOfWeek = 1
for result in query_list:
    if result.get('isWeeklyDownloaded') == False:
        dayOfWeek = result.get('dayOfWeek')
        if ((dayOfWeek - lastDayOfWeek) > 1):# 如果时间间隔大于1天，则说明已经是下一周
            if (not isFirstResult) :
                break
        elif ((dayOfWeek - lastDayOfWeek) < 0):# 如果有时间间隔小于0的情况，则必定是从周日到周一，否则就是下一周了
            if (dayOfWeek != 1 or lastDayOfWeek != 7) :
                break
        classroom = result.get('classroom')
        absent[dayOfWeek][classroom] = result.get('absent')
        print(absent[dayOfWeek][classroom])
        lastDayOfWeek = dayOfWeek
        endDate = str(result.get('createdAt'))[0:10]
        if isFirstResult :
            isFirstResult = False
            beginDate = str(result.get('createdAt'))
        #result.set('isWeeklyDownloaded',True)
        #result.save();


# 处理数据·处理扣分数据
score = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]
isFirstResult1 = True
lastDayOfWeek = 1
for result in query_list2:
    if result.get('isWeeklyDownloaded') == False:
        dayOfWeek = result.get('dayOfWeek')
        if ((dayOfWeek - lastDayOfWeek) > 1):# 如果时间间隔大于1天，则说明已经是下一周
            if (not isFirstResult1) :
                break
        elif ((dayOfWeek - lastDayOfWeek) < 0):# 如果有时间间隔小于0的情况，则必定是从周日到周一，否则就是下一周了
            if (dayOfWeek != 1 or lastDayOfWeek != 7) :
                break
        classroom = result.get('classroom')
        score[dayOfWeek][classroom] += result.get('score')
        print(score[dayOfWeek][classroom])
        lastDayOfWeek = dayOfWeek
        isFirstResult1 = False
        #result.set('isWeeklyDownloaded',True)
        #result.save();

if isFirstResult == False :
    MakeExcel(int(grade),absent,score,beginDate,endDate)
    import tkinter.messagebox
    tkinter.messagebox.showinfo("智慧纪检", "成功！请到桌面查看")
else :
    import tkinter.messagebox
    tkinter.messagebox.showinfo("智慧纪检", "没有待收集的检查表")