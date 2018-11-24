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
def MakeExcel(grade):

    if grade == SENIOR_1 :
        gradeString = '高一'
    elif grade == SENIOR_2:
        gradeString = '高二'
    else :
        gradeString = '高三'

    excel = xlwt.Workbook(encoding='Utf-8')
    sheet = excel.add_sheet('sheet1')

    #设置列宽
    sheet.col(0).width = 3048 # 1088（列宽）*2.8
    sheet.col(1).width = 2346
    sheet.col(2).width = 3010
    sheet.col(3).width = 3186
    sheet.col(4).width = 12250
    #设置行高
    height_style = xlwt.easyxf('font:height 480;')
    sheet.row(0).set_style(height_style)
    height_style = xlwt.easyxf('font:height 440;')
    sheet.row(1).set_style(height_style)
    height_style = xlwt.easyxf('font:height 220;')
    sheet.row(2).set_style(height_style)
    height_style = xlwt.easyxf('font:height 220;')
    sheet.row(3).set_style(height_style)
    i = 4
    height_style = xlwt.easyxf('font:height 440;')
    while i <= 23:
        sheet.row(i).set_style(height_style)
        i += 1;

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
    sheet.write_merge(0, 0, 0, 4, gradeString+'课室'+patString+'情况登记表', style)

    #表头部分
    font = xlwt.Font()
    font.height = 220
    font.bold = False
    font.name = '宋体'
    style.font = font
    sheet.write_merge(1, 1, 0, 3, '检查时间：'+str(time)[0:10]+' 星期'+dayOfWeek, style)
    sheet.write_merge(1, 1, 4, 4, ' 检查人：'+checker, style)

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
    sheet.write_merge(2, 3, 0, 0, '班级', style2)
    sheet.write_merge(2, 3, 1, 1, patString+'人数', style2)
    sheet.write_merge(2, 2, 2, 3, '检查内容', style2)
    sheet.write_merge(3, 3, 2, 2, '到位（4）', style2)
    sheet.write_merge(3, 3, 3, 3, '纪律（6）', style2)
    sheet.write_merge(2, 3, 4, 4, '备注', style2)

    #班级数据部分
    i = 4;j = 1;
    while j <= 18 :
        sheet.write_merge(i, i, 0, 0, gradeString+'（'+str(j)+'）', style2)
        sheet.write_merge(i, i, 1, 1, str(fact[j])+' / '+str(ought[j]), style2)
        situationString = ''
        if absent[j] != 0 :
            sheet.write_merge(i, i, 2, 2, '-'+str(absent[j]), style2)
            situationString += '缺席' +str(absent[j])+'人 '
        else :
            sheet.write_merge(i, i, 2, 2, '', style2)
        if score[j] != 0:
            sheet.write_merge(i, i, 3, 3, '-' + str(score[j]), style2)
        else :
            sheet.write_merge(i, i, 3, 3, '', style2)
        if leave[j] != 0 :
            situationString += '请假'+str(leave[j])+'人 '
        if temporary[j] != 0 :
            situationString += '临'+patString[1::1]+str(temporary[j])+'人 '
        if event[j] != 0 :
            situationString += event[j]
        if situationString != '':
            sheet.write_merge(i, i, 4 ,4 ,situationString,style2)
        else : sheet.write_merge(i, i, 4, 4, '', style2)
        i += 1
        j += 1

    #对三个年级的19/20号班级分类讨论
    if grade == SENIOR_1 :
        i = 0
        while i <= 4 :
            sheet.write_merge(22, 22, i, i, '', style2)
            sheet.write_merge(23, 23, i, i, '', style2)
            i += 1
    elif grade == SENIOR_2:
        # 抓取初二应到实到数据
        ClassData = leancloud.Object.extend('ClassData')
        gradeQuery = ClassData.query
        isDownloadedQuery = ClassData.query
        gradeQuery.equal_to('grade', JUNIOR_2)
        isDownloadedQuery.does_not_exist("isDailyDownloaded")
        patternQuery = ClassData.query
        patternQuery.equal_to('pattern',int(pat))
        query = leancloud.Query.and_(gradeQuery, isDownloadedQuery, patternQuery)
        query_list = query.find()
        # 处理应到实到数据
        ought1 = [0, 0, 0]
        fact1 = [0, 0, 0]
        leave1 = [0, 0, 0]
        temporary1 = [0, 0, 0]
        absent1 = [0, 0, 0]
        for result in query_list:
            classroom = result.get('classroom')
            ought1[classroom] = result.get('ought')
            fact1[classroom] = result.get('fact')
            leave1[classroom] = result.get('leave')
            temporary1[classroom] = result.get('temporary')
            absent1[classroom] = result.get('absent')
            print(JUNIOR_2,classroom,ought1[classroom], fact1[classroom], temporary1[classroom], absent1[classroom])
            # result.set('isDailyDownloaded',True)
            # result.save();
        # 抓取扣分数据
        SituationData = leancloud.Object.extend('SituationData')
        situationQuery = SituationData.query
        isDownloadedQuery = SituationData.query
        situationQuery.equal_to('grade', JUNIOR_2)
        patternQuery = SituationData.query
        patternQuery.equal_to('pattern', int(pat))
        isDownloadedQuery.equal_to("isDailyDownloaded",False)
        query2 = leancloud.Query.and_(situationQuery, isDownloadedQuery, patternQuery)
        query_list2 = query2.find()
        # 处理扣分数据
        event1 = ["", "", ""]
        score1 = [0, 0, 0]
        for result in query_list2:
            classroom = result.get('classroom')
            event1[classroom] = event[classroom]+result.get('location')+result.get('event')+'('+result.get('date')[11:-3]+')'
            score1[classroom] = score[classroom]+result.get('score')
            print(event1[classroom])
            #result.set('isDailyDownloaded',2)
            #result.save();
        #写入表格
        sheet.write_merge(22, 22, 0, 0, '初二（1）', style2)
        sheet.write_merge(22, 22, 1, 1, str(fact1[1])+' / '+str(ought1[1]), style2)
        situationString = ''
        if absent1[1] != 0 :
            sheet.write_merge(22, 22, 2, 2, '-'+str(absent1[1]), style2)
            situationString += '缺席' +str(absent1[1])+'人 '
        else :
            sheet.write_merge(22, 22, 2, 2, '', style2)
        if score1[1] != 0:
            sheet.write_merge(22, 22, 3, 3, '-' + str(score1[1]), style2)
        else :
            sheet.write_merge(22, 22, 3, 3, '', style2)
        if leave1[1] != 0 :
            situationString += '请假'+str(leave1[1])+'人 '
        if temporary1[1] != 0 :
            situationString += '临'+patString[1::1]+str(temporary[j])+'人 '
        if event1[1] != 0 :
            situationString += event1[1]
        if situationString != '':
            sheet.write_merge(22, 22, 4 ,4 ,situationString,style2)
        else :
            sheet.write_merge(22, 22, 4, 4, '', style2)

        sheet.write_merge(23, 23, 0, 0, '初二（2）', style2)
        sheet.write_merge(23, 23, 1, 1, str(fact1[2])+' / '+str(ought1[2]), style2)
        situationString = ''
        if absent1[2] != 0 :
            sheet.write_merge(23, 23, 2, 2, '-'+str(absent1[2]), style2)
            situationString += '缺席' +str(absent1[2])+'人 '
        else :
            sheet.write_merge(23, 23, 2, 2, '', style2)
        if score1[2] != 0:
            sheet.write_merge(23, 23, 3, 3, '-' + str(score1[2]), style2)
        else :
            sheet.write_merge(23, 23, 3, 3, '', style2)
        if leave1[2] != 0 :
            situationString += '请假'+str(leave1[2])+'人 '
        if temporary1[2] != 0 :
            situationString += '临'+patString[1::1]+str(temporary[j])+'人 '
        if event1[2] != 0 :
            situationString += event1[2]
        if situationString != '':
            sheet.write_merge(23, 23, 4 ,4 ,situationString,style2)
        else :
            sheet.write_merge(23, 23, 4, 4, '', style2)
    elif grade == SENIOR_3:
        # 抓取初三应到实到数据
        ClassData = leancloud.Object.extend('ClassData')
        gradeQuery = ClassData.query
        isDownloadedQuery = ClassData.query
        gradeQuery.equal_to('grade', JUNIOR_3)
        isDownloadedQuery.does_not_exist("isDailyDownloaded")
        patternQuery = ClassData.query
        patternQuery.equal_to('pattern',int(pat))
        query = leancloud.Query.and_(gradeQuery, isDownloadedQuery, patternQuery)
        query_list = query.find()
        # 处理应到实到数据
        ought1 = [0, 0, 0]
        fact1 = [0, 0, 0]
        leave1 = [0, 0, 0]
        temporary1 = [0, 0, 0]
        absent1 = [0, 0, 0]
        for result in query_list:
            classroom = result.get('classroom')
            ought1[classroom] = result.get('ought')
            fact1[classroom] = result.get('fact')
            leave1[classroom] = result.get('leave')
            temporary1[classroom] = result.get('temporary')
            absent1[classroom] = result.get('absent')
            print(JUNIOR_3,classroom,ought1[classroom], fact1[classroom], temporary1[classroom], absent1[classroom])
            # result.set('isDailyDownloaded',True)
            # result.save();
        # 抓取扣分数据
        SituationData = leancloud.Object.extend('SituationData')
        situationQuery = SituationData.query
        isDownloadedQuery = SituationData.query
        situationQuery.equal_to('grade',JUNIOR_3)
        patternQuery = SituationData.query
        patternQuery.equal_to('pattern',int(pat))
        isDownloadedQuery.equal_to("isDailyDownloaded",False)
        query2 = leancloud.Query.and_(situationQuery,isDownloadedQuery,patternQuery)
        query_list2 = query2.find()
        # 处理扣分数据
        event1 = ["", "", ""]
        score1 = [0, 0, 0]
        for result in query_list2:
            classroom = result.get('classroom')
            event1[classroom] = event[classroom] + result.get('location') + result.get('event') + '(' + result.get('date')[11:-3] + ')'
            score1[classroom] = score[classroom] + result.get('score')
            print(event1[classroom])
            # result.set('isDownloaded',2)
            # result.save();
        # 写入表格
        sheet.write_merge(22, 22, 0, 0, '初三（1）', style2)
        sheet.write_merge(22, 22, 1, 1, str(fact1[1]) + ' / ' + str(ought1[1]), style2)
        situationString = ''
        if absent1[1] != 0:
            sheet.write_merge(22, 22, 2, 2, '-' + str(absent1[1]), style2)
            situationString += '缺席' +str(absent1[1])+'人 '
        else:
            sheet.write_merge(22, 22, 2, 2, '', style2)
        if score1[1] != 0:
            sheet.write_merge(22, 22, 3, 3, '-' + str(score1[1]), style2)
        else:
            sheet.write_merge(22, 22, 3, 3, '', style2)
        if leave1[1] != 0:
            situationString += '请假' + str(leave1[1]) + '人 '
        if temporary1[1] != 0:
            situationString += '临'+patString[1::1]+str(temporary[j])+'人 '
        if event1[1] != 0:
            situationString += event1[1]
        if situationString != '':
            sheet.write_merge(22, 22, 4, 4, situationString, style2)
        else:
            sheet.write_merge(22, 22, 4, 4, '', style2)

        situationString = ''
        sheet.write_merge(23, 23, 0, 0, '初三（2）', style2)
        sheet.write_merge(23, 23, 1, 1, str(fact1[2]) + ' / ' + str(ought1[2]), style2)
        if absent1[2] != 0:
            sheet.write_merge(23, 23, 2, 2, '-' + str(absent1[2]), style2)
            situationString += '缺席' +str(absent1[2])+'人 '
        else:
            sheet.write_merge(23, 23, 2, 2, '', style2)
        if score1[2] != 0:
            sheet.write_merge(23, 23, 3, 3, '-' + str(score1[2]), style2)
        else:
            sheet.write_merge(23, 23, 3, 3, '', style2)
        if leave1[2] != 0:
            situationString += '请假' + str(leave1[2]) + '人 '
        if temporary1[2] != 0:
            situationString += '临'+patString[1::1]+str(temporary[j])+'人 '
        if event1[2] != 0:
            situationString += event1[2]
        if situationString != '':
            sheet.write_merge(23, 23, 4, 4, situationString, style2)
        else:
            sheet.write_merge(23, 23, 4, 4, '', style2)

    path = GetDesktopPath()
    excel.save(path+"\\"+time[5:10]+gradeString+"课室"+patString+"情况登记表.xls")

# 获取桌面路径
import os
def GetDesktopPath():
    return os.path.join(os.path.expanduser("~"), 'Desktop')

# 获取需要抓取的年级
def GetGradeAndPattern():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("-intA", "--Grade", required=True, help="Grade")
    ap.add_argument("-intB", "--Pattern", required=True, help="Pattern")
    args = vars(ap.parse_args())
    grade = args["Grade"]
    pattern = args["Pattern"]
    return [grade,pattern]

#主程序
initLeanCloud()
[grade,pat] = GetGradeAndPattern()
if int(pat) == 0: patString = "午休"
else: patString = "晚修"
print("获取"+patString+"检查表...")

# 抓取应到实到数据
ClassData = leancloud.Object.extend('ClassData')
gradeQuery = ClassData.query
isDownloadedQuery = ClassData.query
gradeQuery.equal_to('grade',int(grade))
patternQuery = ClassData.query
patternQuery.equal_to('pattern',int(pat))
isDownloadedQuery.does_not_exist("isDailyDownloaded")
query = leancloud.Query.and_(gradeQuery,patternQuery,isDownloadedQuery)
query_list = query.find()

# 抓取扣分数据
SituationData = leancloud.Object.extend('SituationData')
situationQuery = SituationData.query
isDownloadedQuery = SituationData.query
situationQuery.equal_to('grade',int(grade))
patternQuery = SituationData.query
patternQuery.equal_to('pattern',int(pat))
isDownloadedQuery.equal_to("isDailyDownloaded",False)
query2 = leancloud.Query.and_(situationQuery,patternQuery,isDownloadedQuery)
query_list2 = isDownloadedQuery.find()
print (query_list2)

# 处理数据·处理应到实到数据
ought = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
fact = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
leave = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
temporary = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
absent = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
isFirstResult = True
time = ''
for result in query_list:
    pattern = result.get('pattern')
    checker = result.get('checker')
    dayOfWeek1 = result.get('dayOfWeek')
    if dayOfWeek1 == 1:
        dayOfWeek = "日"
    elif dayOfWeek1 == 2:
        dayOfWeek = "一"
    elif dayOfWeek1 == 3:
        dayOfWeek = "二"
    elif dayOfWeek1 == 4:
        dayOfWeek = "三"
    elif dayOfWeek1 == 5:
        dayOfWeek = "四"
    elif dayOfWeek1 == 6:
        dayOfWeek = "五"
    elif dayOfWeek1 == 7:
        dayOfWeek = "六"

    checkTime = str(result.get('date'))
    if (checkTime[0:10] != time[0:10]) and (not isFirstResult):
        continue
    classroom = result.get('classroom')
    ought[classroom] = result.get('ought')
    fact[classroom] = result.get('fact')
    leave[classroom] = result.get('leave')
    temporary[classroom] = result.get('temporary')
    absent[classroom] = result.get('absent')
    print(grade,classroom,ought[classroom], fact[classroom], temporary[classroom], absent[classroom])
    time = checkTime
    isFirstResult = False
    # result.set('isDailyDownloaded',True)
    # result.save();

# 处理数据·处理扣分数据
event = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
score = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
for result in query_list2:
    classroom = result.get('classroom')
    event[classroom] = event[classroom] + result.get('location') + result.get('event') + '(' + result.get('date')[11:-3] + ')'
    score[classroom] = score[classroom] + result.get('score')
    if (checkTime[0:10] != time[0:10]): continue
    print(event[classroom])
    # result.set('isDailyDownloaded',True)
    # result.save();

if isFirstResult == False :
    MakeExcel(int(grade))
    import tkinter.messagebox
    tkinter.messagebox.showinfo("智慧纪检", "成功！请到桌面查看")
else :
    import tkinter.messagebox
    tkinter.messagebox.showinfo("智慧纪检", "没有待收集的检查表")