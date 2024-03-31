import sys


# 或者在判断的同时，获取节日名
import chinese_calendar as calendar
from chinese_calendar import is_workday

import xlwt
import xlrd      # pip install xlrd==1.2.0
import datetime
import os


class Employee:
    def __init__(self, name, monthDates):
        self.name = name

        self.chidaoZaotuiMinutes = 0
        self.quekaCount = 0
        self.queqinCount = 0

        self.detailChidaoZaotui = {}
        for date in monthDates:
            self.detailChidaoZaotui[date] = {}
            self.detailChidaoZaotui[date]['normal'] = []
            self.detailChidaoZaotui[date]['serious'] = []

        self.quekaoDicts = {}
        for date in monthDates:
            self.quekaoDicts[date] = {}
            self.quekaoDicts[date]['am'] = 0
            self.quekaoDicts[date]['pm'] = 0
            self.quekaoDicts[date]['twoDetail'] = ''
            self.quekaoDicts[date]['oneDetail'] = ''


def process_kaoqin(outputPath, inputFileName):

    now = datetime.datetime.now()
    now_str = now.strftime("%Y%m%d_%H%M%S")
    write_filename = 'wtt_' + now_str + '.xls'
    write_workbook_name = outputPath + "/" + write_filename



    # 创建一个workbook设置编码
    write_workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    write_worksheet = write_workbook.add_sheet('汇总sheet')


    # 打开Excel文件
    read_workbook = xlrd.open_workbook(inputFileName)

    # 获取Sheet对象
    read_sheet = read_workbook.sheet_by_name('原始考勤记录报表')

    # 获取行数和列数
    rows = read_sheet.nrows
    cols = read_sheet.ncols


    write_worksheet.write(0, 0, '姓名')
    write_worksheet.write(0, 1, '实际出勤天数')
    write_worksheet.write(0, 2, '实际缺勤天数')
    write_worksheet.write(0, 3, '迟到早退（分钟）')
    write_worksheet.write(0, 4, '普通迟到早退')
    write_worksheet.write(0, 5, '严重迟到早退')

    write_worksheet.write(0, 8, '漏卡次数')
    write_worksheet.write(0, 9, '详细漏卡半天')
    write_worksheet.write(0, 10, '详细漏卡整天')


    time_str1 = "2024-04-01 08:00:00"
    time_str2 = "2024-04-01 17:00:00"
    time_str3 = "2024-04-01 08:30:00"
    time_str4 = "2024-04-01 16:30:00"
     
    time1 = datetime.time(8, 0, 0)
    time2 = datetime.time(8, 30, 0)
    time3 = datetime.time(17, 0, 0)
    time4 = datetime.time(16, 30, 0)
    time5 = datetime.time(12, 30, 0)
    # 将字符串转换为datetime对象 
    # time1 = datetime.strptime(time_str1, "%Y-%m-%d %H:%M:%S").time()
    # time2 = datetime.strptime(time_str2, "%Y-%m-%d %H:%M:%S").time()
    # time3 = datetime.strptime(time_str3, "%Y-%m-%d %H:%M:%S").time()
    # time4 = datetime.strptime(time_str4, "%Y-%m-%d %H:%M:%S").time()
    today = datetime.datetime.today().date()
    datetime1 = datetime.datetime.combine(today, time1)
    datetime2 = datetime.datetime.combine(today, time2)
    datetime3 = datetime.datetime.combine(today, time3)
    datetime4 = datetime.datetime.combine(today, time4)
    datetime5 = datetime.datetime.combine(today, time5)
    # print(time1)

    monthDates = []
    workdays = []
    specialDays = {}

    # 第一个key是name，第二个key是day（无time），内容是time
    dicts = {}

    
    # 输出每行的数据
    for i in range(1,rows):
        row_data = read_sheet.row_values(i)
        # print(row_data)
        # 这里的日期是时间戳显示（float类型），需要转换
        name = row_data[1]
        allDate = xlrd.xldate_as_datetime(row_data[3], 0)

        if len(workdays) == 0:
            for i in range(1, 32):
                try:
                    thisdate = datetime.date(allDate.year, allDate.month, i)
                    thisdateStr = thisdate.strftime("%Y-%m-%d")
                    monthDates.append(thisdate)
                    # write_worksheet.write(4, 9 + thisdate.day, thisdateStr)
                except(ValueError):
                    break
                on_holiday, holiday_name = calendar.get_holiday_detail(thisdate)
                # if is_workday(thisdate):
                # 是周一到周六，并且不是节假日（去除周六周日）
                if thisdate.weekday() < 6 and (not on_holiday or (on_holiday and holiday_name == None)): # Monday == 0, Sunday == 6 
                    workdays.append(thisdate)
                elif on_holiday and holiday_name != None:
                    specialDays[thisdate] = holiday_name

            write_worksheet.write(1, 0, '本月是：{0}月，总共{1}天，工作日是{2}天, {3}'.format(allDate.month, len(monthDates), len(workdays), '; '.join([x.strftime("%Y-%m-%d") for x in workdays])))
            write_worksheet.write(2, 0, '特殊日期：{0}'.format('; '.join(['{0}: {1}'.format(x.strftime("%Y-%m-%d"), specialDays[x]) for x in specialDays.keys()])))


        # print(type(allDate))
        # print(allDate)
        if dicts.get(name) == None:
            dicts[name] = {}
        # print(allDate.day)
        # print(allDate.date())
        # print(allDate.time())
        if dicts.get(name).get(allDate.date()) == None:
            dicts[name][allDate.date()] = []
        dicts[name][allDate.date()].append(allDate.time())



    for index, (name, kaoqins) in enumerate(dicts.items()):
        employee = Employee(name, monthDates)
        

        write_worksheet.write(index + 5, 0, name)

        # 每天的时间，取一个最大的，取一个最小的，表示最早打卡时间和最晚打卡时间
        for dateStr in kaoqins.keys():
            if (dateStr not in workdays):
                continue
            times = kaoqins[dateStr]
            amTimes = [x for x in times if (datetime.datetime.combine(today, min(times)) - datetime5).total_seconds() < 0]
            pmTimes = [x for x in times if (datetime.datetime.combine(today, max(times)) - datetime5).total_seconds() >= 0]
            # print(name)
            # print(max(times))
            # print(min(times))

            if len(amTimes) > 0:
                employee.quekaoDicts[dateStr]['am'] = 2
                datetimeMin = datetime.datetime.combine(today, min(amTimes))
                diffTime = int(( datetimeMin - datetime1).total_seconds() / 60);
                if diffTime > 0:
                    if diffTime > 30:
                        employee.detailChidaoZaotui[dateStr]['serious'].append('{0} {1} 上午迟到 {2} 分钟'.format(dateStr, min(amTimes).strftime('%H:%M'), diffTime))
                    else:
                        employee.chidaoZaotuiMinutes += diffTime
                        employee.detailChidaoZaotui[dateStr]['normal'].append('{0} {1} 迟到 {2} 分钟'.format(dateStr, min(amTimes).strftime('%H:%M'), diffTime))

            if len(pmTimes) > 0:
                employee.quekaoDicts[dateStr]['pm'] = 2
                datetimeMax = datetime.datetime.combine(today, max(pmTimes))
                diffTime = int((datetime3 - datetimeMax).total_seconds() / 60);
                if diffTime > 0:
                    if diffTime > 30:
                        employee.detailChidaoZaotui[dateStr]['serious'].append('{0} {1} 下午早退 {2} 分钟'.format(dateStr, max(pmTimes).strftime('%H:%M'), diffTime))
                    else:
                        employee.chidaoZaotuiMinutes += diffTime
                        employee.detailChidaoZaotui[dateStr]['normal'].append('{0} {1} 早退 {2} 分钟'.format(dateStr, max(pmTimes).strftime('%H:%M'), diffTime))


        for i in workdays:
            quekaDict = employee.quekaoDicts[i]
            if quekaDict.get('am') == 0 and quekaDict.get('pm') == 0:
                quekaDict['twoDetail'] = ('{0}整天缺卡'.format(i))
                employee.queqinCount += 1
            elif quekaDict.get('am') == 0:
                employee.quekaCount += 1
                quekaDict['oneDetail'] = ('{0}上午缺卡'.format(i))
            elif quekaDict.get('pm') == 0:
                employee.quekaCount += 1
                quekaDict['oneDetail'] = ('{0}下午缺卡'.format(i))
            

        write_worksheet.write(index + 5, 1, len(workdays) - employee.queqinCount)

        if employee.queqinCount != 0:
            write_worksheet.write(index + 5, 2, employee.queqinCount)
        if employee.chidaoZaotuiMinutes != 0:
            write_worksheet.write(index + 5, 3, employee.chidaoZaotuiMinutes)
        if employee.quekaCount != 0:
            write_worksheet.write(index + 5, 8, employee.quekaCount)


        write_worksheet.write(index + 5, 4, '。\n '.join(['; '.join(employee.detailChidaoZaotui[x]['normal']) for x in employee.detailChidaoZaotui if len(employee.detailChidaoZaotui[x]['normal']) > 0]))
        write_worksheet.write(index + 5, 5, '。\n '.join(['; '.join(employee.detailChidaoZaotui[x]['serious']) for x in employee.detailChidaoZaotui if len(employee.detailChidaoZaotui[x]['serious']) > 0]))
        write_worksheet.write(index + 5, 9, '。\n '.join([employee.quekaoDicts[x]['oneDetail'] for x in employee.quekaoDicts.keys() if employee.quekaoDicts[x]['oneDetail'] != '']))
        write_worksheet.write(index + 5, 10, '。\n '.join([employee.quekaoDicts[x]['twoDetail'] for x in employee.quekaoDicts.keys() if employee.quekaoDicts[x]['twoDetail'] != '']))

        # 存储写入的
        write_workbook.save(write_workbook_name)





    # 存储写入的
    # write_workbook.save(write_workbook_name)

# # 自己电脑处理得注释掉，放到云端需要放开。
# process_kaoqin(sys.argv[1], sys.argv[2])

# 为了在自己电脑处理用的
if __name__ == '__main__':
    current_work_dir = os.path.abspath(os.path.dirname(__file__))           # 当前文件所在的目录，不能在命令行运行，会__file__ not defined
    process_kaoqin(current_work_dir, r'/Users/dinosaur/jcwang/IDEA/WTT_taima_kaoqin/baobiao1.xlsx')


