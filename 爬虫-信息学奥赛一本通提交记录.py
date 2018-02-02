import requests
from lxml import etree
from matplotlib import pyplot as plt
import xlwt
import xlrd
from datetime import datetime, timedelta
import os
import sys


def GetList(username, name, filename, save_file):
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36'}
    url = 'http://ybt.ssoier.cn:8088/status.php?start=%d&showname=%s&showpid=&showres=%s&showlang='
    ac_cnt = 0
    ce_cnt = 0
    all_cnt = 0
    markp = 0
    time_list = []
    pro_list = []
    ach_list = []
    pro_ac_num = {}

    pagenum = 0
    # print('开始抓取数据')
    while True:

        req = requests.get(url % (pagenum * 20, username, ''), header).text.encode("ISO-8859-1", "ignore").decode("utf8", "ignore")
        pagenum += 1
        # print(pagenum, '页')
        html = etree.HTML(req)
        res = html.xpath('//center/table/tr')
        if not res[3::]:
            break
        for i in res[3::]:
            all_cnt += 1
            thisce = 0

            achievement_list = i.xpath('./script')
            if achievement_list:
                achievement = achievement_list[0].text
                if 'Accepted' in achievement:
                    # print("AC!")
                    ach_list.append("AC!")
                    ac_cnt += 1
                else:
                    # print("WA!")
                    ach_list.append("WA!")
            else:
                # print("CE!")
                ach_list.append("CE!")
                ce_cnt += 1
                thisce = 1

            time_record = i.xpath('./td')[5 + thisce].text
            if time_record:
                # print(j)
                time_list.append(time_record)

            pro_num_record = i.xpath('./td/a')[1].text
            # print(pro_num_record)
            pro_list.append(pro_num_record)
            if ach_list[markp] is 'AC!':
                if pro_num_record not in pro_ac_num:
                    pro_ac_num[pro_num_record] = 0
                pro_ac_num[pro_num_record] += 1
            markp += 1

    duty_list = []
    for key in pro_ac_num:
        if pro_ac_num[key] >= 5:
            duty_list.append(int(key))

    ffinal = open("./all.txt","a")
    print("成功抓取学生%s的数据\n------------------------" % name)
    print("提交次数\t\tAC次数\t\tWA次数\t\tCE次数\t\tAC题目数")
    print('%-8d\t%-8d\t%-8d\t%-8d\t%-8d' % (all_cnt, ac_cnt, all_cnt-ac_cnt-ce_cnt, ce_cnt, len(pro_ac_num)))
    print('------------------------\n正确率%f%%\n' % (100*ac_cnt/all_cnt))

    print('搜索到脏数据（AC5次及以上）%d条：' % len(duty_list))
    if len(duty_list):
        for i in duty_list:
            print('\t第%d题，提交成功%d次' % (i, pro_ac_num[str(i)]))

    save_sheet1 = save_file.add_sheet(u'sheet1')
    save_sheet1.write(0, 0, '%s提交记录' % name)
    save_sheet1.write(1, 0, '题目号')
    save_sheet1.write(1, 1, '提交时间')
    save_sheet1.write(1, 2, '判题结果')
    for i in range(len(pro_list)):
        save_sheet1.write(i + 2, 0, pro_list[i])
        save_sheet1.write(i + 2, 1, time_list[i])
        save_sheet1.write(i + 2, 2, ach_list[i])

    subtract_time_list = []
    duty_list2 = []
    contiudirty_num = 0
    for i in range(1, len(time_list)):
        if int(pro_list[i]) <= 1020:
            if contiudirty_num >= 5:
                duty_list2.extend(subtract_time_list)
                contiudirty_num = 0
                subtract_time_list = []
            continue
        if ach_list[i] == 'AC!' and ach_list[i-1] == 'AC!':
            time_a = datetime.strptime(time_list[i-1], '%Y-%m-%d %H:%M:%S')
            time_b = datetime.strptime(time_list[i], '%Y-%m-%d %H:%M:%S')
            if time_a - time_b < timedelta(days=0, minutes=2):
                contiudirty_num += 1
                subtract_time_list.append(i)
                continue
            elif contiudirty_num >= 5:
                duty_list2.extend(subtract_time_list)
            contiudirty_num = 0
            subtract_time_list = []
    print('搜索到脏数据（连续5次以上，在120秒内连续AC）%d条：' % len(duty_list2))
    if len(duty_list2):
        for i in duty_list2:
            print('\t时间：%s\t题号：%s' % (time_list[i], pro_list[i]))

    ffinal = open("./all.txt", "a")
    ffinal.write('%d %d %d ' % (all_cnt, len(pro_ac_num), len(duty_list)+len(duty_list2)))
    ffinal.close()

    fig1 = plt.figure(figsize=(10, 10))
    plt.axes(aspect=1)
    plt.pie([ac_cnt, all_cnt-ac_cnt-ce_cnt, ce_cnt], explode=[0.1, 0, 0], labels=['AC', 'WA', 'CE'], autopct='%f %%', colors=['yellowgreen','red', 'lightblue'], shadow=True, startangle=90)
    plt.title('%s Correct Rate' % username)
    plt.legend()
    plt.savefig('./%s//%s提交记录饼图.jpg' % (filename, name), dpi=140)
    # plt.show()
    # print(pro_list)
    # print(time_list)
    # print(pro_ac_num)
    return time_list, pro_list, ach_list


def AnlysTime(filename, name, time_list, ach_list, save_file):
    datestr_list = []
    num_list = []
    num_list2 = []

    work_info = {}
    ac_info = {}
    markp = 0
    starttime = datetime(2017, 9, 7)
    endtime = datetime.now()
    step = timedelta(days=1)

    days = (endtime-starttime).days
    for strtime in time_list:
        dttime = datetime.strptime(strtime, '%Y-%m-%d %H:%M:%S')
        month = dttime.strftime('%m')
        day = dttime.strftime('%d')
        timeinfo = ''.join(month + '-' + day)
        # print(timeinfo)
        if timeinfo not in work_info:
            work_info[timeinfo] = 0
        work_info[timeinfo] += 1
        if ach_list[markp] == 'AC!':
            if timeinfo not in ac_info:
                ac_info[timeinfo] = 0
            ac_info[timeinfo] += 1
        markp += 1

    dates = [starttime + step * i for i in range(days+1)]
    for i in dates:
        month = i.strftime('%m')
        day = i.strftime('%d')
        datestr = ''.join(month+'-'+day)
        if datestr in work_info:
            num_list.append(work_info[datestr])
        else:
            num_list.append(0)
        if datestr in ac_info:
            num_list2.append(ac_info[datestr])
        else:
            num_list2.append(0)
        datestr_list.append(datestr)

    fig2 = plt.figure(figsize=(20, 15))
    pagenum = 310
    stepnum = 0
    save_sheet2 = save_file.add_sheet(u'sheet2')
    for step in [1, 3, 7]:
        stepnum += 1
        new_num_list = []
        new_num_list2 = []
        tmpsum = 0
        for i in range(len(num_list)):
            tmpsum += num_list[i]
            if i == len(num_list)-1:
                new_num_list.append(tmpsum)
                tmpsum = 0
            elif i % step == step-1:
                new_num_list.append(tmpsum)
                tmpsum = 0
        for i in range(len(num_list2)):
            tmpsum += num_list2[i]
            if i == len(num_list2)-1:
                new_num_list2.append(tmpsum)
            elif i % step == step-1:
                new_num_list2.append(tmpsum)
                tmpsum = 0

        lazy_week = 0
        for i in new_num_list:
            if i == 0:
                lazy_week += 1

        ffinal = open("./all.txt", "a")
        ffinal.write("%d\n" % lazy_week)
        ffinal.close()

        save_sheet2.write(0, 4 * (stepnum - 1), '%s每%d天提交情况' % (name, step))
        save_sheet2.write(1, 4 * (stepnum - 1), '日期')
        save_sheet2.write(1, 4 * (stepnum - 1) + 1, '提交次数')
        save_sheet2.write(1, 4 * (stepnum - 1) + 2, 'AC次数')
        dateFormat = xlwt.XFStyle()
        dateFormat.num_format_str = 'yyyy-mm-dd'
        for i in range(len(new_num_list)):
            save_sheet2.write(i + 2, 4 * (stepnum - 1), dates[i * step], dateFormat)
            save_sheet2.write(i + 2, 4 * (stepnum - 1) + 1, new_num_list[i])
            save_sheet2.write(i + 2, 4 * (stepnum - 1) + 2, new_num_list2[i])

        pagenum += 1
        fig2.add_subplot(pagenum)
        plt.plot(dates[::step], new_num_list, 'o', color='blue')
        plt.plot(dates[::step], new_num_list, '-', label='Submit', color='blue')
        plt.plot(dates[::step], new_num_list2, 'o', color='red')
        plt.plot(dates[::step], new_num_list2, '-', label='AC', color='red')
        plt.xlabel('DATE')
        plt.ylabel('Number')
        plt.title('%d days Records of submission' % step)
        plt.grid(True)
        plt.xlim(starttime, endtime)
        plt.legend()
    plt.savefig('./%s//%s提交记录折线.jpg' % (filename, name), dpi=140)
    # plt.show()
    return dates, num_list, num_list2


def MainWord(username, name, stdnum):
    filename = stdnum + name
    if not os.path.isdir('./%s' % filename):
        os.mkdir('./%s' % filename)

    save_file = xlwt.Workbook()
    printouttmp = sys.stdout
    outf = open('./%s//%s数据统计报告.txt' % (filename, name), 'w')
    sys.stdout = outf

    print('学生：%s   学号：%s   用户名：%s' % (name, stdnum, username))

    time_list, pro_list, ach_list = GetList(username, name, filename, save_file)
    AnlysTime(filename, name, time_list, ach_list, save_file)

    save_file.save('./%s//%s提交记录.xls' % (filename, name))
    sys.stdout = printouttmp
    plt.close('all')
    outf.close()
    print('学生%s数据统计已完成，数据分析结果、统计图、统计表已保存\n--------------------------------------------------------' % name)


if __name__ == '__main__':
    read_file = xlrd.open_workbook('./15级计师算法与信息学奥赛选课名单.xlsx')
    sh1 = read_file.sheet_by_index(0)
    for stdn in range(1, sh1.nrows):
        name = sh1.cell_value(stdn, 0)
        stdnum = sh1.cell_value(stdn, 1)
        username = sh1.cell_value(stdn, 2).strip()
        MainWord(name=name, username=username, stdnum=stdnum)

    print('%d个用户全部结束' % (int(sh1.nrows)-1))
