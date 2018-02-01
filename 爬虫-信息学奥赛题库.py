import requests
from lxml import etree
import xlwt


def bashuoj(save_file):
    save_sheet = save_file.add_sheet(u'sheet1')
    urls = 'http://ybt.ssoier.cn:8088/problem_show.php?pid=%d'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36'}

    for pid in range(1000, 1418):

        req = requests.get(url=urls % pid, headers=headers).text.encode('ISO-8859-1', 'ignore').decode('utf-8', 'ignore')
        print('开始抓取%d题' % pid)
        reqhtml = etree.HTML(req)

        res = reqhtml.xpath('//center/table/td/center/h3')
        tmmc = res[0].text
        tmms = ''
        sr = ''
        sc = ''
        sryl = ''
        scyl = ''
        ts = ''
        tyb = 0
        res = reqhtml.xpath('//p|//h3|//font/div/pre|//font/p|//p/img/@src')
        for i in res:
            if not hasattr(i, 'text'):
                with open('./%s' % str(i), 'wb') as f:
                    f.write(requests.get('http://ybt.ssoier.cn:8088/%s' % str(i)).content)
                if tyb == 1:
                    tmms += str(i)
                elif tyb == 2:
                    sr += str(i)
                elif tyb == 3:
                    sc += str(i)
                elif tyb == 4:
                    sryl += str(i)
                elif tyb == 5:
                    scyl += str(i)
                elif tyb == 6:
                    ts += str(i)
                continue

            if(i.text):
                if i.text == '【题目描述】':
                    tyb = 1
                elif i.text == '【输入】':
                    tyb = 2
                elif i.text == '【输出】':
                    tyb = 3
                elif i.text == '【输入样例】':
                    tyb = 4
                elif i.text == '【输出样例】':
                    tyb = 5
                elif i.text == '【提示】':
                    tyb = 6
                elif i.text == '【来源】':
                    break

                elif tyb == 1:
                    tmms += i.text
                elif tyb == 2:
                    sr += i.text
                elif tyb == 3:
                    sc += i.text
                elif tyb == 4:
                    sryl += i.text
                elif tyb == 5:
                    scyl += i.text
                elif tyb == 6:
                    ts += i.text

        save_sheet.write(pid - 999, 0, pid)
        # print('【题目名称】')
        # print(tmmc)
        save_sheet.write(pid-999, 1, tmmc)
        # print('【题目描述】')
        # print(tmms)
        save_sheet.write(pid - 999, 2, tmms)
        # print('【输入】')
        # print(sr)
        save_sheet.write(pid - 999, 3, sr)
        # print('【输出】')
        # print(sc)
        save_sheet.write(pid - 999, 4, sc)
        # print('【输入样例】')
        # print(sryl)
        save_sheet.write(pid - 999, 5, sryl)
        # print('【输出样例】')
        # print(scyl)
        save_sheet.write(pid - 999, 6, scyl)
        # print('【提示】')
        # print(ts)
        save_sheet.write(pid - 999, 7, ts)
        print(pid, '完毕')


if __name__ == '__main__':
    save_file = xlwt.Workbook()
    bashuoj(save_file)
    save_file.save('./信息学奥赛题库.xls')
