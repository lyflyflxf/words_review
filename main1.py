#!/sr/bin/env python
# -*- coding: gbk -*-

# inter=[0,1,2,4,7,15,30]
# for i in range(1, len(inter)):
#     day_list.append(sum(inter[:i]))
# print(day_list)

from datetime import *
import pyperclip
import pandas as pd
# from pdb import set_trace as st
import numpy as np

# review_day = [1, 3, 7, 14, 29, 59]
interval = [1, 2, 4, 7, 15, 30]
inter_l = len(interval)

# f_dir = r"E:\py\tools\s_plan\\"
# f_name = 't.xlsx'
dir = r"E:\py\tools\s_plan\t.xlsx"

init = '初次背诵日期'
content = '内容'
count = '已复习次数'
last = '上次复习日期'
next = '下次复习日期'
dl = 'deadline'
sn = '序号'
erase = [last, next, dl]

now = pd.Timestamp(date.today())

def future(l):
    out = []
    for sn in l:
        sn = int(sn)
        if sn < inter_l:
            out.append(interval[sn])
        else:
            out.append(60)
    return [pd.Timedelta(days=each) for each in out]


def initial():
    def check_n(col):
        return pd.isnull(s[col]).any()

    def write_n(col, value):
        if check_n(col):
            s.loc[pd.isnull(s[col]), col] = value

    # 编号
    write_n(count, 0)
    # 上次
    write_n(last, now)
    # 下次时间
    if check_n(next):
        n = pd.isnull(s[next])
        m= s.loc[n,:]
        s.loc[n, next] = future(m[count]) + m[last]
    # deadline时间
    if check_n(dl):
        n = pd.isnull(s[dl])
        m = s.loc[n, :]
        s.loc[n, dl] = future(m[count] + 1) + m[next]


xls = pd.ExcelFile(dir)
writer = pd.ExcelWriter(dir)

for name in xls.sheet_names:
    s = xls.parse(name)
    initial()

    nexts = s.loc[:, next]
    dls = s.loc[:, dl]
    today = s[(nexts <= (now + pd.Timedelta(days=1))) & (dls >= now)]
    today = pd.DataFrame({sn: pd.Series(np.arange(1, 1 + len(today)),
                                        index=today.index)}).join(today)

    tomorrow = s[nexts == (now + pd.Timedelta(days=1))]

    print('今天是', date.today(),'\n', '学生：', name)

    def tasks(day, df):
        l = len(df)
        if l == 0:
            print(day + '无复习任务。')
            return 0
        else:
            print(day + '的复习任务为：')
            return l


    tasks('明天', tomorrow)
    tom_heads = ['高中单词']
    l = list(tomorrow[content])
    copy = ''

    for head in tom_heads:
        out =''
        h_len = len(head)
        for item in l:
            if item.startswith(head):
                out += item[h_len:]+ '、'
        if len(out) != 0:
            copy += head + out
    copy= copy[:-1]
    print(copy)
    pyperclip.copy(copy)


    today_no = tasks("今天", today)
    print(today[[sn, content]].to_string(index=False))

    warn = s[(nexts < now) & (dls >= now)][dl]
    if len(warn) != 0:
        print('其中需要注意的任务为')
        print(warn)

    while today_no:
        type_in = input('请输入完成的任务号，用空格间隔:')
        if type_in == '':
            break

        type_in = [int(i) for i in type_in.split(' ')]
        error=False
        for each in type_in:
            if not today[sn].any() == each:
                print('输入错误，请重新输入')
                error=True
                break
        if not error:
            type_in = today[today[sn] == type_in].index
            s.loc[type_in, count] += 1
            s.loc[type_in, erase] = np.nan
            initial()
            break
    s.to_excel(writer, sheet_name=name)
writer.save()
# if len(make_up)!=0:
#     print('补作业！')
#     for each in make_up:
#         print(each[content],'deadline为',each[dl])

# wb.save(f_dir + f_name)

# os.startfile(f_dir + f_name)