#!/sr/bin/env python
# -*- coding: gbk -*-

# inter=[0,1,2,4,7,15,30]
# for i in range(1, len(inter)):
#     day_list.append(sum(inter[:i]))
# print(day_list)

from datetime import *
import pyperclip
import pandas as pd
from pdb import set_trace as st
import numpy as np

interval = [1, 2, 4, 7, 15, 30]
review_day = [1, 3, 7, 14, 29, 59]
inter_l = len(interval)

f_dir = r"E:\py\tools\s_plan\\"
f_name = 't.xlsx'
dir = r"E:\py\tools\s_plan\t.xlsx"

init = '初次背诵日期'
content = '内容'
sn = '已复习次数'
last = '上次复习日期'
next = '下次复习日期'
dl = 'deadline'
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

    def write_n1(col, value):
        if check_n(col):
            s.loc[pd.isnull(s[col]), col] = value

    # index
    # if pd.isnull(s.index).any():
    #     pd.isnull
    # 编号
    write_n1(sn, 0)
    # 上次
    write_n1(last, now)
    # 下次时间
    if check_n(next):
        nan = pd.isnull(s[next])
        s.loc[nan, next] = future(s.loc[nan, sn]) + s.loc[nan, last]
    # deadline时间
    if check_n(dl):
        nan = pd.isnull(s[dl])
        s.loc[nan, dl] = future(s.loc[nan, sn] + 1) + s.loc[nan, next]


xls = pd.ExcelFile(dir)
writer = pd.ExcelWriter(dir)

for name in xls.sheet_names:
    s = xls.parse(name)
    initial()

    nexts = s.loc[:, next]
    dls = s.loc[:, dl]
    today = s[(nexts <= (now+pd.Timedelta(days=1))) & (dls >= now)]
    tomorrow = s[nexts == (now + pd.Timedelta(days=1))]

    # tasks = {'今天': (today.index, today[content]),
    #          '明天': (tomorrow.index, tomorrow[content])}
    # a= tasks['明天']
    # for i,b in enumerate(a[1]):
    #     print(i,b)
    # print(list(tomorrow[content]))
    today=pd.DataFrame({'序号': pd.Series(np.arange(1, 1 + len(today))),
                        '任务': (list(today[content])),
                        'index':pd.Series(today.index)},
                       columns=['序号', '任务', 'index'])

    print('今天是', date.today())
    print('学生：', name)
    def tasks(df):
        l= len(df)
        if l == 0:
            print('无复习任务。')
            return 0
        else:
            print('的复习任务为：')
            print(df.to_string(index=False))
            return l

    print("今天",end='')
    today_no= tasks(today.loc[:,['序号', '任务']])
    print('明天',end='')
    tasks(tomorrow[content])
    tom_head=''
    st()



    for k, v in tasks.items():
        copy = ''

        for i, item in enumerate(v[1]):
            print(i, item,)
            copy += (item + ' ')
        pyperclip.copy(copy)

    warn = s[(nexts < now) & (dls >= now)][dl]
    if len(warn) != 0:
        print('其中需要注意的任务为')
        print(warn)

    while today_no:
        type_in=input('请输入完成的任务号，用空格间隔:')
        if type_in == '':
            break

        else:
            type_in = [int(i) for i in type_in.split(' ')]
        try:
            for each in type_in:
                if int(each) not in today[content].index:
                    raise IndexError('输入错误，请重新输入')
        except IndexError as e:
            print(e)
        else:
            s.loc[type_in, sn] += 1
            s.loc[type_in, erase] = np.nan
            initial()
            break

    s.to_excel(writer,sheet_name=name)
writer.save()
# if len(make_up)!=0:
#     print('补作业！')
#     for each in make_up:
#         print(each[content],'deadline为',each[dl])

# wb.save(f_dir + f_name)

# os.startfile(f_dir + f_name)