#! /usr/bin/env python
# -*- coding: gbk -*-

import pandas as pd
import main
from pdb import set_trace as st
import os
import pyttsx3
files= {"high":"high school.xlsx",
    '21':"21_ran.xls"}
unit='Unit'
english='单词'

if __name__=="__main__":
    if input('选择：1.杨瀚森 2.尹嘉禾')=='2':
        xls = pd.ExcelFile(files['21'])
    else:
        xls = pd.ExcelFile(files['high'])
    s=xls.parse('Sheet1')
    units=s.loc[:, unit]
    words=list(s[units==1][english])
    # print(words)

    engine = pyttsx3.init()
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - 80)
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)

    for i, word in enumerate(words):
        def this():
            engine.say(word)
            engine.runAndWait()
        while 1:
            os.system('cls')
            print('第', i + 1, '个单词，共', len(words), '个。\n回车：下一词  其他键+回车：重新播放；')# 1+回车：退出
            this()
            type_in=input()
            if type_in=='':
                break
            #elif type_in=='1':
            #    os.exit()
            else:
                continue#print('输入错误，请重新输入')
        # if type_in==
    # while 1:
    #     tasks=main.today('杨瀚森')#input('请输入你的全名：'))
    #     print(tasks)
    #     if tasks==None:
    #         print('恭喜你，今天没有任务。')
    #         exit()
    #     elif tasks=='error':
    #         os.system('cls')
    #         print('输入错误，请重新输入')
    #         continue
    #     else:
    #         os.system('cls')
    #         print('现在开始听写...')
    #         break
    # st()
    # for k,v in files.items():
    #     xls = pd.ExcelFile(v)
    #     print(k)