#! /usr/bin/env python
# -*- coding: gbk -*-

import pandas as pd
import main
from pdb import set_trace as st
import os
import pyttsx3

files = {"高中单词": "high_ran.xlsx",
         '21天list': "21_ran.xls"}
unit = 'Unit'
english = '单词'

if __name__ == "__main__":
    while 1:
        student = input('请输入学生姓名：')  # '2':#
        if student not in main.names or student == '张瑞斌':
            print('输入错误，请重新输入。')
            continue
        else:
            break
    if student == '尹嘉禾':
        v_down = 20
    elif student == '杨瀚森':
        v_down = 90

    f_name, today = main.today(student)
    xls = pd.ExcelFile(files[f_name])
    s = xls.parse('Sheet1')
    units = s.loc[:, unit]
    w_dict = {}
    for each in today:
        w_dict.update({each: list(s[units == int(each)][english])})

    engine = pyttsx3.init()
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - v_down)
    voices = engine.getProperty('voices')

    for wl_name, words in w_dict.items():
        word_l = len(words)
        i = 0
        while i < word_l:
            word = words[i]


            def this():
                engine.setProperty('voice', voices[1].id)
                engine.say(word)
                engine.setProperty('voice', voices[2].id)
                engine.say(word)

                engine.runAndWait()


            while 1:
                os.system('cls')
                print('list', wl_name, '的第', i + 1, '个单词，共', word_l, '个。\n回车：下一词； 1+回车：上一词；其他键+回车：重新播放；')
                this()
                type_in = input()
                if type_in == '':
                    break
                elif type_in == '1':
                    i -= 2
                    break
                # sys.exit()
                else:
                    continue  # print('输入错误，请重新输入')

            i += 1
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