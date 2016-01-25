# encoding:utf-8

import docx2txt
import os
import pandas as pd
import re
import string

def printPath(level, path):
    global allFileNum
    '''
    打印一个目录下的所有文件夹和文件
    '''
    # 所有文件夹，第一个字段是次目录的级别
    dirList = []
    # 所有文件
    fileList = []
    # 返回一个列表，其中包含在目录条目的名称(google翻译)
    files = os.listdir(path)
    # 先添加目录级别
    dirList.append(str(level))
    for f in files:
        if(os.path.isdir(path + '/' + f)):
            # 排除隐藏文件夹。因为隐藏文件夹过多
            if(f[0] == '.'):
                pass
            else:
                # 添加非隐藏文件夹
                dirList.append(f)
        if(os.path.isfile(path + '/' + f)):
            # 添加文件
            fileList.append(f)
    # 当一个标志使用，文件夹列表第一个级别不打印
    i_dl = 0
    for dl in dirList:
        if(i_dl == 0):
            i_dl = i_dl + 1
        else:
            # 打印至控制台，不是第一个的目录
            print '-' * (int(dirList[0])), dl
            # 打印目录下的所有文件夹和文件，目录级别+1
            printPath((int(dirList[0]) + 1), path + '/' + dl)
    return fileList

file_list = printPath(1, os.getcwd())

def cut2sentence(text):
    all_sentence = []
    for key in string.split(text, '\n'):
        if key != u'':
            all_sentence.append(key)
    return all_sentence


result = []

tc_pattern = [re.compile(u' [\u4e00-\u9fa5]+'), re.compile(r' \w+'), re.compile(u'：[\u4e00-\u9fa5]+'), re.compile(u':[\u4e00-\u9fa5]+')]
chinese_pattern = re.compile(u'[\u4e00-\u9fa5]{5}')

for filename in file_list:
    if 'docx' in filename and '~$' not in filename:
        print filename
        text = docx2txt.process(filename)
        all_paragraph = cut2sentence(text)

        # english subject
        english_subject = all_paragraph[0]

        # chinese subject
        chinese_subject = all_paragraph[1]

        # name
        translator_name = ''
        checker_name = ''

        for i in xrange(1, 10):
            paragraph = all_paragraph[i]
            first_word = paragraph[0]

            if first_word == u'翻' or first_word  == u'译':
                c = ''
                for k in xrange(0, len(tc_pattern)):
                    pattern = tc_pattern[k]
                    a = pattern.search(paragraph)
                    if a:
                        c = a.group()
                        break

                for j in xrange(1, len(c)):
                    if c[j] != ' ':
                        translator_name += c[j]

            elif first_word == u'审' or first_word == u'校':
                c = ''
                for k in xrange(0, len(tc_pattern)):
                    pattern = tc_pattern[k]
                    a = pattern.search(paragraph)
                    if a:
                        c = a.group()
                        break

                for j in xrange(1, len(c)):
                    if c[j] != ' ':
                        checker_name += c[j]

        # the number of words:

        num_chinese = 0
        num_english = 0
        for paragraph in all_paragraph:
            if chinese_pattern.search(paragraph):
                #print paragraph
                for words in paragraph:
                    if re.compile(u'[\u4e00-\u9fa5]').search(words) :
                        num_chinese += 1

            else:
                '''for word in paragraph:
                    if word in string.punctuation:
                        num_english += 1'''
                num_english += string.split(paragraph).__len__()
        if translator_name == u'M':
            translator_name = u'Meatle'
        if checker_name == u'M':
            checker_name = u'Meatle'
        result.append([english_subject, chinese_subject, translator_name, checker_name, num_english, num_chinese])

len = result.__len__()
result = pd.DataFrame(result)
result.set_axis(1, [u'英文题目', u'中文题目', u'翻译', u'审校', u'英文字数', u'中文字数'])
result.set_axis(0, range(1, len+1))
result.to_csv('统计结果.csv', encoding='utf-8')






