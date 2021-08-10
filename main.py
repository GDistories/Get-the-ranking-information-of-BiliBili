import requests
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils import copy
import time


# Count String Length
def maxLen(list):
    maxlen = 0
    for i in list:
        if maxlen < len(i):
            maxlen = len(i)
    return maxlen


# param
title_name = []
up_name = []
play_number = []
barrage_number = []
VideoList = []
maxtitlelen = 0
maxupnamelen = 0
maxplaynumlen = 0
maxbarragenumlen = 0
counter = 0
url = "https://www.bilibili.com/v/popular/rank/all"

# Get Html Page
html = requests.get(url)
soup = BeautifulSoup(html.text, 'lxml')

# Get Video Name
for i in soup.find_all(class_='title', target="_blank"):
    title_name.append(i.text)

# Get UP Name, Play Number, Barrage Number
for i in soup.find_all(class_='data-box'):
    VideoList_html = i.text.replace(' ', '')
    VideoList_html = VideoList_html.replace('\n', '')
    VideoList.append(VideoList_html)

# Get UP Name
up_name = VideoList[2::3]

# Get Play Number
play_number = (VideoList[::3])

# Get Barrage Number
barrage_number = (VideoList[1::3])

# Count length
maxtitlelen = maxLen(title_name)
maxupnamelen = maxLen(up_name)
maxplaynumlen = maxLen(play_number)
maxbarragenumlen = maxLen(barrage_number)

# Create Workbook
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet(str(time.strftime('%Y-%m-%d-%H.%M', time.localtime(time.time()))))

# Set Row Width
worksheet.col(0).width = 256 * 2 * 2
worksheet.col(1).width = 256 * 2 * maxtitlelen
worksheet.col(2).width = 256 * 2 * maxupnamelen
worksheet.col(3).width = 256 * 2 * maxplaynumlen
worksheet.col(4).width = 256 * 2 * maxbarragenumlen

# Write Data
excel_title = ['名次', '视频名称', 'up主', '播放量', '弹幕量']
for i in range(0, 5):
    worksheet.write(0, i, excel_title[i])

for i in range(0, 100):
    worksheet.write(i + 1, 0, i + 1)
    worksheet.write(i + 1, 1, title_name[i])
    worksheet.write(i + 1, 2, up_name[i])
    worksheet.write(i + 1, 3, play_number[i])
    worksheet.write(i + 1, 4, barrage_number[i])

# Save File
workbook.save(str(time.strftime('%Y-%m-%d-%H.%M', time.localtime(time.time())))+' BiliBili播放排行榜.xls')


