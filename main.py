# -*- coding:utf-8 -*-
'''
Author: Kevin.Zhang
E-Mail: testcn@vip.qq.com
'''

import subprocess
import sys
import xlwt
import json

windowsLib = '/lib/Windows/'
macLib = '/lib/MAC/'

sysPlatform = sys.platform
if 'win32' in sysPlatform:
    useLib = sys.path[0] + windowsLib + 'MediaInfo.exe'
elif 'darwin' in sysPlatform:
    # libPath = sys.path[0] + macLib
    print("Your system is not support!")

print("Your Lib Path: ", useLib)

mediaFile = r'E:\TestVideo\测试视频\6K\6K全景迪拜.mp4'
outJSON = 'testJson.json'

# cmd = [useLib, '--Output=JSON', mediaFile, '--LogFile=%s' % outJSON]
cmd = [useLib, '--Output=JSON', mediaFile]
print(cmd)
run = subprocess.Popen(cmd, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
# outPut = run.stdout.read()

# for i in outPut:
#     print(i.decode('utf-8'))
# print(outPut)
# print(type(outPut))

def save2xls(list):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('MediaInfo')

    for line in list:
        w = len(line)
    ws.write(2, 0, 1)
    ws.write(2, 1, 1)

    wb.save('MediaInfo.xls')


# 读取数据
with open(r'./6k_JSON.json', 'r',encoding='UTF-8') as f:
    data = json.load(f)

track = data['media']['track']

for i in track:
    print(i['@type'])
# 扩展名 FileExtension
# 格式 Format
# 文件大小 FileSize
# 平均码率 OverallBitRate

# 视频编码格式 Format
# 格式概况 Format_Level
# 码率 BitRate
# 宽度 Width
# 高度 Height
# 画面比例 DisplayAspectRatio
# 帧率 FrameRate
# 色彩空间 ColorSpace
# 色彩抽样 ChromaSubsampling
# 位深 BitDepth

# 格式 Format
# 码率 164088
# 声道 Channels
# 采样率 SamplingRate
# 压缩模式 Compression_Mode
# 语言 Language
def jsonTest(data):
    pass