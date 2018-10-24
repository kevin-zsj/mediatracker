# -*- coding:utf-8 -*-
'''
Author: Kevin.Zhang
E-Mail: testcn@vip.qq.com
'''

import subprocess
import sys
import xlwt
import xml.etree.ElementTree as ET

windowsLib = '/lib/Windows/'
macLib = '/lib/MAC/'

sysPlatform = sys.platform
if 'win32' in sysPlatform:
    useLib = sys.path[0] + windowsLib + 'MediaInfo.exe'
    print("Your Lib Path: ", useLib)
elif 'darwin' in sysPlatform:
    # libPath = sys.path[0] + macLib
    print("Your system is not support!")


# mediaFile = r'E:\TestVideo\测试视频\6K\6K全景迪拜.mp4'
# outJSON = 'testJson.json'
#
# # cmd = [useLib, '--Output=JSON', mediaFile, '--LogFile=%s' % outJSON]
# cmd = [useLib, '--Output=JSON', mediaFile]
# print(cmd)
# run = subprocess.Popen(cmd, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
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


class getInfo:

    def __init__(self, xmlFile):
        if xmlFile:
            self.tree = ET.parse(xmlFile)
            self.root = self.tree.getroot()[1]
            self.info = dict()

    def getGeneralInfo(self):
        # 文件路径
        print('文件路径：', self.root.attrib['ref'])
        self.info['generalVideoFilePath'] = self.root.attrib['ref']
        for child in self.root[0]:
            # 扩展名 FileExtension
            # print('扩展名：')
            # 格式 Format
            if child.tag.strip().endswith('Format'):
                print('视频格式', ":", child.text)
                self.info['generalFormat'] = child.text
            # 文件大小 FileSize
            if child.tag.strip().endswith('FileSize'):
                print('视频大小', ":", child.text)
                self.info['generalVideoFileSize'] = child.text
            # 平均码率 OverallBitRate
            if child.tag.strip().endswith('OverallBitRate'):
                print('视频平均码率', ":", child.text)
                self.info['generalVideoOverallBitRate'] = child.text

    def getVideoInfo(self):
        for child in self.root[1]:
            # 视频编码格式 Format
            if child.tag.strip().endswith('Format'):
                print('视频编码格式', ":", child.text)
                self.info['videoEncodeFormat'] = child.text
            # 格式概况 Format_Level
            if child.tag.strip().endswith('Format_Level'):
                print('视频格式概况', ":", child.text)
                self.info['videoFormat_Level'] = child.text
            # 码率 BitRate
            if child.tag.strip().endswith('BitRate'):
                print('视频码率', ":", child.text)
                self.info['videoBitRate'] = child.text
            # 宽度 Width
            if child.tag.strip().endswith('Width'):
                print('视频宽度', ":", child.text)
                self.info['videoWidth'] = child.text
            # 高度 Height
            if child.tag.strip().endswith('Height'):
                print('视频高度', ":", child.text)
                self.info['videoHeight'] = child.text
            # 画面比例 DisplayAspectRatio
            if child.tag.strip().endswith('DisplayAspectRatio'):
                print('视频画面比例', ":", child.text)
                self.info['videoDisplayAspectRatio'] = child.text
            # 帧率 FrameRate
            if child.tag.strip().endswith('FrameRate'):
                print('视频画面帧率', ":", child.text)
                self.info['videoFrameRate'] = child.text
            # 色彩空间 ColorSpace
            if child.tag.strip().endswith('ColorSpace'):
                print('视频色彩空间', ":", child.text)
                self.info['videoColorSpace'] = child.text
            # 色彩抽样 ChromaSubsampling
            if child.tag.strip().endswith('ChromaSubsampling'):
                print('视频色彩抽样', ":", child.text)
                self.info['videoChromaSubsampling'] = child.text
            # 位深 BitDepth
            if child.tag.strip().endswith('BitDepth'):
                print('视频位深', ":", child.text)
                self.info['videoBitDepth'] = child.text

    def getAudioInfo(self):
        for child in self.root[2]:
            # 格式 Format
            if child.tag.strip().endswith('Format'):
                print('音频编码格式', ":", child.text)
                self.info['audioEncodeFormat'] = child.text
            # 码率 BitRate
            if child.tag.strip().endswith('BitRate'):
                print('音频码率', ":", child.text)
                self.info['audioBitRate'] = child.text
            # 声道 Channels
            if child.tag.strip().endswith('Channels'):
                print('音频声道', ":", child.text)
                self.info['audioChannels'] = child.text
            # 采样率 SamplingRate
            if child.tag.strip().endswith('SamplingRate'):
                print('音频采样率', ":", child.text)
                self.info['audioSamplingRate'] = child.text
            # 压缩模式 Compression_Mode
            if child.tag.strip().endswith('Compression_Mode'):
                print('音频压缩模式', ":", child.text)
                self.info['audioCompression_Mode'] = child.text
            # 语言 Language
            if child.tag.strip().endswith('Language'):
                print('音频语言', ":", child.text)
                self.info['audioLanguage'] = child.text

    def main(self):
        self.getGeneralInfo()
        self.getVideoInfo()
        self.getAudioInfo()
        return self.info


if __name__ == '__main__':
    xml = '6K-Video.xml'
    info = getInfo(xml).main()
    print(info)
    print('Count :', len(info))
