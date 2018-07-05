#!/usr/bin/python3
# -*- coding: utf-8 -*-
import xlrd, xlwt, os
from xlutils.copy import copy
from download import request
from bs4 import BeautifulSoup
import time
import platform

system = platform.system()
xlsPath = ''
#根据系统识别路径
if system == 'Darwin':#mac
	originPath = os.path.abspath('.')
	xlsPath = os.path.join(originPath,'words rug pad.xls')
elif system == 'Windows':
	originPath = 'C:/Users/Administrator/Desktop/amzkeyword'
	xlsPath = os.path.join(originPath,'words rug pad.xls')


#保存xls表格
def savexls():
	os.remove(xlsPath)
	workbookCopy.save(xlsPath)
#获取当前表格的信息
def get_sheet_mes():
	workbook = xlrd.open_workbook(xlsPath)
	workbookCopy = copy(workbook)

	sheet_name = workbook.sheet_names()[0]
	sheet_one = workbook.sheet_by_name(sheet_name)
	wordlist = sheet_one.col_values(0)
	selllist = sheet_one.col_values(2)
	return workbookCopy, wordlist, selllist

def getdata(keyword):
	wordurl = 'https://www.amazon.com/s/ref=nb_sb_noss?url=search-alias=aps&field-keywords=%s&rh=i:aps,k:%s'%(keyword, keyword)
	html = request.get(wordurl,timeout = 10)
	html_soup = BeautifulSoup(html.text, 'lxml')
	span = html_soup.find('span', id = 's-result-count')
	if span != None:
		spanvalues = span.strings
		spanvalue = ''
		for x in spanvalues:
			spanvalue = x
			break
		values = spanvalue.split()[-3]
		values = values.replace(',', '')
		print(values)
		return values
	else:#无搜索量或者是网络问题
		print('无搜索量',wordurl)
		return ''

x, y, z = get_sheet_mes()

curr_num = 1      #当为第0行为表头
word_num = len(y) #表格总长度

while curr_num < word_num:
	workbookCopy, wordlist, selllist = get_sheet_mes()
	sheet = workbookCopy.get_sheet(0)
	sells = selllist[curr_num]

	if sells == '':
		current = curr_num
		for x in range(0,20):
			current = curr_num+x
			if current < word_num:
				currsells = selllist[current]
				if currsells == '':
					word = wordlist[current]
					#可以抓取数据
					print(current,word)
					string = getdata(word)
					if string == '':
						sheet.write(current, 2, string)
					else:
						sheet.write(current, 2, int(string))
					time.sleep(1)
			else:
				break
		savexls()
		curr_num = current + 1
	else:
		#直接定位到下一个为 空的位置
		curr_num += 1
		for x in range(curr_num,len(selllist)):
			print(x)
			curr_num = x
			sell = selllist[x]
			if sell == '':
				break




