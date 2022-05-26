#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os,re
path=input('请输入文件路径(结尾加上/)：')       


fileList=os.listdir(path)			#获取该目录下所有文件，存入列表中

'''
#part 1

n=0
for i in fileList:

	name_separate = os.path.splitext(i) #分割后缀名
	name_forward = name_separate[0]
	name_suffix = name_separate[1]

	# 设置旧文件名（就是路径+文件名）
	oldname = path + os.sep + fileList[n]  # os.sep添加系统分隔符

	if re.match(r'^([a-zA-Z]+)(\d+)(.*)',name_forward) or re.match(r'^([a-zA-Z]+)(\-{2,})(\d+)(.*)',name_forward): 	#检测是否有分割线-

		m = re.match(r'^([a-zA-Z]*)(\-*)(\d+)(.*)', name_forward)
		group1 = m.group(1)
		group2 = m.group(2)
		group3 = m.group(3)
		group4 = m.group(4)

		#newname = os.path.join(path, group1 + '-' + group3 + '-RION' + '-' + group4 + name_suffix)
		newname = os.path.join(path, group1 + '-' + group3 + '-RION' + name_suffix)

		os.rename(oldname,newname )  #用os模块中的rename方法对文件改名
		print(oldname,'======>',newname) 

	n+=1
'''

'''
# #part 2
# n=0
# for i in fileList:

	# name_separate = os.path.splitext(i) #分割后缀名
	# name_forward = name_separate[0]
	# name_suffix = name_separate[1]

	# oldname = path + os.sep + fileList[n]
	
	# newname = os.path.join(path, 'QQ视频 ' + str(n+1) + name_suffix)
	
	# os.rename(oldname, newname)
	# print(oldname,'======>',newname)
	
	# n+=1
	
	'''


n = 0
for i in fileList:
	name_separate = os.path.splitext(i)
	name_forwar = name_separate[0]
	name_suffix = name_separate[1]
	
	oldname = path + os.sep +fileList[n]
	newname = os.path.join(path, str(n+1) + name_suffix)
	os.rename(oldname,newname )  #用os模块中的rename方法对文件改名
	print(oldname,'======>',newname) 

	n+=1