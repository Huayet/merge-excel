# -*- coding: utf-8 -*- 

''' 
Created on 2016-08-03
@author：孙华琛
@description: 扫描指定文件夹（即根目录），把它下面所有子文件夹的文件，复制到根目录下
			29-30行定义了，小于30kb的文件不复制，因为本场景小于30kb的都是空文件
''' 

import os
import shutil

#
#rootDir:递归文件夹（初始是目的文件夹）
#root   :目的文件夹，相当于整个函数中的静态的变量
#
def moveFile(rootDir,root):
	print u'目的目录：',root
	for lists in os.listdir(rootDir):
		print u'正在移动的文件名是:',lists
		path = os.path.join(rootDir, lists) #实际目录+文件名
		print u'文件大小：',os.path.getsize(path)
		rootpath = root+'\\'+lists #os.path.join(root, lists) #目标目录+文件名
		print u'原始路径：',path
		print u'目的路径：',rootpath
		if os.path.exists(rootpath):
			pass
		else:
			if(os.path.getsize(path)>30000): #没有数据的文件都是小于30kb,就不拿出来了
				shutil.copy(path, rootpath)  #copy如果换成move就是移动
		print u'————————————————————————————————————————'
		#如果是文件夹 递归
		if os.path.isdir(path):
			moveFile(path,root)

#把当前路径下的所有文件，复制到根目录下
moveFile(os.getcwd(),os.getcwd())