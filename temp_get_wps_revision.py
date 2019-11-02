'''
Created on 2019年10月25日

@author: liushucheng
@brief:  Review议事录中记载的成果物路径没有svn信息，
      :  这个脚本用来获得这个成果物的svn revision信息，
      :  路径信息需要手动从review议事录中拷贝到指定路径下的文件中
'''
# coding=utf-8
import os
import re
import pysvn

svn_cmd_prefix = "TortoiseProc.exe /command:update /path:"
svn_cmd_suffix = "/closeonend:1"

# 把要获取svn版本号的文件路径保存到这个文件中,文件路径需要从review议事录中拷贝
file_loc_list = "C:\\Users\\liushucheng\\Desktop\\new.txt"
# 根据从上面文件中读取的文件路径信息构造文件更新脚本文件
svn_update_cript = "C:\\Users\\liushucheng\\Desktop\\update.bat"

# 文件路径中包含的公共部分
common_path = "https://172.29.44.209/automotive/autosar/svnASG_D008633/dev/bsw/Autosar_R40/branch/dev_mcal_F1Kx_R40_R42_all_variant"
# 以读方式打开文件
open_file = open(file_loc_list,'r',encoding='utf-8')
# 以写方式打开文件
script_file = open(svn_update_cript,'w',encoding='utf-8')

for line,value in enumerate(open_file):
    value = re.sub(common_path,'U:',value)
    if '\n' in value:
        value = re.sub(r'\s*\n$','',value)
        svn_cmd = svn_cmd_prefix + '\"' + value + '\" ' + svn_cmd_suffix
        # print(svn_cmd)
        print(svn_cmd, file = script_file)
    else:
        pass
# 关闭文件
script_file.close()
# 执行脚本文件,从SVN服务器更新文件
os.system(svn_update_cript)
# 关闭文件
open_file.close()
# 打开文件
open_file = open(file_loc_list,'r',encoding='utf-8')

client = pysvn.Client()

item_id = 1
# 遍历文件
for line,value in enumerate(open_file):
    # 替换路径中的公共部分
    value = re.sub(common_path,'U:',value)
    # 替换路径中的制表回车符号
    value = re.sub(r'\s*\n','',value)
    # print(value)
    try:
        entry = client.info(value)
        # 打印文件路径及svn版本信息
        print(value + ',ID:' + str(item_id) + ',svn revision: ' + str(entry.commit_revision.number))
    except:
        print(value + ' path is not correct.')
    item_id += 1

