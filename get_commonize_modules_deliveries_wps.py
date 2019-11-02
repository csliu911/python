'''
Created on 2019年10月21日
@author: liushucheng
@brief : 这个脚本是临时用途，仅用来遍历共通化的三个模块(CAN, FR, SPI)中的文件列表,从中提取需要交付的成果物
       : 来填充safety plan products deliveries中的成果物清单，正确列出需要提交的成果物清单后，以后只需要
       : 根据成果物清单的路径用get_doc_num..._attachment.py脚本获取文件版本信息即可
'''
# coding=utf-8
import os
import pysvn

print('****************************************************************************************************************')
print('-------------------- Walking through folder of commonize module to make a list of WPs --------------------------')
print('****************************************************************************************************************')

client = pysvn.Client()

msn = ['can','spi','fr']
for index in range(len(msn)):
    path_common_platform = "U:\\external\\X1X\\common_platform\\modules\\" + msn[index] + "\\"
    for root, dirs, files in os.walk(path_common_platform, topdown = False):
        for name in files:
            full_path = os.path.join(root, name)
            entry = client.info(full_path)
            # print(full_path + ',(SVN:' + str(entry.commit_revision.number) + ')')
            print(full_path)
        # for name in dirs:
        #     print(os.path.join(root, name))
    print('')
    path_F1x = "U:\\external\\X1X\\F1x\\modules\\" + msn[index] + "\\"
    for root, dirs, files in os.walk(path_F1x, topdown = False):
        for name in files:
            full_path = os.path.join(root, name)
            entry = client.info(full_path)
            # print(full_path + ',(SVN:' + str(entry.commit_revision.number) + ')')
            print(full_path)
        # for name in dirs:
        #     print(os.path.join(root, name))
    print('')
path_generic = "U:\\external\\X1X\\common_platform\\generic\\"
for root, dirs, files in os.walk(path_generic, topdown = False):
    for name in files:
        full_path = os.path.join(root, name)
        entry = client.info(full_path)
        # print(full_path + ',(SVN:' + str(entry.commit_revision.number) + ')')
        print(full_path)
    # for name in dirs:
    #     print(os.path.join(root, name))
