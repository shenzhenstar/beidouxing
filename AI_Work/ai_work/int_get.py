#coding=utf-8
#date=2020/06/13 0013

import configparser
import os

class IniOps(object):
    """
    从ini格式配置文件中读取配置信息
    """
    def __init__(self,filepath):
        cp=configparser.ConfigParser()
        cp.read(filepath,encoding='utf8')
        self.cp=cp
        self.file_conf='excel_file'
        self.excel_conf='excel_data'
        return

    def get_file_conf(self):
        """
        返回文件的配置信息
        :return:
        """
        options=self.cp.options(self.file_conf)
        conf={}
        # 遍历，迭代 iteration
        for option in options:
            value=self.cp.get(self.file_conf,option)
            conf[option]=value
        return conf
    def get_excel_conf(self):
        """
        返回Excel文件的格式配置信息
        :return:
        """
        options=self.cp.options(self.excel_conf)
        conf={}
        # 遍历，迭代 iteration
        for option in options:
            value=self.cp.get(self.excel_conf,option)
            if option.lower()=='CASE_START_LINE'.lower():
                value=int(value)
            conf[option]=value
        return conf


if __name__ == '__main__':
    file_name_list = os.listdir('../file_data')
    # print(file_name_list)
    for f in file_name_list:
        if '薪薪乐考勤数据' in f :
            print(f)
        if '考勤统计表' in f :
            print(f)
    obj=IniOps('../config.ini')
    print(obj.get_file_conf())
    # data11 = obj.get_excel_conf()
    # print(data11['put_weekend_s'])
    # print(data11)
    # for i in data11.values():
    #     print(i,type(i))
    # print(type(data11['pno']))