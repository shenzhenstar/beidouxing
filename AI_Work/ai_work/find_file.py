#coding=utf-8
#date=2020/06/23

import os
PATH=lambda p: os.path.abspath(
    os.path.join(os.path.dirname(__file__),p)
)


class FindFile(object):
    """
    从目录里找Excel表格
    """
    def __init__(self,mohu_filename):
        """
        初始化
        :param mohu_filename: 模糊查询的文件名字
        """
        self.mohu_filename = mohu_filename
    def find_file(self):
        """
        遍历文件，并返回文件名称
        :return:
        """
        file_name_list = os.listdir(PATH('../file_data'))
        for f in file_name_list:
            if self.mohu_filename in f:
                return f
            if self.mohu_filename in f:
                return f


if __name__ == '__main__':
    # ce = FindFile('考勤统计表')
    ce = FindFile('薪薪乐考勤数据')
    print(ce.find_file(),type(ce.find_file()))