#coding=utf-8
#date=2020/06/23

from ai_work.excel_get import ExcelOps
from ai_work.excel_write import WriteExcel
from ai_work.find_file import FindFile
from ai_work.int_get import IniOps
import time
import sys
import os
PATH=lambda p: os.path.abspath(
    os.path.join(os.path.dirname(__file__),p)
)


"""临时一个想法，把print的内容写到txt中，方便去分析。就先在这里声明这个类吧！！！"""
class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "a",encoding="utf8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass


def main():
    """把print的内容写到日志"""
    type = sys.getfilesystemencoding()
    sys.stdout = Logger(PATH('./log/运行日志.txt'))

    """入口"""
    print("======================开始运行=========================")
    file_conf = IniOps(PATH(r'./config.ini')).get_file_conf()
    file_name1 = FindFile(file_conf['a_excel'])
    file_name2 = FindFile(file_conf['b_excel'])
    file_string1 = './file_data/'+str(file_name1.find_file())
    file_string2 = './file_data/'+str(file_name2.find_file())
    ce1 = ExcelOps(PATH(file_string1))
    ret1 = ce1.read_data()
    ce2 = WriteExcel(PATH(file_string2))
    ret = ce2.working()

if __name__ == '__main__':
    main()
    print(time.asctime())
    print("----------------------执行完毕-------------------------")