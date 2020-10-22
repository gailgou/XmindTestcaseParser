from lib.readXmind import *
from lib.writeExcel import *
import logging


def get_all_xmind_case(param):
    # 生成测试用例
    read_xmind = ReadXmindList(param['xmind_path'])
    write_excel = WriteExcel(param['excel_path'])
    testcase_list = []
    read_xmind.read_all_case(read_xmind.content, testcase_list, write_excel, param)
    write_excel.save_excel()
    logging.info("Generate Xmind file successfully: {}".format(param['excel_path']))


def get_develop_xmind_case(param):
    # 生成测试用例
    read_xmind = ReadXmindList(param['xmind_path'])
    write_excel = WriteExcel(param['excel_path'])
    testcase_list = []
    read_xmind.read_develop_case(read_xmind.content, testcase_list, write_excel, param)
    # write_excel.write_analysis_worksheet()
    write_excel.save_excel()
    logging.info("Generate Xmind file successfully: {}".format(param['excel_path']))


def get_num_and_rate(param):
    read_xmind = ReadXmindList(param['xmind_path'])
    res = read_xmind.get_case_num_and_rate()
    res['pass_rate'] = int(res['pass']/res['total']*100)
    res['perform_rate'] = int((res['total']-res['not_run'])/res['total']*100)
    return res
