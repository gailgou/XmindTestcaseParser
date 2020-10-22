import sys
from xmindparser import xmind_to_dict
from .writeExcel import *
from .constant import Const


class ReadXmindList(object):
    def __init__(self, filename):
        self.filename = filename  # xmind文件路径
        self.content, self.canvas_name, self.excel_title = self.__get_dic_content(self.filename)
        self.title_length = 4
        self.develop_count = 0

        Const.PASS = 0

    @staticmethod
    def __get_dic_content(filename):
        """
        #从Xmind读取字典形式的数据
        :param filename: xmind文件路径
        :return:
        """
        if not os.path.exists(filename):
            print("[ERROR] 文件不存在")
            sys.exit(-1)
        out = xmind_to_dict(filename)
        dic_content = out[0]
        canvas_name = dic_content.get('title')  # 获取画布名称
        canvas_values = dic_content.get('topic')
        if canvas_values:
            excel_title = canvas_values.get('title')  # 获取模块名称
        content = [dic_content.get('topic')]
        return content, canvas_name, excel_title

    def __format_list(self, testcase_list, param):
        """
        格式化为excel需要的列表
        :param testcase_list: 需要处理的列表
        :return:
        """
        step_list = []
        testcase_title, except_result, test_module = "", "", ""
        step_index = 1
        is_develop_case = False
        for item in testcase_list:
            if item['is_develop_case']:
                is_develop_case = True
            item_index = testcase_list.index(item)
            if 1 < item_index <= 1 + self.title_length:
                if testcase_title == "":
                    testcase_title = item['title']
                else:
                    testcase_title += '-' + item['title']
            if item_index == 1:
                test_module = item['title']
            if param['is_need_result'] == 'y':
                if item_index != len(testcase_list) - 1 and item_index > 1:
                    if step_index != 1:
                        step_list.append('\n' + str(step_index) + '.' + item['title'])
                    else:
                        step_list.append(str(step_index) + '.' + item['title'])
                    step_index += 1
                elif item_index == len(testcase_list) - 1:
                    except_result = item['title']
            elif param['is_need_result'] == 'n':
                if item_index > 1:
                    if step_index != 1:
                        step_list.append('\n' + str(step_index) + '.' + item['title'])
                    else:
                        step_list.append(str(step_index) + '.' + item['title'])
                    step_index += 1
        testcase = {
            'title': testcase_title,
            'module': test_module,
            'result': except_result,
            'steps': step_list,
            'creator': param['creator'],
            'is_develop_case': is_develop_case
        }
        return testcase

    def __format_develop_list(self, testcase_list, param):
        """
        格式化为excel需要的列表
        :param testcase_list: 需要处理的列表
        :return:
        """
        step_list = []
        testcase_title, except_result, test_module = "", "", ""
        step_index = 1
        for item in testcase_list:
            item_index = testcase_list.index(item)
            if 1 < item_index <= 1 + self.title_length:
                if testcase_title == "":
                    testcase_title = item['title']
                else:
                    testcase_title += '-' + item['title']
            if item_index == 1:
                test_module = item['title']
            if param['is_need_result'] == 'y':
                if item_index != len(testcase_list) - 1 and item_index > 1:
                    if step_index != 1:
                        step_list.append('\n' + str(step_index) + '.' + item['title'])
                    else:
                        step_list.append(str(step_index) + '.' + item['title'])
                    step_index += 1
                elif item_index == len(testcase_list) - 1:
                    except_result = item['title']
            elif param['is_need_result'] == 'n':
                if item_index > 1:
                    if step_index != 1:
                        step_list.append('\n' + str(step_index) + '.' + item['title'])
                    else:
                        step_list.append(str(step_index) + '.' + item['title'])
                    step_index += 1
        testcase = {
            'title': testcase_title,
            'module': test_module,
            'result': except_result,
            'steps': step_list,
            'creator': param['creator'],
            'is_develop_case': True
        }
        return testcase

    def read_all_case(self, content, testcase_list, write_excel, param):
        for dic_val in content:
            if 'makers' in dic_val and 'priority-1' in dic_val['makers']:
                is_develop_case = True
                self.develop_count += 1
            else:
                is_develop_case = False
            if 'title' in dic_val:
                title = dic_val['title']
            else:
                title = ""
            dot = {
                'title': title,
                'is_develop_case': is_develop_case
            }
            testcase_list.append(dot)
            if 'topics' in dic_val:
                self.read_all_case(dic_val['topics'], testcase_list, write_excel, param)
            else:
                testcase = self.__format_list(testcase_list, param)
                write_excel.write_testcase_excel(testcase)  # 写入测试用例
                testcase_list.pop()
        if testcase_list:
            testcase_list.pop()

    def read_develop_case(self, content, testcase_list, write_excel, param):
        for dic_val in content:
            if 'makers' in dic_val and 'priority-1' in dic_val['makers']:
                is_develop_case = True
                self.develop_count += 1
            else:
                is_develop_case = False
            if 'title' in dic_val:
                title = dic_val['title']
            else:
                title = ""
            dot = {
                'title': title,
                'is_develop_case': is_develop_case
            }
            testcase_list.append(dot)

            if 'topics' in dic_val:
                self.read_develop_case(dic_val['topics'], testcase_list, write_excel, param)
            else:
                if self.develop_count > 0:
                    testcase = self.__format_develop_list(testcase_list, param)
                    write_excel.write_testcase_excel(testcase)  # 写入测试用例
                case = testcase_list.pop()
                if case['is_develop_case']:
                    self.develop_count -= 1
        if testcase_list:
            case = testcase_list.pop()
            if case['is_develop_case']:
                self.develop_count -= 1

    def get_case_num_and_rate(self):
        res = {
            'total': 0,
            'pass': 0,
            'fail': 0,
            'not_run': 0,
            'not_run_list': []
        }
        testcase = []
        self.get_result(self.content, testcase, res)
        return res

    def get_result(self, content, testcase, res):
        for dic_val in content:
            if 'makers' in dic_val:
                if 'flag-green' in dic_val['makers']:
                    status = 0
                elif 'flag-red' in dic_val['makers']:
                    status = 1
                else:
                    status = 2
            else:
                status = 2
            title = dic_val['title'] if ('title' in dic_val) else ""
            dot = {
                'title': title,
                'status': status
            }
            testcase.append(dot)

            if 'topics' in dic_val:
                self.get_result(dic_val['topics'], testcase, res)
            else:
                res['total'] += 1
                case_detail = ""
                for dot in reversed(testcase):
                    case_detail = dot['title'] + '_' + case_detail if (case_detail != "") else dot['title']
                    if dot['status'] == 0:
                        res['pass'] += 1
                        break
                    if dot['status'] == 1:
                        res['fail'] += 1
                        break
                    else:
                        index = testcase.index(dot)
                        if index == 0:
                            res['not_run'] += 1
                            res['not_run_list'].append(case_detail)
                        else:
                            continue
                testcase.pop()
        if testcase:
            testcase.pop()
