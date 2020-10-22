import os, xlwt, datetime


# from conf.settings import filename_path

class WriteExcel:
    style = xlwt.easyxf(
        'pattern: pattern solid, fore_colour 0x31; font: bold on,height 250;alignment:HORZ CENTER,VERT CENTER;'
        'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
    style_nocenter = xlwt.easyxf('pattern: pattern solid, fore_colour White;'
                                 'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')  # 未居中无背景颜色0x34
    style_center = xlwt.easyxf('pattern: pattern solid, fore_colour White;alignment:HORZ CENTER;'
                               'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')  # 无背景颜色居中
    style_num = xlwt.easyxf('pattern: pattern solid, fore_colour White;alignment:HORZ CENTER;'
                            'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A',
                            num_format_str='#,##0.00')  # 无背景颜色居中

    def __init__(self, output_file):
        self.testcase_filename = output_file
        self.workbook = self.__init_excel()
        self.testcase_worksheet = self.__init_testcase_worksheet()
        self.analysis_worksheet = self.__init_analysis_worksheet()
        self.__row = 1
        self.__temp_list = []

    def __init_excel(self):
        """
        初始化excel
        :return: excel工作表
        """
        f = open(self.testcase_filename, 'w')
        f.close()
        workbook = xlwt.Workbook()
        return workbook

    def __init_outline_worksheet(self):
        """
        初始化测试大纲sheet
        :return:
        """
        outline_worksheet = self.workbook.add_sheet('测试大纲', cell_overwrite_ok='True')  # 测试大纲
        for i in range(13):
            outline_worksheet.col(i).width = (13 * 367)
            if i in [6, 7]:
                outline_worksheet.col(i).width = (25 * 500)
            if i in [7]:
                outline_worksheet.col(i).width = (20 * 400)
            if i in [4, 8, 9, 10, 11, 12]:
                outline_worksheet.col(i).width = (13 * 200)
        # 表头标题
        head = ['需求编号', '功能模块', '功能名称', '功能点', '用例类型', '检查点', '用例设计', '预期结果', '类别',
                '责任人', '状态', '更新日期', '用例编号']  # 子功能名称
        index = 0
        for head_item in head:
            outline_worksheet.write(0, index, head_item, self.style)
            index += 1
        self.save_excel()
        return outline_worksheet

    def __init_testcase_worksheet(self):
        """
        初始化测试用例sheet
        :return:
        """
        testcase_worksheet = self.workbook.add_sheet('测试用例', cell_overwrite_ok='True')
        for i in range(7):
            testcase_worksheet.col(i).width = (15 * 367)
            if i in [1, 2]:
                testcase_worksheet.col(i).width = (25 * 500)
            if i in [3]:
                testcase_worksheet.col(i).width = (20 * 400)
            if i in [4, 5]:
                testcase_worksheet.col(i).width = (13 * 200)
            if i in [6]:
                testcase_worksheet.col(i).width = (20 * 220)
        testcase_worksheet.row(0).height_mismatch = True
        testcase_worksheet.row(0).height = 20 * 40
        head = ['用例目录', '用例名称', '用例步骤', '预期结果', '创建人', '测试结果', '是否开发自测']
        index = 0
        for head_item in head:
            testcase_worksheet.write(0, index, head_item, self.style)
            index += 1
        self.save_excel()
        return testcase_worksheet

    def init_scope_worksheet(self):
        """
        初始化测试范围
        :return:
        """
        scope_worksheet = self.workbook.add_sheet('测试范围', cell_overwrite_ok='True')  # 测试范围
        for i in range(7):
            scope_worksheet.col(i).width = (13 * 367)
        # 表头标题
        head = ['序号', '功能模块', '功能名称', '角色', '责任人', '更新日期', '备注']  # 子功能名称
        index = 0
        for head_item in head:
            scope_worksheet.write(0, index, head_item, self.style)
            index += 1
        self.save_excel()
        return scope_worksheet

    def __init_analysis_worksheet(self):
        """
        初始化测试分析
        :return:
        """
        analysis_worksheet = self.workbook.add_sheet('测试分析', cell_overwrite_ok='True')
        for i in range(10):
            analysis_worksheet.col(i).width = (10 * 367)
            analysis_worksheet.write_merge(1, 1, 1, 9, '测试覆盖范围及执行结果', self.style)
        analysis_worksheet.write(2, 1, '项目', self.style_nocenter)
        analysis_worksheet.write_merge(2, 2, 2, 3, '', self.style_nocenter)
        analysis_worksheet.write(2, 4, '需求编号', self.style_nocenter)
        analysis_worksheet.write_merge(2, 2, 5, 7, '', self.style_nocenter)
        analysis_worksheet.write(2, 8, '产品', self.style_nocenter)
        analysis_worksheet.write(2, 9, '', self.style_nocenter)
        analysis_worksheet.write(3, 1, '投产日期', self.style_nocenter)
        analysis_worksheet.write_merge(3, 3, 2, 3, '', self.style_nocenter)
        analysis_worksheet.write(3, 4, '迭代编号', self.style_nocenter)
        analysis_worksheet.write_merge(3, 3, 5, 7, '', self.style_nocenter)
        analysis_worksheet.write(3, 8, '测试周期', self.style_nocenter)
        analysis_worksheet.write(3, 9, '', self.style_nocenter)
        analysis_worksheet.write(4, 1, '测试人员', self.style_nocenter)
        analysis_worksheet.write_merge(4, 4, 2, 9, '', self.style_nocenter)
        analysis_worksheet.write(5, 1, '环境', self.style_nocenter)
        analysis_worksheet.write_merge(5, 5, 2, 9, 'SIT/UAT/生产', self.style_nocenter)
        # 第 7 行内容
        lines = ['模块', 'Total', 'Pass', 'Fail', 'Block', 'NA', 'Not Run', 'Run Rate', 'Pass Rate']
        index = 1
        for head_item in lines:
            analysis_worksheet.write(6, index, head_item, self.style)
            index += 1
        self.save_excel()
        return analysis_worksheet

    def write_outline_excel(self, new_testcase):
        """
        写入测试大纲excel
        :param new_testcase:写入的列表信息
        :return:
        """
        style = xlwt.easyxf('borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
        for i in range(13):
            self.outline_worksheet.write(self.__row, i, "", style)
        col = 1
        for item in new_testcase[0][1:]:
            if col == 4:
                self.outline_worksheet.write(self.__row, col, '功能', style)
                col += 1
                self.outline_worksheet.write(self.__row, col, item, style)
            else:
                self.outline_worksheet.write(self.__row, col, item, style)
            col += 1
        if new_testcase[1]:
            self.outline_worksheet.write(self.__row, 7, new_testcase[1], style)
        self.outline_worksheet.write(self.__row, 4, '功能', style)
        self.outline_worksheet.write(self.__row, 10, 'Not Run', style)
        date_time = datetime.date.today()
        self.outline_worksheet.write(self.__row, 11, str(date_time), style)
        self.__row += 1

    def write_testcase_excel(self, testcase):
        """
        写入测试用例excel
        :param testcase:写入的列表信息
        :return:
        """
        style = xlwt.easyxf(
            'align: wrap on; font:height 200;'
            'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')

        self.testcase_worksheet.write(self.__row, 0, testcase['module'], style)  # 模块
        self.testcase_worksheet.write(self.__row, 1, testcase['title'], style)  # 用例名称
        self.testcase_worksheet.write(self.__row, 2, testcase['steps'], style)  # 用例名称
        self.testcase_worksheet.write(self.__row, 3, testcase['result'], style)  # 预期结果
        self.testcase_worksheet.write(self.__row, 4, testcase['creator'], style)

        self.testcase_worksheet.write(self.__row, 5, '', style)
        if testcase['is_develop_case']:
            self.testcase_worksheet.write(self.__row, 6, '是', style)
        else:
            self.testcase_worksheet.write(self.__row, 6, '否', style)

        self.__row += 1

    def write_scope_worksheet(self, new_testcase):
        """
        写入测试范围sheet
        :param new_testcase:
        :return:
        """
        style = xlwt.easyxf('borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
        item_list = new_testcase[0][1:3]
        col = 1
        if item_list not in self.__temp_list:
            self.__temp_list.append(item_list)
            for i in range(7):
                self.scope_worksheet.write(self.__scope_row, i, "", style)
            for item in item_list:
                self.scope_worksheet.write(self.__scope_row, col, item, style)
                col += 1
            self.__scope_row += 1

    def write_analysis_worksheet(self):
        """
        写入测试分析excel
        :return:
        """
        row = 7
        temp_list = []
        for item in self.__temp_list:
            if len(item) >= 1:
                if item[0] not in temp_list:
                    temp_list.append(item[0])
                    self.analysis_worksheet.write(row, 1, item[0], self.style_center)
                    self.analysis_worksheet.write(row, 2, "=COUNTIF(测试用例!A:A,B" + str(row + 1) + ")", self.style_center)
                    self.analysis_worksheet.write(row, 3,
                                                  "=COUNTIFS(测试用例!A:A,B" + str(row + 1) + ''',测试用例!K:K,"Pass")''',
                                                  self.style_center)
                    self.analysis_worksheet.write(row, 4,
                                                  "=COUNTIFS(测试用例!A:A,B" + str(row + 1) + ''',测试用例!K:K,"Fail")''',
                                                  self.style_center)
                    self.analysis_worksheet.write(row, 5,
                                                  "=COUNTIFS(测试用例!A:A,B" + str(row + 1) + ''',测试用例!K:K,"Block")''',
                                                  self.style_center)
                    self.analysis_worksheet.write(row, 6, "=COUNTIFS(测试用例!A:A,B" + str(row + 1) + ''',测试用例!K:K,"NA")''',
                                                  self.style_center)
                    self.analysis_worksheet.write(row, 7,
                                                  "=COUNTIFS(测试用例!A:A,B" + str(row + 1) + ''',测试用例!K:K,"Not Run")''',
                                                  self.style_center)
                    self.analysis_worksheet.write(row, 8, xlwt.Formula(
                        "SUM(D" + str(row + 1) + ":F" + str(row + 1) + ")/(C" + str(row + 1) + "-G" + str(
                            row + 1) + ")"), self.style_num)
                    self.analysis_worksheet.write(row, 9, xlwt.Formula(
                        "D" + str(row + 1) + "/(C" + str(row + 1) + "-G" + str(row + 1) + ")"), self.style_num)
                    row += 1
            else:
                lines = ['', 0, 0, 0, 0, 0, 0, '0.00%', '0.00%']
                index = 1
                for head_item in lines:
                    self.analysis_worksheet.write(row, index, head_item, self.style)
                    index += 1
                row += 1
        self.analysis_worksheet.write(row, 1, '总计', self.style_center)
        self.analysis_worksheet.write(row, 2, xlwt.Formula("SUM(C8:C" + str(row) + ")"), self.style_center)
        self.analysis_worksheet.write(row, 3, xlwt.Formula("SUM(D8:D" + str(row) + ")"), self.style_center)
        self.analysis_worksheet.write(row, 4, xlwt.Formula("SUM(E8:E" + str(row) + ")"), self.style_center)
        self.analysis_worksheet.write(row, 5, xlwt.Formula("SUM(F8:F" + str(row) + ")"), self.style_center)
        self.analysis_worksheet.write(row, 6, xlwt.Formula("SUM(G8:G" + str(row) + ")"), self.style_center)
        self.analysis_worksheet.write(row, 7, xlwt.Formula("SUM(H8:H" + str(row) + ")"), self.style_center)
        self.analysis_worksheet.write(row, 8, xlwt.Formula(
            "SUM(D" + str(row + 1) + ":F" + str(row + 1) + ")/(C" + str(row + 1) + "-G" + str(row + 1) + ")"),
                                      self.style_num)
        self.analysis_worksheet.write(row, 9, xlwt.Formula(
            "D" + str(row + 1) + "/(C" + str(row + 1) + "-G" + str(row + 1) + ")"), self.style_num)
        row += 2
        self.analysis_worksheet.write(row, 1, '说明:')
        row += 1
        self.analysis_worksheet.write(row, 1, 'Pass-验证通过  Fail-验证未通过  Block-阻塞  NA-本期不涉及  Not Run-尚未执行')
        row += 1
        self.analysis_worksheet.write(row, 1, 'Run Rate=(Pass+Fail+Block)/(Total-NA)')
        self.analysis_worksheet.write(row + 1, 1, 'Pass Rate=Pass/(Total-NA)')

    def save_excel(self):
        """
        保存excel
        :return:
        """
        self.workbook.save(self.testcase_filename)
