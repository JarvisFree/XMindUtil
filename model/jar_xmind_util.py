#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@Time    ：2020/9/4 15:07
@Author  ：维斯
@File    ：jar_excel_util.py
@Version ：1.0
@Function：XMind解析工具工具
"""

import json
import xlrd
import xlwt
import xmind
from xlutils.copy import copy
from model.jar_excel_util import JarExcelUtil
from common.jar_project_util import JarProjectUtil


class JarXMindUtil:
    def __init__(self):
        # markers字段 "markers": ["star-orange"]
        self.model = ['star-dark-gray', '$所属模块$']  # [所属模块（灰色星星标志）,标识]
        self.check = ['star-orange', '$验证点$']  # [验证点（黄色星星标志）,标识]
        self.except_case = ['people-red', '$异常$']  # [异常用例（红色人像标志）,标识]
        self.normal_case = ['people-green', '$正常$']  # [正常用例（绿色人像标志）,标识]
        self.priority_1 = ['priority-1', '$高$']  # [重要级别-高（1号优先级标志）,标识]
        self.priority_2 = ['priority-2', '$中$']  # [重要级别-中（2号优先级标志）,标识]
        self.priority_3 = ['priority-3', '$低$']  # [重要级别-低（3号优先级标志）,标识]
        self.result_pass = ['flag-green', '$通过$']  # 用例执行结果 通过
        self.result_no_pass = ['flag-red', '$未通过$']  # 用例执行结果 未通过
        self.result_is_success = [
            self.result_pass,
            self.result_no_pass
        ]

        # 汇总
        self.all_markers = [self.model,
                            self.check,
                            self.except_case,
                            self.normal_case,
                            self.priority_1,
                            self.priority_2,
                            self.priority_3,
                            ]

        self.note = ['note', '$备注节点$']
        self.comment = ['comment', '$批注节点$']

        # 节点分割符
        self.node_split = ' /'  # 节点分割符
        self.node_split_excel = ' /'  # 写入表格中的多节点分割符
        # 预期结果分割符
        self.node_expect_split = '预期结果：'

        # 用例编号前缀
        self.case_number_model = 'CASE_200917'  # 生成的用例编号 如 CASE_200917000001、CASE_200917000002

        # XMind备注信息模板解析（严格按照此顺序解析）
        self.xmind_note_model = [['【前置条件】', '内容'],
                                 ['【操作步骤】', '内容'],
                                 ['【SQL校验】', '内容'],
                                 ['【预期结果】', '内容'],
                                 ['【备注】', '内容']]

        # 测试用例表格
        self.headers = [
            ['', 15],
            ['用例编号', len(self.case_number_model) + 5 + 5],  # 长度多5个字符
            ['用例类型'],
            ['重要级别'],
            ['所属模块', 30],
            ['验证点', 20],
            ['用例标题', 40],
            ['前置条件', 15],
            ['操作步骤', 15],
            ['SQL校验', 15],
            ['预期结果', 15],
            ['备注', 15]
        ]

    def to_excel(self, out_file, result_node: list):
        """
        解析的XMind数据写入表格文件
        :param out_file: 表格文件
        :param result_node: 列表 每一个元素为一个完整的xmind路径
        """
        body_data = []
        count = 0
        for re in result_node:
            count += 1
            # 1 分割某链路所有节点
            node_list = re.split(self.node_split)
            # 2 判断节点属性
            re_model = ''  # 所属模块
            re_check = ''  # 验证点
            re_case_type = ''  # 用例类型
            re_case_name = ''  # 用例标题
            re_priority = ''  # 重要级别
            re_expect = ''  # 预期结果
            re_note = ''  # 备注（XMind中的备注）
            re_comment = ''  # 批注（XMind中的批注）
            re_result = ''  # 用例执行结果

            for n in node_list:
                n: str
                # 2.1 所属模块
                if n.startswith(self.model[1]):
                    re_model += n.replace(self.model[1], '') + self.node_split_excel
                # 2.3 验证点
                if n.startswith(self.check[1]):
                    re_check += n.replace(self.check[1], '') + self.node_split_excel
                # 2.4 用例类型&用例标题
                if n.startswith(self.except_case[1]) or n.startswith(self.normal_case[1]):
                    if n.startswith(self.except_case[1]):
                        re_case_type = self.except_case[1][1:-1]
                        re_case_type_swap1 = n.split(self.except_case[1])
                        if re_case_type_swap1 == 3:
                            re_case_name += re_case_type_swap1[1] + self.node_split_excel
                        else:
                            re_case_name += n.replace(self.except_case[1], '') + self.node_split_excel
                    if n.startswith(self.normal_case[1]):
                        re_case_type = self.normal_case[1][1:-1]
                        re_case_type_swap2 = n.split(self.normal_case[1])
                        if re_case_type_swap2 == 3:
                            re_case_name += re_case_type_swap2[1] + self.node_split_excel
                        else:
                            re_case_name += n.replace(self.normal_case[1], '') + self.node_split_excel
                # 2.5 重要级别
                if n.startswith(self.priority_1[1]):
                    re_priority = self.priority_1[1][1:-1]
                if n.startswith(self.priority_2[1]):
                    re_priority = self.priority_2[1][1:-1]
                if n.startswith(self.priority_3[1]):
                    re_priority = self.priority_3[1][1:-1]
                # 2.6 预期结果
                re_expect = n[n.find(self.node_expect_split):]
                re_expect = re_expect[len(self.node_expect_split):]
                re_expect = JarProjectUtil.del_endswith_none(re_expect)
                # 2.7 备注
                re_note_swap = n.split(self.note[1])
                if len(re_note_swap) == 3:
                    re_note = re_note_swap[1]
                # 2.8 批注
                re_comment_swap = n.split(self.comment[1])
                if len(re_comment_swap) == 3:
                    re_comment = re_comment_swap[1]
                # 2.9 用例执行结果
                if n.startswith(self.result_pass[1]) or n.startswith(self.result_no_pass[1]):
                    # 通过
                    if n.startswith(self.result_pass[1]):
                        re_result = self.result_pass[1][1:-1]
                    # 未通过
                    if n.startswith(self.result_no_pass[1]):
                        re_result = self.result_no_pass[1][1:-1]

            # 删除末尾的分割符
            length = len(self.node_split_excel)
            re_model = re_model[:len(re_model) - length]
            re_check = re_check[:len(re_check) - length]
            re_case_name = re_case_name[:len(re_case_name) - length]
            # 删除用例名称后的 预期结果、空格、换行
            index = re_case_name.find(self.node_expect_split)
            re_case_name = re_case_name[:None if index == -1 else index]
            re_case_name = JarProjectUtil.del_endswith_none(re_case_name)

            # 解析XMind中的备注信息
            for n in range(len(self.xmind_note_model)):
                # 提取名称（如：【前置条件】 提取为 前置条件） 分别去除前后1个字符
                n_name = self.xmind_note_model[n][1:-1]
                # 提取此名称下的数据（如：提取哪些内容是前置条件）  截取XMind中备注信息中 n至n+1中的数据
                if n < len(self.xmind_note_model) - 1:
                    n_data = re_note[
                             re_note.find(self.xmind_note_model[n][0]):re_note.find(self.xmind_note_model[n + 1][0])]
                else:
                    n_data = re_note[re_note.find(self.xmind_note_model[n][0]):]
                n_data = n_data[len(self.xmind_note_model[n][0]):]
                n_data = JarProjectUtil.del_endswith_none(n_data)
                self.xmind_note_model[n][1] = n_data

            # 3 拼接节点至指定表格字段
            body = [
                re_result,  # 用例执行结果
                self.case_number_model + str(count).zfill(6),  # 用例编号
                re_case_type,  # 用例类型
                re_priority,  # 重要级别
                re_model,  # 所属模块
                re_check,  # 验证点
                re_case_name,  # 用例标题
                self.xmind_note_model[0][1],  # 前置条件
                self.xmind_note_model[1][1],  # 操作步骤
                self.xmind_note_model[2][1],  # SQL校验
                self.xmind_note_model[3][1],  # 预期结果
                self.xmind_note_model[4][1]  # 备注
            ]
            body_data.append(body)
        JarExcelUtil(header_list=self.headers).write(out_file=out_file, data_body=body_data)
        print('数据写入Excel完成！（路径：{}）'.format(out_file))

    @staticmethod
    def analysis(xmind_path):
        """
        解析XMind文件（获取每条完整路径）
        :param xmind_path: XMind文件
        :return: 返回所有路径list
        """
        wb = xmind.load(xmind_path)
        data = wb.to_prettify_json()
        data = json.loads(data)
        result_data = []
        # 画布
        for data_topic in data:
            # 1级
            data_1 = data_topic.get('topic')
            title_1 = data_1.get('title')

            # 递归后面所有级（2级、3级、......）
            JarXMindUtil().__base(title_1, data_1, result_data)
        print('{}，XMind数据解析完成！'.format(xmind_path))
        print(*result_data, sep='\n')
        return result_data

    def __base(self, title_long_x, data_topics_x, result_data_all):
        """
        递归XMind所有节点
        :param title_long_x:
        :param data_topics_x:
        :param result_data_all:
        """
        # 递归所有级
        topics_list = data_topics_x.get('topics')
        for topics in topics_list:  # 循环所有节点
            title = ''
            swap = ''
            # 取节点图标属性
            markers: list = topics.get('markers')
            if len(markers) != 0:
                for marker in markers:  # 循环当前节点的所有属性
                    for all_mar in self.all_markers:  # 循环预定的所有属性
                        if marker == all_mar[0]:  # 当前节点属性中有预定属性值
                            # 是用例名称节点
                            if marker == self.all_markers[2][0] or marker == self.all_markers[3][0]:  # 绿色人像或红色人像
                                # 获取节点备注值
                                s_note = topics.get('note') if topics.get('note') is not None else ''
                                # 获取节点批注值
                                s_comment = topics.get('comment') if topics.get('comment') is not None else ''
                                # 获取节点执行结果（旗子）
                                s_result = ''
                                for i_result in self.result_is_success:  # 遍历预定值（绿旗子、红旗子）
                                    for i_marker in markers:  # 遍历用例节点当前所有标识
                                        # 是某个预定旗子（只找一个就行了 因为多个旗子不可能同时存在）
                                        if str(i_result[0]) == str(i_marker):
                                            s_result = self.node_split + i_result[1] + i_result[1][1:-1] + i_result[
                                                1] + self.node_split
                                title = '{}{}{}{}{}{}'.format(title_long_x + swap,
                                                              self.node_split,
                                                              all_mar[1] + topics.get('title') + all_mar[
                                                                  1] + self.node_split,
                                                              # 备注
                                                              self.note[1] + str(s_note) + self.note[
                                                                  1] + self.node_split,
                                                              # 批注
                                                              self.comment[1] + str(s_comment) + self.comment[
                                                                  1] + self.node_split,
                                                              # 执行结果（旗子）
                                                              s_result)
                                swap = self.node_split + all_mar[1] + topics.get('title') + all_mar[
                                    1] + self.node_split + self.note[
                                           1] + str(s_note) + self.note[1] + self.node_split + self.comment[
                                           1] + str(s_comment) + self.comment[1] + s_result
                            # 非用例节点
                            else:
                                title = '{}{}{}'.format(title_long_x + swap, self.node_split,
                                                        all_mar[1] + topics.get('title') + all_mar[1] + self.node_split)
                                swap = self.node_split + all_mar[1] + topics.get('title') + all_mar[1]
                            break
                        else:
                            # 循环到最后 没匹配到（说明此标志没有在预定标志中）
                            if markers[len(markers) - 1] == marker \
                                    and self.all_markers[len(self.all_markers) - 1] == all_mar:
                                # TODO 特殊场景不用考虑 不影响使用（暂无解决方案） 用例执行结果（通过、未通过）
                                if markers[len(markers) - 1] == self.result_pass[0] or markers[len(markers) - 1] == \
                                        self.result_no_pass[0]: continue
                                print('节点：【{}】，在预定标志中未匹配到此标志（{}）'.format(topics.get('title'), marker))
            else:
                title = '{}{}{}'.format(title_long_x, self.node_split, topics.get('title'))
            # 取节点值
            if topics.get('topics') is not None:
                JarXMindUtil().__base(title, topics, result_data_all)
            else:
                result_data_all.append(title)
                continue

    @staticmethod
    def calc_progress(excel_file):
        """
        计算用例执行进度（sheet表：0 列：0）
        :param excel_file: 解析后的excel文件
        :return: Ture/Fase,message
        """
        calc_sheet_number = 0
        calc_col_number = 0

        wd = xlrd.open_workbook(excel_file, formatting_info=True)
        wd_new = copy(wd)
        sheet = wd.sheet_by_index(calc_sheet_number)

        row_count = sheet.nrows
        row_count_ture = row_count - 1
        ture_count = 0  # 已经执行的个数
        for i_row in range(row_count):
            if i_row == row_count - 1: break
            if str(sheet.row_values(i_row + 1)[0]) != '':
                ture_count += 1
        now_progress = str(round((ture_count / row_count_ture) * 100)) + '%' + '({}/{})'.format(ture_count,
                                                                                                row_count_ture)
        print(now_progress)
        # 进度写入表格中
        progress_value = str(sheet.row_values(0)[0])
        progress_value += now_progress

        ali_horiz = 'align: horiz center'  # 水平居中
        ali_vert = 'align: vert center'  # 垂直居中
        fore_colour = 'pattern: pattern solid,fore_colour pale_blue'  # 设置单元格背景色为pale_blue色

        # 表头格式（垂直+水平居中、表头背景色）
        style_header = xlwt.easyxf('{};{};{}'.format(fore_colour, ali_horiz, ali_vert))

        sheet_new = wd_new.get_sheet(calc_sheet_number)
        sheet_new.write(0, 0, progress_value, style_header)
        wd_new.save(excel_file)
        return now_progress

    @staticmethod
    def check_excel(excel_file, check_col: list):
        """
        校验解析XMind文件后的excel文件正确性（指定列的单元格不能为空）
        :param excel_file: 解析后的excel文件
        :param check_col: 校验列名称（列表）
        :return: Ture/False,message
        """
        wd = xlrd.open_workbook(excel_file)
        sheet = wd.sheet_by_index(0)

        def get_excel_all_col():
            """
            获取表头
            :return:
            """
            col_values = sheet.row_values(0)
            return col_values

        col_values: list = get_excel_all_col()
        print('表头：', col_values)

        # step1 校验文件中是否有check_col中的列
        null_col = []
        for i in check_col:
            if col_values.count(i) <= 0:
                null_col.append(i)
        if len(null_col) != 0:
            return False, 'Excel表头中无此字段：{}'.format(null_col)

        # step2 校验每一行中指定列是否为空
        null_value = []
        col_z = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                 'U', 'V', 'W', 'X', 'Y', 'Z', ]
        row_count = sheet.nrows
        for i in range(row_count):  # 行
            for i_col in check_col:  # 列
                col_index = col_values.index(i_col)
                value = sheet.row_values(i)[col_index]
                # 单元格为空
                if value == '':
                    null_value.append(str(col_z[col_index]) + str(i + 1))
        if len(null_value) != 0:
            return False, '单元格无值：{}'.format(null_value)
        return True, '校验通过'
