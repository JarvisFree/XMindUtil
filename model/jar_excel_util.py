#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@Time    ：2020/8/20 21:02
@Author  ：维斯
@File    ：jar_excel_util.py
@Version ：1.0
@Function：Excel工具
"""

from typing import List

import xlwt


class JarExcelUtil:
    def __init__(self, header_list: List[list]):
        """
        :param header_list: 如下格式
            例1：默认列宽
            header_list = [
                ['序号'],    # 表格第0列[此列表头名称]
                ['姓名'],
                ['性别'],
                ['爱好'],
                ['生日']
            ]
            例2：自定义列宽（列宽值为int类型 英文字符长度 如：10 表示列宽为10个英文字符长度）
            header = [
                ['序号', 5],  # 表格第0列[此列表头名称,列宽]
                ['姓名', 10], # 表格第1列[此列表头名称,列宽]
                ['性别', 10],
                ['爱好', 10],
                ['生日', 20]
            ]
        """
        self.data = header_list
        self.__color_str = 'aqua 0x31\r\n\
black 0x08\r\n\
blue 0x0C\r\n\
blue_gray 0x36\r\n\
bright_green 0x0B\r\n\
brown 0x3C\r\n\
coral 0x1D\r\n\
cyan_ega 0x0F\r\n\
dark_blue 0x12\r\n\
dark_blue_ega 0x12\r\n\
dark_green 0x3A\r\n\
dark_green_ega 0x11\r\n\
dark_purple 0x1C\r\n\
dark_red 0x10\r\n\
dark_red_ega 0x10\r\n\
dark_teal 0x38\r\n\
dark_yellow 0x13\r\n\
gold 0x33\r\n\
gray_ega 0x17\r\n\
gray25 0x16\r\n\
gray40 0x37\r\n\
gray50 0x17\r\n\
gray80 0x3F\r\n\
green 0x11\r\n\
ice_blue 0x1F\r\n\
indigo 0x3E\r\n\
ivory 0x1A\r\n\
lavender 0x2E\r\n\
light_blue 0x30\r\n\
light_green 0x2A\r\n\
light_orange 0x34\r\n\
light_turquoise 0x29\r\n\
light_yellow 0x2B\r\n\
lime 0x32\r\n\
magenta_ega 0x0E\r\n\
ocean_blue 0x1E\r\n\
olive_ega 0x13\r\n\
olive_green 0x3B\r\n\
orange 0x35\r\n\
pale_blue 0x2C\r\n\
periwinkle 0x18\r\n\
pink 0x0E\r\n\
plum 0x3D\r\n\
purple_ega 0x14\r\n\
red 0x0A\r\n\
rose 0x2D\r\n\
sea_green 0x39\r\n\
silver_ega 0x16\r\n\
sky_blue 0x28\r\n\
tan 0x2F\r\n\
teal 0x15\r\n\
teal_ega 0x15\r\n\
turquoise 0x0F\r\n\
violet 0x14\r\n\
white 0x09\r\n\
yellow 0x0D'
        self.color_list = []  # [[]]   [['aqua', '0x31'], ['black', '0x08'], ...]
        for color in self.__color_str.split('\r\n'):
            color = color.split(' ')
            self.color_list.append(color)

    def write(self, out_file, data_body: List[list], sheet_name='sheet', frozen_row: int = 1, frozen_col: int = 0):
        """
        写入数据
        :param out_file: 保存文件（如：test.xlsx）
        :param data_body: data_body[0]为表格第0行数据  data_body[0][0]为表格第0行第0列单元格值
        :param sheet_name:
        :param frozen_row: 冻结行（默认首行）
        :param frozen_col: 冻结列（默认不冻结）
        """
        # step1 判断数据正确性（每行列数是否与表头相同）
        count = 0
        for pro in data_body:
            if len(pro) != len(self.data):
                raise Exception(
                    'data_body数据错误 第{}行（从0开始） 需为{}个元素 当前行{}个元素：{}'.format(count, len(self.data), len(pro), str(pro)))
            count += 1

        # step2 写入数据
        wd = xlwt.Workbook()
        sheet = wd.add_sheet(sheet_name)

        ali_horiz = 'align: horiz center'  # 水平居中
        ali_vert = 'align: vert center'  # 垂直居中
        fore_colour = 'pattern: pattern solid,fore_colour pale_blue'  # 设置单元格背景色为pale_blue色

        # 表头格式（垂直+水平居中、表头背景色）
        style_header = xlwt.easyxf('{};{};{}'.format(fore_colour, ali_horiz, ali_vert))

        # 表体格式（垂直居中）
        style_body = xlwt.easyxf('{}'.format(ali_vert))

        # 表头
        for col in self.data:
            # 默认列宽
            if len(col) == 1:
                sheet.write(0, self.data.index(col), str(col[0]), style_header)
            # 自定义列宽
            if len(col) == 2:
                sheet.write(0, self.data.index(col), str(col[0]), style_header)
                # 设置列宽
                sheet.col(self.data.index(col)).width = 256 * col[1]  # 256为基数 * n个英文字符
        # 行高（第0行）
        sheet.row(0).height_mismatch = True
        sheet.row(0).height = 20 * 20  # 20为基数 * 20榜

        # 表体
        index = 1
        for pro in data_body:
            sheet.row(index).height_mismatch = True
            sheet.row(index).height = 20 * 20  # 20为基数 * 20榜
            for d in self.data:
                value = pro[self.data.index(d)]
                # 若值类型是int、float 直接写入 反之 转成字符串写入
                if type(value) == int or type(value) == float:
                    sheet.write(index, self.data.index(d), value, style_body)
                else:
                    sheet.write(index, self.data.index(d), str(value), style_body)
            index += 1
        # 冻结（列与行）
        sheet.set_panes_frozen('1')
        sheet.set_horz_split_pos(frozen_row)  # 冻结前n行
        sheet.set_vert_split_pos(frozen_col)  # 冻结前n列

        wd.save(out_file)

    def color_test(self):
        """
        测试颜色
        """
        body_t = []
        for color in self.color_list:
            print(color)
            body_t.append(color)
        wd = xlwt.Workbook()
        sheet = wd.add_sheet('sheet')

        index = 0
        for b in body_t:
            ali = 'align: horiz center;align: vert center'  # 垂直居中 水平居中
            fore_colour = 'pattern: pattern solid,fore_colour {}'.format(
                self.color_list[index][0])  # 设置单元格背景色为pale_blue色
            style_header = xlwt.easyxf(
                '{};{}'.format(fore_colour, ali))
            sheet.write(index, 0, str(b), style_header)
            sheet.col(0).width = 256 * 150  # 256为基数 * n个英文字符
            index += 1

        wd.save('颜色测试.xlsx')


# 测试颜色
# if __name__ == '__main__':
#     header_t = [
#         ['颜色']
#     ]
#     JarExcelUtil(header_t).color_test()

if __name__ == '__main__':
    header = [
        ['序号', 5],
        ['姓名', 10],
        ['性别', 10],
        ['爱好', 10],
        ['生日', 20]
    ]

    # header = [
    #     ['序号'],
    #     ['姓名'],
    #     ['性别'],
    #     ['爱好'],
    #     ['生日']
    # ]

    body = [
        [1, '张三', '男', '篮球', '1994-07-23'],
        [2, '李四', '女', '足球', '1994-04-03'],
        [3, '王五', '男', '兵乓球', '1994-09-13']
    ]

    JarExcelUtil(header_list=header).write(out_file='测试.xlsx', data_body=body)
