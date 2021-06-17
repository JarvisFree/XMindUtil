#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@Time    ：2021/6/17 16:28
@Author  ：维斯
@File    ：start.py
@Version ：1.0
@Function：
"""

from model.jar_xmind_util import JarXMindUtil

if __name__ == '__main__':
    ju = JarXMindUtil()
    # 样例
    xmind_file = '../data/xxx项目测试用例.xmind'  # 编写好的XMind用例文件
    out_excel_file = '../data/xxx项目测试用例.xls'  # 解析后生成的Excel文件

    # 用例编号前缀
    ju.case_number_model = 'CASE_TEST_210617'

    # step1 解析XMind
    result = ju.analysis(xmind_file)
    # step2 写入excel
    ju.to_excel(out_excel_file, result)

    check_col = ['用例编号', '用例类型', '重要级别', '所属模块', '验证点', '用例标题']
    success, error = ju.check_excel(out_excel_file, check_col)
    if success:
        print('校验成功')
    else:
        print(error)
    ju.calc_progress(out_excel_file)
