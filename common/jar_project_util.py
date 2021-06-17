#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@Time    ：2020/9/23 18:16
@Author  ：维斯
@File    ：jar_project_util.py
@Version ：1.0
@Function：
"""
# TODO: 检查完毕 可上传

import os


class JarProjectUtil:
    @staticmethod
    def project_root_path(project_name=None, print_log=True):
        """
        获取当前项目根路径
        :param project_name: 项目名称
                                1、可在调用时指定
                                2、[推荐]也可在此方法中直接指定 将'XmindUitl-master'替换为当前项目名称即可（调用时即可直接调用 不用给参数）
        :param print_log: 是否打印日志信息
        :return: 指定项目的根路径
        """
        p_name = 'XmindUitl-master' if project_name is None else project_name
        project_path = os.path.abspath(os.path.dirname(__file__))
        # Windows
        if project_path.find('\\') != -1: separator = '\\'
        # Mac、Linux、Unix
        if project_path.find('/') != -1: separator = '/'

        root_path = project_path[:project_path.find(f'{p_name}{separator}') + len(f'{p_name}{separator}')]
        if print_log: print(f'当前项目名称：{p_name}\r\n当前项目根路径：{root_path}')
        return root_path

    @staticmethod
    def del_endswith_none(str1: str):
        """
        删除字符串首尾的空字符（空格、换行）
        :param str1:
        """
        s = str1
        if str1 is not None:
            while True:
                if s.endswith(' ') or s.endswith('	') or s.endswith('\r\n') or s.endswith('\r') or s.endswith('\n'):
                    s = s[:-1]
                elif s.startswith(' ') or s.startswith('	') or s.startswith('\r\n') or s.startswith(
                        '\r') or s.startswith('\n'):
                    s = s[1:]
                else:
                    break
        return s


if __name__ == '__main__':
    JarProjectUtil.project_root_path()
