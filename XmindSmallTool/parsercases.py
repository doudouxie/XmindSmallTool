#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# @Time：2020/5/30 16:35
# @Email: am1122@163.com
# @Author: 'Nemo'

from xmindparser import xmind_to_dict
import re
from config import ConfigParser


class ParserCases:

    def __init__(self, path):
        self.msg = ''
        try:
            self.xmind = xmind_to_dict(path)
        except Exception as e:
            self.msg = e

        self.test_cases = []
        self.m = ConfigParser.get_config('mark', 'cases')    # 读取配置中的标识

    def get_case_by_default(self, map):
        """
        以默认5级的方式处理用例
        :param map:
        :return:
        """
        # 第一级为 excel 的 sheet
        map_name = map['title']

        # 模块的描述  module['note']  # 暂时不用
        test_cases = []
        # 页面
        for module in map['topic']['topics']:

            module_name = module['title']
            # page_note = page['note']
            points = module.get('topics')

            if points:
                # 功能点
                for point in points:
                    # 功能点
                    point_name = point.get('title')
                    # 功能点下的用例
                    cases = point.get('topics')

                    if cases:
                        # 用例
                        for case in cases:
                            test_case = self.parse_case(
                                case=case,
                                module=module_name,
                                point_name=point_name,
                                case_name=case['title']
                            )
                            test_cases.append(test_case)
        return {
            'test_cases': test_cases,
            'map_name': map_name
        }

    def is_test_case(self, node):
        '''
        判断该层级是否是测试用例
        可配置：
            1. 默认方式，固定N级，第一级为模块，最后三级分别为用例的标题、步骤、结果，标题包含优先级和前置条件（笔记）
            2. 指定某个关键字，比如 tc：,case: 等，如果没设置默认使用优先级图标，如果都没有则忽略该节点
            3. 使用英文单词区分用例及其步骤，cases 的所有子节点都是用例节点，如果没有cases节点则忽略
        :return:
        '''
        priority = '' if not node.get('makers') else node.get('makers')[0][-1]

        if self.m == 1:
            pass
        elif self.m == 2:
            case_title = node.get('title')
            key  = ConfigParser.get_config('mark_2_string', 'cases')
            if case_title.lower().startswith(key) or priority:
                if case_title.lower().startswith(key):
                    case_title = re.sub(key+r'[:：]', '', case_title, flags=re.I)
                return case_title  # 返回处理后的用例标题
        elif self.m == 3:
            title = node.get('title')
            if title == ConfigParser.get_config('mark_3_key', 'cases'):
                return node.get('topics')  # 返回所有测试用例子节点
        else:
            return None

    @property
    def all_map_case(self):
        '''
        提取整个xmind中所有map的测试用例
        :return:
        '''
        if self.m == 1:
            return [self.get_case_by_default(map) for map in self.xmind]
        else:
            return [self.parse_map(map) for map in self.xmind]

    def find_cases(self, node, module, p=''):
        # 往下遍历节点，遇到用例节点时开始处理用例，并且把前面的节点组合成子模块(子模块组合有问题，暂不考虑)
        child = node.get('topics')
        if p == module:
            p = ''
        point = p + '$' + node.get('title') if p else node.get('title')

        if child:
            for n in child:
                if self.m == 2:
                    case_name = self.is_test_case(n)
                    if case_name:
                        test_case = self.parse_case(n, module, point, case_name)
                        # test_case['test_title'] = case_name
                        self.test_cases.append(test_case)
                    else:
                        self.find_cases(n, module, point)  # 不是用例，继续向下查找
                elif self.m == 3:
                    cases = self.is_test_case(n)
                    if cases:
                        for case in cases:
                            case_name = case.get('title')
                            test_case = self.parse_case(case, module, point, case_name)
                            test_case['test_title'] = case_name
                            self.test_cases.append(test_case)
                    else:
                        self.find_cases(n, module, point)   # 不是用例，继续向下查找

    # 画布
    def parse_map(self, map):
        # 第一级为 excel 的 sheet
        map_name = map['title']
        # 整个画布中的大主题，相当于模块
        modules = map['topic'].get('topics')

        map_cases = []

        if modules:
            for m in modules:
                self.find_cases(m, m.get('title', '未命名模块'))
                map_cases += self.test_cases    # 将map下的所有用例存储到一个列表中
                self.test_cases = []    # 清空 test_cases 中储存的用例，用于存储下一个画布的用例

        return {
            'map_name': map_name,
            'test_cases': map_cases
        }

    def parse_case(self, case, module, point_name, case_name):
        """
        解析测试用例
        :param case:
        :return:
        """
        preset = case.get('note', '')  # 使用笔记作为前置条件

        # 标注 callout：作为用例类别是功能测试、接口测试或其他测试
        '''
        功能测试
        性能测试
        配置相关
        安装部署
        安全相关
        接口测试
        其他
        '''
        category = case.get('callout', ['功能测试'])[0]
        # 标签 labels：作为用例的执行阶段，是冒烟用例、执行用例或者其他用例
        '''
        单元测试阶段
        功能测试阶段
        集成测试阶段
        系统测试阶段
        冒烟测试阶段
        版本验证阶段
        '''
        stage = case.get('labels', ['功能测试阶段'])[0]

        # 获取测试用例所在的标记，用于确定优先级和是否冒烟测试
        priority = 3
        makers = case.get('makers') or ''
        if makers:
            for marker in makers:
                if 'priority' in marker:
                    priority = marker[-1]

            # 是否冒烟测试用例
        is_smoke = '是' if 'flag-red' in makers else ''

        steps = self.parse_case_step(case.get('topics'), case_name)

        return {
            'test_title': case_name,
            'module': module,
            'point': point_name,
            'preset': preset,
            'priority': priority,
            'category': category,
            'stage': stage,
            'is_smoke': is_smoke,
            'test_steps': steps['test_steps'],
            'test_results': steps['test_results']
        }

    def parse_case_step(self, steps, default=None):
        """
        解析测试步骤
        :param steps:
        :return:
        """
        test_steps = []
        test_results = []
        if steps:
            for step in steps:
                step_name = step['title']
                test_steps.append(step_name)

                results = step.get('topics')

                # 预期结果
                if results:
                    for result in results:
                        result = result['title'].splitlines()
                        test_results += result
            test_results = '\n'.join(['{}. {}'.format(i + 1, r) for i, r in enumerate(test_results)])
            test_steps = '\n'.join(['{}. {}'.format(i + 1, r) for i, r in enumerate(test_steps)])
        else:
            # 缺少步骤和结果，则直接使用用例标题填充
            if ConfigParser.get_config('no_step_or_result_fill_by_case_title', 'cases'):
                test_steps = '1. {}'.format(default)
                test_results = '1. {}'.format(default)
            else:
                test_steps = ''
                test_results = ''
            self.msg = 'Xmind中存在用例缺少步骤或结果，将直接使用测试标题填充'

        return {'test_steps': test_steps, 'test_results': test_results}


if __name__ == '__main__':
    from xmindparser import xmind_to_dict
    from pprint import pprint

    d = xmind_to_dict('../test/test.xmind')
    pprint(d)

    p = ParserCases('../test/test.xmind')
    pprint(p.all_map_case)