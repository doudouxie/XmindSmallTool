#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# @Time：2020/5/30 16:35
# @Email: am1122@163.com
# @Author: 'Nemo'
import os,sys
import yaml

# 通过pyinstaller打包后，不能使用__file__读取文件路径，因为该文件会被存放在临时文件夹中
# 可以使用以下几种方式读取运行后的exe所在路径
'''
print(os.path.realpath(sys.argv[0]))
print(os.path.realpath(sys.executable))
print(os.path.dirname(os.path.realpath(sys.argv[0])))
print(os.path.dirname(os.path.realpath(sys.executable)))
'''

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.realpath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG = os.path.join(BASE_DIR, 'conf.yaml')

class ConfigParser:

    @classmethod
    def get_config(cls, option=None, section=None):
        '''
        section = None, 表示没有分块，直接取 option
        option = None, 表示取所有的配置
        :param option: 配置项的 key
        :param section: 配置分块
        :return:
        '''
        config_content = yaml.safe_load(open(CONFIG, encoding='utf-8'))
        if section:
            section = config_content[section]
            return section if not option else section[option]
        else:
            if option:
                return config_content[option]
            else:
                return config_content

    @classmethod
    def set_config(cls, option, value, section=None):
        '''
        section = None, 直接设置配置项
        :param option: 配置项的 key
        :param value: 配置项的值
        :param section: 配置分块
        :return:
        '''
        config_content = yaml.safe_load(open(CONFIG, encoding='utf-8'))
        if section:
            config_content[section][option] = value
        else:
            config_content[option] = value
        yaml.dump(config_content, open(CONFIG, 'w', encoding='utf-8'), allow_unicode=True)


if __name__ == '__main__':
    print(ConfigParser.get_config('mark', 'cases'))
    print(ConfigParser.set_config('mark', 'new', 'cases'))