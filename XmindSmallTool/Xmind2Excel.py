from parsercases import ParserCases
import PySimpleGUI as sg
from case2excel import Case2Excel
import os
from config import ConfigParser, BASE_DIR

base_dir = os.path.dirname(os.path.abspath(__file__))

default_template = os.path.join(BASE_DIR, 'template.xlsx')

# 模板选择
# 选择xmind文件
# 选择存放位置，默认与 xmind 在同一目录，且后缀名改为 xmind
TIPS = {
    'save_file': '如果不选择则会直接保存在xmind所在目录',
    #'r1': '方案1：默认为5级，第一级为模块，最后三级分别为用例标题、步骤和结果',
    #'r2': '方案2：使用定义的关键字识别用例，比如"tc:登录失败"，\n\t使用tc作为标识，会识别该节点为测试用例\n\t后面分别为用例、步骤和结果',
    'r3': '方案3：使用定义的关键字作为用例的前置，比如"cases"，\n\t则认为该节点下的所有节点都是用例',
    'trans': '自动保存为 .xlsx 文件',
    'zentao': '自动保存为 .csv 文件',
    'encoding': '如果CSV出现乱码，则修改编码格式，一般为utf-8或GBK，一种不行就换一种咯！',
    'template_text': '通过这里的关键字去识别模板中的用例，比如会根据“标题”去寻找用例标题',
}

# [sg.Radio('My first Radio!', "RADIO1", default=True),
#     sg.Radio('My second radio!', "RADIO1")]
CASE_CONFIG = ConfigParser.get_config(section='cases')
FIELD = ConfigParser.get_config('field', 'template')

case_column = [
    # [sg.T('用例识别方式：')],
    #[sg.Radio('默认方案', group_id='case_mark', key='_MARK1_', tooltip=TIPS.get('r1'), enable_events=True)],
    # [
    #     sg.Radio('关键字识别', group_id='case_mark', key='_MARK2_', tooltip=TIPS.get('r2'), enable_events=True),
    #     sg.In(key='_MARK2STR_', size=(10,))
    # ],
    [
        sg.Radio('前置关键字', group_id='case_mark', key='_MARK3_', tooltip=TIPS.get('r3'), enable_events=True),
        sg.In(key='_MARK3STR_', size=(10,))
    ],
]

# field_column = [
#     [sg.T('模板字段设置：', tooltip=TIPS.get('template_text'))],
#     [sg.T('所属模块'), sg.I(default_text=FIELD.get('module'), key='_FIELD_MODULE_')],
#     [sg.T('用例标题'), sg.I(default_text=FIELD.get('test_title'), key='_FIELD_TITLE_')],
#     [sg.T('前置条件'), sg.I(default_text=FIELD.get('preset'), key='_FIELD_PRESET_')],
#     [sg.T('操作步骤'), sg.I(default_text=FIELD.get('test_step'), key='_FIELD_STEP_')],
#     [sg.T('预期结果'), sg.I(default_text=FIELD.get('test_result'), key='_FIELD_RESULT_')],
#     [sg.T('优先级'), sg.I(default_text=FIELD.get('priority'), key='_FIELD_PRIORITY_')],
#     [sg.T('用例类型'), sg.I(default_text=FIELD.get('category'), key='_FIELD_CATEGORY_')],
#     [sg.T('适用阶段'), sg.I(default_text=FIELD.get('stage'), key='_FIELD_STAGE_')],
# ]

config_frame = [
    # [sg.Column(case_column), sg.VerticalSeparator(), sg.Column(field_column)],
    # [sg.Frame('用例识别方式', case_column)],
    [sg.T('_'*60)],
    [sg.T('当没有写步骤或预期结果时，是否直接使用用例标题填充'), sg.Checkbox('', default=False, key='_FILL_')],
    [sg.T('用例字体大小：'), sg.In(key='_FONT_SIZE_', default_text=ConfigParser.get_config('font_size', 'style'), size=(5,))],
    [sg.T('csv 编码：', tooltip=TIPS.get('encoding')), sg.In(key='_CODING_', size=(5, ), default_text=ConfigParser.get_config('encoding', 'sys'))],
    [sg.B('保存配置', key='_SAVE_CONFIG_')]
]

layout = [
    [sg.Frame('配置', case_column + config_frame)],
    [sg.T()],
    [sg.In(key='_TEMPLATE_', default_text=default_template), sg.FileBrowse('选择模板')],
    [sg.In(key='_XMIND_FILE_', enable_events=True), sg.FileBrowse('选择Xmind文件', key='_SELECT_XMIND_', file_types=(('Xmind', '*.xmind'),), initial_folder=base_dir)],
    [sg.In(key='_CASE_FILE_', enable_events=True), sg.FileSaveAs('选择用例保存位置', file_types=(('Excel', '*.xlsx'),), tooltip=TIPS.get('save_file'))],
    [sg.B('转化', key='_TRANSLATE_', tooltip=TIPS.get('trans'))],
    [sg.In(key='_EXCEL2XMIND_FILE_', enable_events=True), sg.FileBrowse('选择Excel文件', key='_SELECT_EXCEL_', file_types=(('Excel', '*.xlsx'),), initial_folder=base_dir)],
    [sg.In(key='_EXCEL2XMINDCASE_FILE_', enable_events=True), sg.FileSaveAs('选择脑图保存位置', file_types=(('Xmind', '*.xmind'),), tooltip=TIPS.get('save_file'))],
    [sg.B('重构', key='_ROLLBACK_', tooltip=TIPS.get('backs'))]
]

window = sg.Window('xmind2excel用例转化', layout=layout, finalize=True).finalize()

while True:
    # 处理配置
    if CASE_CONFIG['mark'] == 1:
        window['_MARK1_'].update(value=True)
    elif CASE_CONFIG['mark'] == 2:
        window['_MARK2_'].update(value=True)
        try:
            window['_MARK3STR_'].update(value='', disabled=True)
            window['_MARK2STR_'].update(value=CASE_CONFIG.get('mark_2_string'))
        except:
            pass
    elif CASE_CONFIG['mark'] == 3:
        window['_MARK3_'].update(value=True)
        try:
            #window['_MARK2STR_'].update(value='', disabled=True)
            window['_MARK3STR_'].update(value=CASE_CONFIG.get('mark_3_key'))
        except:
            pass
    if CASE_CONFIG['no_step_or_result_fill_by_case_title']:
        window['_FILL_'].update(value=True)

    # 读取窗口事件
    event, value = window.Read()

    if event == '_MARK1_':
        window['_MARK2STR_'].update(value='', disabled=True)
        window['_MARK3STR_'].update(value='', disabled=True)
        event, value = window.Read()

    elif event == '_MARK2_':
        window['_MARK2_'].update(value=True)
        window['_MARK2STR_'].update(disabled=False, value=CASE_CONFIG.get('mark_2_string'))
        window['_MARK3STR_'].update(value='', disabled=True)
        event, value = window.Read()

    elif event == '_MARK3_':
        window['_MARK3_'].update(value=True)
        window['_MARK3STR_'].update(disabled=False, value=CASE_CONFIG.get('mark_3_key'))
        window['_MARK2STR_'].update(value='', disabled=True)
        event, value = window.Read()

    if event == '_TEMPLATE_':
        event, value = window.Read()

    # 配置保存
    if event == '_SAVE_CONFIG_':
        try:
            if value['_MARK1_']:
                # 写入配置
                ConfigParser.set_config('mark', 1, 'cases')
            elif value['_MARK2_']:
                ConfigParser.set_config('mark', 2, 'cases')
                ConfigParser.set_config('mark_2_string', value['_MARK2STR_'], 'cases')
            elif value['_MARK3_']:
                ConfigParser.set_config('mark', 3, 'cases')
            ConfigParser.set_config('no_step_or_result_fill_by_case_title', value['_FILL_'], 'cases')
            ConfigParser.set_config('font_size', int(value['_FONT_SIZE_']), 'style')
            ConfigParser.set_config('encoding', value['_CODING_'], 'sys')
        except Exception as e:
            sg.popup(e)
        else:
            sg.popup('配置保存成功！')
        event, value =  window.Read()

    # 将生成后的路径与xmind文件在同一路径，并将后缀名修改为xlsx
    if event == '_XMIND_FILE_':
        window['_CASE_FILE_'].update(value=os.path.splitext(value['_XMIND_FILE_'])[0])
        event, value = window.Read()

    # 由于选择excel保存路径时可能会忘记输入后缀名，自动加上.xlsx后缀
    if event == '_CASE_FILE_':
        case_file = value['_CASE_FILE_']
        if os.path.splitext(case_file)[-1] not in ['xlsx', 'xls']:
            case_file = case_file + '.xlsx'
            window['_CASE_FILE_'].update(value=case_file)
        event, value = window.Read()

    # 当点击转化按钮时，获取文件路径，并调用核心代码进行转化
    if event == '_TRANSLATE_':
        template = value['_TEMPLATE_']
        xmind_file = value['_XMIND_FILE_']
        case_file = value['_CASE_FILE_'] + '.xlsx'

        if not template:
            sg.popup('请选择模板文件!')

        elif not xmind_file or not xmind_file.endswith('.xmind'):
            sg.popup('您没有选择xmind文件或文件格式有误!')

        else:
            if not case_file:
                sg.popup('您没有选择用例保存位置，将默认保存在xmind所在的文件夹!')
                case_file = os.path.splitext(xmind_file)[0] + '.xlsx'
                event, value = window.Read()
            xc = ParserCases(xmind_file)
            if xc.msg:
                sg.popup('选择的xmind格式有问题，请检查文件！')
                event, value = window.Read()
            ce = Case2Excel(template, case_file)
            try:
                ce.write_case_to_excel(xc.all_map_case)
            except Exception as e:
                sg.popup('转化出现错误：\n' + str(e))
            else:
                sg.popup('用例生成成功，请前往 {} 查看！'.format(case_file))
                # event, value = window.Read()
        event, value = window.Read()

    if event is None:
        break

window.close()