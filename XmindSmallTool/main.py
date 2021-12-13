import os
import PySimpleGUI as sg
from config import ConfigParser, BASE_DIR
from case2excel import Case2Excel
from parsercases import ParserCases
from excel import Excel2Xmind

base_dir = os.path.dirname(os.path.abspath(__file__))
default_template = os.path.join(BASE_DIR, 'template.xlsx')

#切换窗口颜色
# themelist = sg.theme_list()
# print(themelist)
# sg.preview_all_look_and_feel_themes()
sg.theme('LightBlue')

xmind_layout = [
    [sg.In(key='_TEMPLATE_', default_text=default_template), sg.FileBrowse('选择模板')],
    [sg.In(key='_XMIND_FILE_', enable_events=True),
     sg.FileBrowse('选择Xmind文件', key='_SELECT_XMIND_', file_types=(('Xmind', '*.xmind'),), initial_folder=base_dir)],
    [sg.In(key='_EXCEL_RESULT_', enable_events=True),
     sg.FileSaveAs('用例保存位置', file_types=(('Excel', '*.xlsx'),) )],
    [sg.B('转化', key='_XMIND_TRANSLATE_')],
]

excel_layout = [
    [sg.In(key='_EXCEL_FILE_', enable_events=True),
     sg.FileBrowse('选择Excel文件', key='_SELECT_EXCEL_', file_types=(('Excel', '*.xlsx'),), initial_folder=base_dir)],
    [sg.In(key='_XMIND_RESULT_', enable_events=True),
     sg.FileSaveAs('脑图保存位置', file_types=(('Xmind', '*.xmind'),))],
    [sg.B('重构', key='_EXCEL_TRANSLATE_')],
]

layout = [
    [sg.Frame('脑图转用例',xmind_layout)],
    [sg.T()],
    [sg.Frame('用例转脑图',excel_layout)]
]

window = sg.Window('Excel2Xmind',layout=layout,finalize=True).finalize()

while True:
    event,value = window.read()
    if event == '_TEMPLATE_':
        event, value = window.Read()

    # 将生成后的路径与xmind文件在同一路径，并将后缀名修改为xlsx
    if event == '_XMIND_FILE_':
        window['_EXCEL_RESULT_'].update(value=os.path.splitext(value['_XMIND_FILE_'])[0])
        event,value = window.Read()

    # 由于选择excel保存路径时可能会忘记输入后缀名，自动加上.xlsx后缀
    if event == '_EXCEL_RESULT_':
        excel_result = value['_EXCEL_RESULT_']
        if os.path.splitext(excel_result)[-1] not in ['xlsx', 'xls']:
            excel_result = excel_result + '.xlsx'
            window['_EXCEL_RESULT_'].update(value=excel_result)
        event, value = window.Read()

    # 当点击转化按钮时，获取文件路径，并调用核心代码进行转化
    if event == '_XMIND_TRANSLATE_':
        template = value['_TEMPLATE_']
        xmind_file = value['_XMIND_FILE_']
        excel_result = value['_EXCEL_RESULT_'] + '.xlsx'

        if not template:
            sg.popup('请选择模板文件!')

        elif not xmind_file or not xmind_file.endswith('.xmind'):
            sg.popup('您没有选择xmind文件或文件格式有误!')

        else:
            if not excel_result:
                sg.popup('您没有选择用例保存位置，将默认保存在xmind所在的文件夹!')
                excel_result = os.path.splitext(xmind_file)[0] + '.xlsx'
                event, value = window.Read()
            xc = ParserCases(xmind_file)
            if xc.msg:
                sg.popup('选择的xmind格式有问题，请检查文件！')
                event, value = window.Read()
            ce = Case2Excel(template, excel_result)
            # try:
            #     ce.write_case_to_excel(xc.all_map_case)
            # except Exception as e:
            #     sg.popup('转化出现错误：\n' + str(e))
            # else:
            #     sg.popup('用例生成成功，请前往 {} 查看！'.format(excel_result))
            ce.write_case_to_excel(xc.all_map_case)

        event, value = window.Read()

    # 将生成后的路径与excel文件在同一路径，并将后缀名修改为xmind
    if event == '_EXCEL_FILE_':
        window['_XMIND_RESULT_'].update(value=os.path.splitext(value['_EXCEL_FILE_'])[0])
        event, value = window.Read()

    # 由于选择xmind保存路径时可能会忘记输入后缀名，自动加上.xmind后缀
    if event == '_XMIND_RESULT_':
        case_file = value['_XMIND_RESULT_']
        # if os.path.splitext(case_file)[-1] not in ['xmind']:
        #     case_file = case_file + '.xmind'
        #     window['_XMIND_RESULT_'].update(value=case_file)
        event, value = window.Read()

    # 当点击转化按钮时，获取文件路径，并调用核心代码进行转化
    if event == '_EXCEL_TRANSLATE_':
        xmind_file = value['_EXCEL_FILE_']
        case_file = value['_XMIND_RESULT_'] + '.xmind'

        if not xmind_file or not xmind_file.endswith('.xlsx'):
            sg.popup('您没有选择xlsx文件或文件格式有误!')

        else:
            if not case_file:
                sg.popup('您没有选择脑图保存位置，将默认保存在excel所在的文件夹!')
                case_file = os.path.splitext(xmind_file)[0] + '.xmind'
                event, value = window.Read()
            xc = Excel2Xmind()
            dict_Data = xc.load_excel(xmind_file)
            print(dict_Data)

            try:
                xc.design_sheet(dict_Data,case_file)
            except Exception as e:
                sg.popup('转化出现错误：\n' + str(e))
            else:
                sg.popup('用例生成成功，请前往 {} 查看！'.format(case_file))
                # event, value = window.Read()

        event, value = window.Read()

    if event is None:
        break


window.close()
