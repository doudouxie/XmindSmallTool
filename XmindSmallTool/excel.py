import re
import openpyxl
import xmind
from xmind.core.markerref import MarkerId

def design_sheet(dicts):
    workbook = xmind.load('test1.xmind')
    sheet = workbook.getPrimarySheet()
    sheet.setTitle('First Sheet')
    root_topic = sheet.getRootTopic()
    root_topic.setTitle('root node')
    dict_item(dicts,root_topic)
    xmind.save(workbook)
    print(sheet.getData())

def dict_item(dicts,topic):
    for key in dicts:
        subtopic = topic.addSubTopic()
        subtopic.setTitle(key)
        subtopic.addMarker(MarkerId.priority1)
        subtopic.setPlainNotes('note')
        if isinstance(dicts[key],dict):
            dict_item(dicts[key],subtopic)
        if isinstance(dicts[key],str):
            subtopic = subtopic.addSubTopic()
            subtopic.setTitle(dicts[key])
    return True


def load_excel(filename):
    wb= openpyxl.load_workbook(filename)
    ws = wb.active

    data = []
    for row in range(2,ws.max_row+1):
        if ws[row][4].value is None:
            ws[row][4].value = 'Null'

        if ws[row][5].value is None:
            ws[row][5].value = 'Null'

        if ws[row][3].value is not None:
            case = "/*"+ws[row][6].value+"*/"+ws[row][2].value+"/*"+ws[row][3].value+"*/"
        else:
            case = "/*" + ws[row][6].value + "*/" + ws[row][2].value
            #print(case)


        pre_result = ws[row][0].value+"$"+ws[row][1].value+"$"+"cases"+"$"+case
        if '\n' in ws[row][4].value:
            miaoshu = ws[row][4].value.split('\n')
            jieguo = ws[row][5].value.split('\n')
            for i in range(len(miaoshu)):
                if i<len(jieguo):
                    miaoshu[i] = re.sub('^\d*\.','',miaoshu[i])
                    jieguo[i] = re.sub('^\d*\.','',jieguo[i])
                    result = (pre_result+"$"+miaoshu[i]+"$"+jieguo[i])

                else:
                    miaoshu[i] = re.sub('^\d*\.', '', miaoshu[i])
                    result = (pre_result + "$" + miaoshu[i])
                after = result.split("$")
                print(result)
                final = list2dict(after)
                print(final)
                print('\n')
                data.append(final)
            #after = result.split("$")
            #print(hehe(after)
            # data.append(hehe(after))
        else:
            result= pre_result+"$"+ws[row][4].value+"$"+ws[row][5].value
            after = result.split("$")
            # print(hehe(after))
            # data.append(hehe(after))
            final = list2dict(after)
            #print(final)
            data.append(final)
    print(data)
    result = combine_dict(data)
    print(result)
    return combine_dict(data)

def combine_dict(l1):
    o = {}
    for d in l1:
        n = o
        m = d
        p = None
        while isinstance(m, dict):
            if p is not None:
                if k not in p:
                    p[k] = {}
                n = p[k]
            (k, m), = m.items()
            p = n

        while isinstance(m,str):
            p.setdefault(k,m)
            break
    return o

def list2dict(list_):
    dict_={}
    if len(list_)>2:
        dict_[list_[0]]=list2dict(list_[1:])
    else:
        dict_[list_[0]]=list_[1]
    return dict_

if __name__ == '__main__':
    final = load_excel('test.xlsx')
    design_sheet(final)









