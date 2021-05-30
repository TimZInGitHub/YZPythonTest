# coding:utf-8

from bs4 import BeautifulSoup
import requests
import xlwt

def set_style(name,height,bold=False):
    #初始化表格样式;
    style=xlwt.XFStyle()
    #为样式创建字体
    font=xlwt.Font()
    # print(font)
    font.name=name
    font.bold=bold
    # font.colour_index=4
    font.color_index = 4
    font.height=height

    style.font =font
    return style


def write_excel():
    # 步骤1：打开目标网页"https://www.hkca.com.hk/tc/members-list"；
    # 步骤2：抓包查看网页期显示原理，确定爬取方案，是通过HTML信息解析，还是通过webAPI获取信息；
    # 步骤3：确定当前网页可通过webAPI，检查API参数原理；
    # 步骤4：模拟请求webAPI，发送get请求
    url = "https://hkca-api-new.wtc.work/about/member?page=1&limit=2000&lang=tc"

    r = requests.get(url)
    # 步骤5：获取返回的json数据
    json_result = r.json()

    items = json_result["items"]
    print(json_result)

    # 步骤6：创建Excel表格
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿对象
    sheet3 = f.add_sheet(u"sheet3", cell_overwrite_ok=True)

    # 步骤7：确定目标解析内容，建立表头
    row0 = [u'index', u'en_company_name', u'en_address', u'tc_company_name', u'sc_company_name', u'tel', u'fax', u'email', u'slug', u'website']

    # 步骤8：编辑表头（表格第一行）
    for i in range(0, len(row0)):
        sheet3.write(0, i, row0[i], set_style('Times New Roman', 220, True))

    # 步骤9：遍历解析接口信息，获取目标数据
    for i in range(0, len(items)):
        sheet3.write(i + 1, 0, i+1, set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 1, items[i]["info"]["en"]["company_name"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 2, items[i]["info"]["en"]["address"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 3, items[i]["info"]["tc"]["company_name"], set_style('Times New Roman', 220, True))
        # sheet3.write(i + 1, 4, items[i]["info"]["sc"]["company_name"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 5, items[i]["tel"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 6, items[i]["fax"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 7, items[i]["email"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 8, items[i]["slug"], set_style('Times New Roman', 220, True))
        sheet3.write(i + 1, 9, items[i]["website"], set_style('Times New Roman', 220, True))
    # 步骤10：保存表格
    f.save("zcl.xlsx")
    
write_excel()