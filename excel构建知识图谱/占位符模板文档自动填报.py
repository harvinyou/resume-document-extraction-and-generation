import os
import xlrd
import datetime
import time
from mailmerge import MailMerge

def remvocePiont(str):
    index = str.find(".")
    return str[0:index]
path ='文档生成\模型.xlsx'
xl = xlrd.open_workbook(path)
print("开始填报模板文档，请稍等······ ")
# 读取第一个表
table = xl.sheet_by_name(xl.sheet_names()[0])
# 获取表中行数
nrows = table.nrows
# 生成Word文档存储目录
path_base_name = '文档生成'

for i in range(nrows):  # 循环逐行打印
    if i > 0:
        template_path = "文档生成\个人简历模板.docx"
        doc = MailMerge(template_path)  # 打开模板文件
        # print(doc.get_merge_fields())
        id=str(table.row_values(i)[18])
        name = str(table.row_values(i)[0])
        borndate = str(table.row_values(i)[1])
        sex = str(table.row_values(i)[2])
        tel = str(table.row_values(i)[3])
        bestedu = str(table.row_values(i)[4])
        nativeplace = str(table.row_values(i)[5])
        city = str(table.row_values(i)[6])
        trust = str(table.row_values(i)[7])
        edu = str(table.row_values(i)[8])
        edutime = str(table.row_values(i)[9])
        worktime = str(table.row_values(i)[10])
        itemtime =str(table.row_values(i)[11])
        eduschool = str(table.row_values(i)[12])
        company = str(table.row_values(i)[13])
        worklist = str(table.row_values(i)[14])
        worktitle = str(table.row_values(i)[15])
        itemname=str(table.row_values(i)[16])
        itemresponsibility=str(table.row_values(i)[17])





        # 以下为填充模板中对应的域，
        doc.merge(name=name,
                  borndate=borndate,
                  sex=sex,
                  tel=tel,
                  bestedu=bestedu,
                  nativeplace=nativeplace,
                  city=city,
                  trust=trust,
                  edu=edu,
                  edutime=edutime,
                  worktime=worktime,
                  itemtime=itemtime,
                  eduschool=eduschool,
                  company=company,
                  worklist=worklist,
                  worktitle=worktitle,
                  itemname=itemname,
                  itemresponsibility=itemresponsibility


                  )
        # 使用文件名 日期名称
        path_name = os.path.join(path_base_name,datetime.datetime.now().strftime("%Y-%m-%d"))
        if not os.path.exists(path_name):
            os.makedirs(path_name)
        word_name = path_name +"\\"+ str(i)+"_"+ id +"_"+ name + '.docx'

        doc.write(word_name)
        print("第"+str(i)+"个文档填报成功，文档名为：" + id)
        doc.close()
print("填报成功,WORD文件保存在 :" + path_base_name)
time.sleep(5)

