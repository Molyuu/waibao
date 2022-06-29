#!/usr/bin/python

from docxtpl import DocxTemplate
from openpyxl import load_workbook, Workbook
from shutil import move

print("======xlsx批量填入docx======")
print("请仔细阅读本目录下的使用说明.pdf！！！！")
print("请仔细阅读本目录下的使用说明.pdf！！！！")
print("请仔细阅读本目录下的使用说明.pdf！！！！")
print("========================")
a = input("你是否已经阅读了 使用说明.pdf 并完全理解了如何操作？(是/否)")
if a == "是":
    print("好的！")
else:
    print("好的！")
    b = input("请按下回车键键以退出程序！")
    exit(0)


filename = ""

doc = DocxTemplate("SRC.docx")
wb = load_workbook("SRC.xlsx", data_only=True)
sheet = wb.active
for row in sheet:
    rowMap = map(lambda x: x.value, row)
    if row[0].row == 1:
        title = list(rowMap)
    else:
        context = dict(zip(title, rowMap))
        filename = context["filename"]
        print("处理", filename, "中", end="\r")
        doc.render(context)
        doc.save("%s.docx" % filename)
        move("%s.docx" % filename, "./OUT")
        print("处理", filename, "完毕！.")

c = input("处理完成！按下回车键来退出")
