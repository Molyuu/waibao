from docxtpl import DocxTemplate
from openpyxl import Workbook, load_workbook

def find(xlsx):
    #读取Excel中的第一行内容
    cell=[]
    wb=load_workbook(filename=xlsx)
    sheet=wb.active
    for i in sheet[1]:
        cell.append(i.value)
    return cell

def replace(xlsx,docx):
    doc=DocxTemplate(docx)
    wb=load_workbook(filename=xlsx)
    sheet=wb.active
    for j in find(xlsx):
        for k in sheet:
            context= {'{%s}'% j : ""}
    

if __name__ == "__main__":
    print(find("./gork.xlsx"))
    