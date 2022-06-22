import os
import shutil
import pandas as pd
import docx


def porcess_word(name,mon1,mon2,mon3,stat,left,company,year,month,day,numbers,person):
    doc1=docx.Document('src.docx') #读取文件
    list_row = ['客户档案名称','债权金额','债务金额','抵消金额','剩余债权债务确认','抵消后余额','清理账簿公司' , 'A' , 'B','C','序号','税务经理']#word中需要替换的参数
    list_replace=[name , mon1 , mon2 , mon3 , stat , left , company , year , month , day , numbers , person]#读取Excel的参数
    for kk in range(len(list_row)):#通过循环进行逐个参数的替换
        text_row= list_row[kk]
        text_replace = list_replace[kk]
        for p in doc1.paragraphs:
            if text_row in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if text_row in inline[i].text:
                        text = inline[i].text.replace(str(text_row), str(text_replace))
                        inline[i].text = text

    doc1.save('%s-%s-%s.docx'%(person,company,numbers)) #保存文件
    shutil.move('%s-%s-%s.docx'%(person,company,numbers),'out')
if __name__ == "__main__":
    file_path = r'./'  #指定路径
    os.chdir(file_path)
    excel_file = pd.read_excel('src.xlsx') #读取Excel数据

    for i in range(len(excel_file)):  #逐行进行数据处理
        data_temp=excel_file.loc[i]
        name = str(data_temp['客户档案名称'])
        mon1_prev = float(data_temp['债权金额'])
        mon1 = f'{mon1_prev:,.2f}'
        mon2_prev = float(data_temp['债务金额'])
        mon2 = f'{mon2_prev:,.2f}'
        mon3_prev = float(data_temp['抵消金额'])
        mon3 = f'{mon3_prev:,.2f}'
        stat = str(data_temp['剩余债权债务确认'])
        left_prev = float(data_temp['抵消后余额'])
        allleft = f'{left_prev:,.2f}'
        company = str(data_temp['清理账簿公司'])
        year = data_temp['年']
        month = data_temp['月']
        day = data_temp['日']
        numbers = data_temp['序号']
        person = data_temp['税务经理'] 

        porcess_word(name,mon1,mon2,mon3,stat,allleft,company,year,month,day,numbers,person)#调用处理函数
    print("完成！")
    input("按 Enter键/回车 退出")
