import pandas as pd
import docx


def search(word):
    doc = docx.Document("./gork.docx")
    all_paragraphs = doc.paragraphs
    list1 = []
    for paragraph in doc.paragraphs:
        str1 = paragraph.text
        if str1.find(word) != -1:
            list1.append(str1)
    if list1 == []:
        list1.append("failed,not found")
    return list1


def replace():
    data_src = pd.read_excel(io=r"./gork.xlsx", nrows=1)
    for i in data_src:
        ser_word = "{%s}" % i
        if search(ser_word) == ["failed,not found"]:
            print("抱歉，没有在表格中找到要求替换的数据，请检查格式是否正确")


if __name__ == "__main__":
    replace()
    