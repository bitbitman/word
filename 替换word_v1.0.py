import xlrd
from docx import Document
import os


#替换doc内容
def change(doc,changed_txt,change_txt):
    for i in doc.paragraphs:
        for run in i.runs:
            run.text=run.text.replace(changed_txt,change_txt)
    
    return doc




if __name__ == '__main__':

    print("="*40)
    print(" ")
    print("将word与exe放入同一目录")
    print("不支持整段替换")
    print(" "*30+"by 徐阳")
    print(" ")
    print("="*40)



    print("输入替换次数")
    n=int(input())
    print("输入模板文档文件名")
    name=input()
    print("输入替换位置文本")
    changed_txt=input()

    for i in range(n):
        print("输入替换文本")
        change_txt=input()
        doc=Document(os.getcwd()+"/"+name+".docx")
        doc=change(doc,changed_txt,change_txt)
        h=i+1
        print("第",h,"个文件")
        print("="*25)
        doc.save(str(h)+".docx")
    print("完成 按回车退出")
    ppp=input()


