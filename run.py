import os
import sys
from win32com import client as wc


def save_doc_to_docx(file_name: str) -> None:  # doc转docx
    word = wc.Dispatch("Word.Application")
    # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件的）
    if file_name.endswith('.doc') and not file_name.startswith('~$'):
        print(file_name, "开始转换")
        doc = word.Documents.Open(os.getcwd() + "/" + file_name)  # 将文件名与后缀分割
        doc.SaveAs(os.getcwd() + "/" + os.path.splitext(file_name)[0] + '.docx', 12)  # 12表示docx格式
        doc.Close()
        os.remove(os.getcwd() + "/" + file_name)  # 删除原文件
    word.Quit()


if __name__ == '__main__':
    save_doc_to_docx(sys.argv[1])
