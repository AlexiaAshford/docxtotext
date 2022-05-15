import os
import sys

from win32com import client as wc


def save_doc_to_docx(dir1: str, dir2: str) -> None:  # doc转docx
    word = wc.Dispatch("Word.Application")
    for file_name in os.listdir(os.path.join(os.getcwd(), dir1)):
        # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件的）
        if file_name.endswith('.doc') and not file_name.startswith('~$'):
            print(file_name, "开始转换")
            try:
                doc = word.Documents.Open(os.path.join(os.getcwd(), dir1, file_name))  # 将文件名与后缀分割
                # 12表示docx格式
                doc.SaveAs(os.path.join(os.getcwd(), dir1, os.path.splitext(file_name)[0] + '.docx'), 12)
                doc.Close()
                os.remove(os.path.join(os.getcwd(), dir1, file_name))  # 删除原文件
            except Exception as error:
                print(file_name, "转换失败", error)
                continue
    word.Quit()


if __name__ == '__main__':
    save_doc_to_docx(sys.argv[1], sys.argv[2])
