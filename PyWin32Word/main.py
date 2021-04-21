import win32com.client
import re


if __name__ == '__main__':
    app = win32com.client.Dispatch('Word.Application')
    app.Visible = 0  # 后台运行
    app.DisplayAlerts = 0  # 不显示，不警告
    doc = app.Documents.Open('D:\\MES0205\\01_doc\\05_実装段階\\PPA\\省エネ最適化_取扱説明書_1.2.3版_横.docx')   # 打开一个已有的word文档
    # new_doc = word.Documents.Add() # 创建新的word文档

    # read
    # data = doc.words(9).text   # 从1开始
    # print(data)

    for w in doc.Words:
        if re.search(r'[\u3040-\u30FF\uFF61-\uFF9F]', str(w)) is not None:
            print(w)
            print(w.Information(3))     # get page

    doc.Close()  # 关闭 word 文档
    # app.Documents.Close(wc.wdDoNotSaveChanges)  # 保存并关闭 word 文档
    app.Quit()  # 关闭 office
