import os
import win32com.client
import configparser

config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8-sig')
filepath = config['DEFAULT']['path']

def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)

word = win32com.client.Dispatch('Word.Application')
if filepath.endswith('.docx'):
    try:
        wdFormatPDF = 17
        i_path = filepath
        o_path = rreplace(filepath, '.docx', '.pdf', 1)
        inputFile = os.path.abspath(i_path)
        outputFile = os.path.abspath(o_path)

        doc = word.Documents.Open(inputFile)
        doc.SaveAs(outputFile, FileFormat = wdFormatPDF)
        doc.Close()
    except Exception as e:
        print('中斷，出現錯誤：{}'.format(filepath))
        print(e)
    else:
        print('成功轉檔')
else:
    print('請確認副檔名')

word.Quit()