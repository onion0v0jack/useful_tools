import os
import win32com.client
import configparser

config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8')
folder_path = config['DEFAULT']['path']

def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)

word = win32com.client.Dispatch('Word.Application')
for filename in os.listdir(folder_path):
    filepath = os.path.join(folder_path, filename)
    try:
        if filename.endswith('.docx'):
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
        print('已完成：{}'.format(filepath))

word.Quit()