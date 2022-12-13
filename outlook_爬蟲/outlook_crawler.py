import win32com.client
import re
import pandas as pd
import PySimpleGUI as sg
import sys
from styleframe import StyleFrame

pd.options.mode.chained_assignment = None
sg.ChangeLookAndFeel('GreenTan')

print('開始運行...')

def processfolder(folder, filter, collection):
    ignoredfolders = []
    if not folder.Name in ignoredfolders:
        count = 0
        for mail in folder.Items.Restrict(filter):
            if mail.Sender.GetExchangeUser() != None:
                address = mail.Sender.GetExchangeUser().PrimarySmtpAddress
            else:
                address = mail.SenderEmailAddress
            #savemsg(mail)
            # print('|'.join([
            #     mail.SentOn.strftime('%Y/%m/%d %H:%M:%S'),
            #     mail.SenderName, 
            #     mail.Subject, 
            #     #i.body
            # ]))
            collection.append([
                folder.Name,
                mail.SentOn.strftime('%Y/%m/%d %H:%M:%S'),
                address,
                mail.Subject
            ])
            count += 1
        # print("{} 封信於資料夾 {}".format(count, folder.Name))
        for fld in folder.Folders:
            collection = processfolder(fld, filter, collection)
    return collection

try:
    with open('Filter.txt', 'rt', encoding = 'UTF-8') as f:
        Filter = f.read()
    Filter = Filter.replace('\n', '').replace('\\', '')

    for i in re.findall(r'【(.*?)】', Filter): # 去掉註解的範圍
        Filter = Filter.replace('【'+ i + '】', '')

    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    df_output = pd.DataFrame()

    for main_folder in outlook.Folders: # 搜尋所有信箱
        main_inbox = outlook.Folders[main_folder.Name]
        inbox = main_inbox.Folders(2) # 2表示收件匣
        lines = []

        lines = processfolder(inbox, Filter, lines)

        df_suboutput = pd.DataFrame(lines, columns = ['資料夾名稱', '收信時間', '寄件者', '主旨'])
        df_suboutput.insert(0, '收信匣名稱', main_inbox.Name)
        df_output = pd.concat([df_output, df_suboutput])
        print(f'收件匣 {main_inbox.Name}   OK')
    df_output = df_output.reset_index(drop = True)
except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print('【發生錯誤】中斷第{}行\n錯誤訊息：{} {}'.format(exc_tb.tb_lineno, exc_type, str(e)))
        sg.PopupOK('請確認資料來源。', font = ('Microsoft YaHei', 10))
else:
    if len(df_output) > 0:
        layout_output = [[
            [sg.Text('儲存輸出資料', font = ('Microsoft YaHei', 10))],
            [
                sg.InputText(key = 'save_filename', visible = True, enable_events = True),
                sg.FileSaveAs(target = 'save_filename', file_types = (('xlsx', '*.xlsx'), ('All Files', '*.*')))
            ],
            [sg.Submit(button_text = '確定', font = ('Microsoft YaHei', 9))]
        ]]
        window_savefile = sg.Window('儲存輸出檔案', layout_output, size = (500, 120), font = ('Microsoft YaHei', 10))

        while True:
            event, values_output = window_savefile.Read()
            if event == sg.WIN_CLOSED or event == 'Exit':
                break
            if event == 'save_filename':
                if not values_output['save_filename']:
                    sg.PopupOK('請確認輸出檔案名稱', font = ('Microsoft YaHei', 10))
            if event == '確定':
                window_savefile.close()
                StyleFrame(df_output).to_excel(values_output['save_filename'], index = False, best_fit = df_output.columns.tolist()).save()

                sg.PopupOK('儲存成功，執行完畢！', font = ('Microsoft YaHei', 10))
                break
    else:
        sg.PopupOK('無資料產出。', font = ('Microsoft YaHei', 10))



# df_output.to_csv('output.csv', encoding = 'Big5', index = False)
# print('檔案輸出完畢')