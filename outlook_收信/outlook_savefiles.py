import re
import os
import win32com.client

folder = r'D:\Topunion_KJ\Python_related\outlook收信測試\test'
subject_rule = '^(高通新需求)(.*?)(檔案)$'

# folder = r'\\192.168.1.20\業務訂單資訊共用檔\3AL-高通\PO'
# subject_rule = '^(EBCRPPRD:)(.*?)(TOP UNION ELECTRONICS CORP)$'

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.Folders['KJ-2021A'].Folders['收件匣'].Folders['部門委託'].Folders['李蕙安']    #######
# inbox = outlook.Folders['anna_li@topunion.com.tw'].Folders['收件匣'].Folders['高通客戶'].Folders['Jasmine']    #######
items = inbox.Items.Restrict('[UnRead] = True') # 抓取 Inbox 未讀取郵件
sub = [i for i in items if any(re.findall(subject_rule, i.Subject, re.IGNORECASE))]
mail = [i for i in sub if i.Attachments.Count > 0] # 判斷是否有附件

m, n = 0, 0
for i in mail:
    i.UnRead = False # 將郵件標示為讀取
    m += 1
    tfg1 = i.SentOn.strftime('%Y %m月') # 收件時間戳記
    tfg2 = i.SentOn.strftime('%m%d') # 收件時間戳記
    try:
        for j in i.Attachments:
            if j.FileName.rsplit('.')[-1].lower() in ['pdf', 'xlsx', 'xls']: # 附件格式 pdf 才存檔
                if not os.path.exists(os.path.join(folder, tfg1)):
                    os.mkdir(os.path.join(folder, tfg1))
                if not os.path.exists(os.path.join(folder, tfg1, tfg2)):
                    os.mkdir(os.path.join(folder, tfg1, tfg2))
                j.SaveAsFile(os.path.join(folder, tfg1, tfg2, j.FileName))
                print('已執行收信時間 {} 的信，有附檔，檔名為 {}。'.format(i.SentOn.strftime('%Y/%m/%d %H:%M:%S'), j.FileName))
                n += 1
            else:
                print('已執行收信時間 {} 的信，無符合格式(pdf、xlsx、xls)的檔案。')
                #continue
    except:
        print('執行收信時間 {} 的信出現錯誤！'.format(i.SentOn.strftime('%Y/%m/%d %H:%M:%S')))
        pass
print(f'程式執行完成，共讀取 {m} 封信並儲存 {n} 個檔案。')