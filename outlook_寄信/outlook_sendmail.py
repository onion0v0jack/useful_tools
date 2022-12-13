import win32com.client

file_path_1 = r'D:\Topunion_KJ\Python_related\outlook_sendmail\test\test_1.txt'
file_path_2 = r'D:\Topunion_KJ\Python_related\outlook_sendmail\test\test_2.xlsx'

# 注意這兩行不能直接合併寫 outlook = win32com.client.Dispatch('Outlook.Application').GetNameSpace('MAPI') 
# 否則後面的 outlook.CreateItem(0) 會跑不出來
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')  
mail = outlook.CreateItem(0)

# 如果最後面有加上「mail.send()」，即寄出，那有沒有加這行沒差。
# 但如果沒有加上「mail.send()」，那有加行就會跳出草稿視窗，否則不會做任何事情
# mail.Display() 

# 收件者與cc名單
# 如果收件者們是一批名單，且可能有變動，建議步驟如下：
# 1. 建立一個xlsx檔案(mail.xlsx)，然後把收件者寫在裡面。
# 2. 利用pandas讀取mail.xlsx，讀取紀錄mail的欄位，記得讀取後面加上「.value.tolist()」，暫時命名為receiver_list。若有類別需求就用篩選。
# 3. 最後把這個裝滿string的list：receiver_list，轉成一個分號分隔的string。用「'; '.join(receiver_list)」就可以了
mail.To = 'kj_chen@topunion.com.tw'
# mail.CC = 'somebody@company.com' ; alice_tung@topunion.com.tw

mail.Subject = 'Test Email'  # 主旨

# 內文有兩種寫法：
# Body：純文字
# HTMLBody：有美化或甚至放表格、圖片、超連結等需求才用，且該文字要用html語法寫。
# mail.Body = "This is the normal body"
mail.HTMLBody = '<h3>This is HTML Body</h3>'

# 夾帶檔案
# 如果只是固定一兩個檔案就分開寫；但如果是一群，甚至某資料夾內部所有檔案，那建議善用os套件，取得內部所有檔案「路徑」，再用for就解決了。
mail.Attachments.Add(file_path_1)
mail.Attachments.Add(file_path_2)

# 最後一定要加上這行，才會寄出，否則就會變草稿(參考「mail.Display() 」)。
mail.Send()


print('Done')
