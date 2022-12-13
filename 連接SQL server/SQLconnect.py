import pyodbc

server = 'N-MIS-KJ' #伺服器名稱
database = 'test' # 資料庫名稱
username = 'sa' # 登入帳號
password = 'Sa12345678' # 登入密碼
datatable = 'dbo.Table_1' # 資料表名稱

cnxn = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, 
    unicode_results = True
)
cursor = cnxn.cursor()

cursor.fast_executemany = True # 加速用的

# 準備要使用的query
query = "SELECT * FROM {}".format(datatable)
# query = "INSERT INTO {} ({}) VALUES {};".format(datatable, columns_query, ','.join(data_query[start: end]))

cursor.execute(query) # 送出query

# 取得查詢的資料欄位名稱
column_names = [column[0] for column in cursor.description]
print(column_names)

# 讀取查詢的資料內容
rows = cursor.fetchall()
for row in rows:
    print(row)   # 亦可直接取得欄位資料，如row.place，即取得資料中欄位place的資料

cnxn.commit() # 必須加入這行，否則對資料庫的所有操作將會失效
cursor.close()