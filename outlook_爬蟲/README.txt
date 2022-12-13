Outlook爬蟲程式說明：

此程式可以依照查詢條件，檢索含地端與雲端收件匣內所有資料夾的信件，並以xlsx檔產出。
檢索語法開放編輯於Fitrer.txt，其查詢語言為DASL，且其檔案編碼為UTF-8。

以下為部分語法說明：

1. urn:schemas:httpmail:subject：表示主旨。
2. urn:schemas:httpmail:textdescription：表示內文。
3. urn:schemas:httpmail:datereceived：表示信件接收日期。
4. urn:schemas:httpmail:hasattachment：表示附檔數量。
5. 多重條件可以使用OR與AND，但每一個指令必須用「()」包覆，否則程式可能不會理會。
6. 已知模糊查詢(Like)後面帶的文字要用「'% 」與「%'」包覆，其他則用「'」與「'」包覆即可。
7. 「【】」為註解框架，即框選範圍內為註解，須留意不可多層框選(此不為DASL規則，為此程式限定)。
8. 可接受換行或縮排。
9. 以下為查詢語法範例：
@SQL=
(
	("urn:schemas:httpmail:subject" Like \'%Receive_Purchase_Orders%\')【主旨】
	OR ("urn:schemas:httpmail:textdescription" Like \'%工單建立結果%\')【內文】
) 
AND ("urn:schemas:httpmail:datereceived" > \'2022/04/01\')【接收日期】
AND ("urn:schemas:httpmail:datereceived" < \'2022/06/01\')
AND ("urn:schemas:httpmail:hasattachment" = 0)【附檔數量】