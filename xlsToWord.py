#這版本修復了逗點課別問題

import tkinter as tk
import os
import datetime
import shutil
import time
from tkinter.constants import NO
import win32com.client as win32 
import openpyxl
import pyautogui

#建立視窗
root = tk.Tk()
root.title("一鍵請購驗收單")

#創建標籤
label1 = tk.Label(root, text="輸入 num：")
label1.grid(row=0, column=0)

#創建文字框
num = tk.StringVar()
entry1 = tk.Entry(root, textvariable=num)
entry1.grid(row=0, column=1)

#創建執行狀況顯示文字框
label2 = tk.Label(root, text="執行結果")
label2.grid(row=1)

v1 = []
lblString = ""

def tax(money):
    s = int(money)
    tax =str(round(s*0.05))
    return addComma(tax)

def moneyWithTax(money):
    s = int(money)
    moneyWithTax =str(round(s*1.05))
    return addComma(moneyWithTax)

def addComma(s):
    #s = s.replace("NTD", "") # 刪除 NTD
    result = ""
    count = 0
    for digit in s[::-1]:
        if count == 3:
            result = "," + result
            count = 0
        result = digit + result
        count += 1

    #result = "NTD " + result
    print(result)
    return result


#按下Enter按鈕時執行的功能
def execute():
    global num
    global v1
    global lblString

    #取得num的值
    num_value = num.get()
    
        ##這是將excel檔案複製到txt檔案中


    # 指定 Excel 檔案路徑
    path = r"D:\Users\E941\Desktop\手工驗收單\112備品俊龍板_0410.xlsx"

    # 打開 Excel 檔案
    wb = openpyxl.load_workbook(path)

    # 宣告欲尋找的字串
    #num = "20231185"

    # 宣告一個空的 List 變數
    #v1 = []

    # 尋找 num 的位置
    for sheet in wb.worksheets:
        for row in sheet.rows:
            for cell in row:
                if str(cell.value) == num_value:
                    for i in range(25):
                        v1.append(str(row[i].value))
                        print(v1[i])


    if len(v1)==0:
        lblString = "於找不到這一筆資料。\n"
        label2['text'] = lblString
        return
    else:
        for i in range(25):
            if v1[i]==None:
                v1[i]="0"

    # 複製文件
    src_file = r"D:\Users\E941\Desktop\採購表單\07_驗收表單\F-AD-2-09-04A請購驗收單.docx"

    new_filename =v1[3]
    # 檢查檔案名稱是否合法，若不合法則替換特殊符號
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|',' ']
    for char in invalid_chars:
        if char in new_filename:
            new_filename = new_filename.replace(char, '_')
    dst_file_name = num_value + new_filename + "_驗收"

    dst_folder = r"D:\Users\E941\Desktop\手工驗收單"
    dst_file = os.path.join(dst_folder, dst_file_name + ".docx")
    os.system(f"copy {src_file} {dst_file}")

    # 打開文件
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(dst_file)

    # 移動游標並進行文字操作
    word.Selection.MoveRight()
    word.Selection.Find.Execute("部門：")
    word.Selection.MoveRight()
    if v1[6]=="儀控" or v1[6]=="機械" or v1[6]=="電氣":
        word.Selection.TypeText("維修課")
    elif v1[6]=="管理課":
        word.Selection.TypeText("管理課")
    elif v1[6]=="工安課":
        word.Selection.TypeText("工安課")
    elif v1[6]=="水處理":
        word.Selection.TypeText("運轉課")
    else:
        msg =tk.Message(root,text="注意XX課有誤",font=("Algerian",18,"bold"),bg='#ADFEDC',fg='#00CACA')
        word.Selection.TypeText("")

    word.Selection.Find.Execute("日期:")
    word.Selection.MoveRight()
    today = datetime.date.today().strftime("%Y/%m/%d")
    word.Selection.TypeText(today)
    for i in range(8):
        word.Selection.Delete()
    word.Selection.Find.Execute("編號:")
    word.Selection.MoveRight()
    word.Selection.TypeText(v1[1])
    word.Selection.Find.Execute("編號")
    word.Selection.MoveDown()
    word.Selection.TypeText("1")#輸入1
    word.Selection.MoveRight()#前往品名
    word.Selection.TypeText(v1[3])
    word.Selection.MoveRight(1,2)#跳過規範前往數量1式
    word.Selection.TypeText(v1[4])
    word.Selection.MoveRight()#前往單價
    #pyautogui.hotkey('ctrl', 'E')#置中
    word.Selection.TypeText("NTD " + addComma(v1[19]))
    word.Selection.MoveRight(1,2)#跳過用途前往驗收數量
    word.Selection.TypeText(v1[4]+"\n\n合計\n稅額\n總計")
    word.Selection.MoveRight()#前往驗收金額
    #pyautogui.hotkey('ctrl', 'E')#置中
    word.Selection.TypeText("NTD " + addComma(v1[19])+"\n\nNTD " + addComma(v1[19])+"\nNTD "+tax(v1[19])+"\nNTD "+moneyWithTax(v1[19]))

    word.Selection.Find.Execute("1.廠商:")
    word.Selection.MoveRight()#前往品名
    word.Selection.TypeText(v1[16])#輸入廠商
    
    word.Selection.Find.Execute("經管")
    word.Selection.MoveRight(1,7)#
    if v1[6]=="儀控" or v1[6]=="機械" or v1[6]=="電氣":
        word.Selection.TypeText("維修課")
    elif v1[6]=="管理課":
        word.Selection.TypeText("管理課")
    elif v1[6]=="工安課":
        word.Selection.TypeText("工安課")
    elif v1[6]=="水處理":
        word.Selection.TypeText("運轉課")
    else:
        msg =tk.Message(root,text="注意XX課有誤",font=("Algerian",18,"bold"),bg='#ADFEDC',fg='#00CACA')
        word.Selection.TypeText("")

    #word.Selection.Find.Execute("2.統編:")

    # 儲存文件
    doc.Save()

    # 關閉文件
    doc.Close()

    # 關閉Word應用程序
    word.Quit()
    lblString =  "已成功創建"+dst_file_name+".docx\n\t目前仍未支援自動輸入統一編號，再請自行上網查詢，感謝。\n" + lblString
    label2['text'] = lblString

    #最後資料清空
    v1 = []

btn1 = tk.Button(root, text="確定",command=execute)
btn1.grid(row=0, column=2)

root.mainloop()
