# -excel-
import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import messagebox as msgbox
import openpyxl
#訊息箱子
def ca():
    if(gender.get()==0):
        G=float(B[0])
        T=float(aa.get())
        AA=float(T/G)
    elif(gender.get()==1):
        G=float(B[1])
        T=float(aa.get())
        AA=float(T/G)
    elif(gender.get()==2):
        G=float(B[2])
        T=float(aa.get())
        AA=float(T/G)
    AA = round(AA, 2)
    messagebox = msgbox.askokcancel('台幣和美金互換', AA, icon='info')
    
#網頁擷取
ur1 = 'https://rate.bot.com.tw/xrt?Lang=zh-TW'
a1 = requests.get(ur1)
a1.encoding = 'utf-8'
a2 = BeautifulSoup(a1.text, 'html.parser')
title = a2.find_all('div', class_="visible-phone print_hide")
math = a2.find_all('td', class_="rate-content-cash text-right print_hide")
#資料轉excel
A=[0]*19
B=[0]*19
n=0
workbook = openpyxl.Workbook()
sheet = workbook.worksheets[0]
sheet['A1'] = '幣別:   '
sheet['B1'] = '現金匯率:'
for i in range(19):
    A[i]=title[i].text.strip()
    B[i]=math[n].text
    #ia='A'+str(i+1)
    sheet['A'+str(i+2)]= A[i]
    sheet['B'+str(i+2)]= B[i]    
    n=n+2
workbook.save('test.xlsx')
#小視窗設定
win = tk.Tk()
win.title('台幣匯率')
win.geometry('400x200')
#視窗物件設定
lbl00 = tk.Label(win, text='請輸入金額', bg='white', fg='black', font=('標楷體', 15))
lbl01 = tk.Label(win, text='台幣(NTW)', bg='white', fg='black', font=('標楷體', 14))
gender = tk.IntVar()
aa=tk.IntVar()
c1 = tk.Radiobutton(win, text=A[0], variable=gender, value='0')
c2 = tk.Radiobutton(win, text=A[1], variable=gender, value='1')
c3 = tk.Radiobutton(win, text=A[2], variable=gender, value='2')
btnok=tk.Button(win, text='確認', command=ca)
a=tk.Entry(win, width=12, textvariable=aa)
#位置
lbl00.grid(row=0, column=0)
lbl01.grid(row=1, column=0)
a.grid(row=1, column=1)
c1.grid(row=2, column=1)
c2.grid(row=2, column=2)
c3.grid(row=2, column=3)
btnok.grid(row=3, column=3)

win.mainloop()
