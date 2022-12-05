#pip install pandas pyodbc openpyxl xlrd pandastable pyautogui python-docx

import paxlovid_ddi
import search_nhi_med
import paxlovid_word
import copy_nhi_cloud
import download_cloud
import tkinter as tk
from tkinter import ttk
import pandastable #製作比較好看的tkinter表格
import pandas as pd
from time import sleep
import os
import webbrowser

#全域變數集中區
#把用藥清單做成全域df，後續再次評估交互作用均用這個丟回後端分析
result_original=pd.DataFrame(columns=['身分證字號','姓名','日期','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE'])

def refresh_result_original(): #重置放病人用藥明細的全域變數
    global result_original
    result_original=pd.DataFrame(columns=['身分證字號','姓名','日期','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE'])


def export_word():
    global result_original
    result_original=result_original[['日期','科別','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率']]
    paxlovid_word.save_word(output_ddi.get("1.0","end"),result_original)

def export_excel():
    global result_original
    paxlovid_ddi.output_medication_excel(result_original)

def update_nhi_database(): #更新健保用藥品項表
    def callback(url):
        webbrowser.open_new_tab(url)
    def warrning_msg():
        msg=tk.Toplevel()
        tk.Label(msg,text='檔案錯誤，請點選下方連結下載新檔案\n再重新更新資料庫\n請下載csv檔').pack()
        link = tk.Label(msg, text="https://data.gov.tw/dataset/23715", fg="blue", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: [callback("https://data.gov.tw/dataset/23715"),msg.destroy()])
        tk.Button(msg,text='關閉視窗',command=lambda:msg.destroy()).pack()
    
    file = tk.filedialog.askopenfile(title='更新健保品項檔',mode='r', filetypes=[('CSV檔', '*.csv')])
    if file:
        filepath = os.path.abspath(file.name)
        update_result=search_nhi_med.update_nhi_database(filepath)
        if update_result==True:
            tk.messagebox.showinfo(title='更新結果', message='更新成功')
        else:
            warrning_msg()

def renew_frame_table(df):   
    global frame_table
    #TK沒有辦法清除treeview或frame，所以這邊先直接破壞掉本來的frame_table，再重做一個
    if frame_table!=None:
        frame_table.destroy()
        frame_table=tk.LabelFrame(frame_table_main,bg='lightyellow',text='用藥清單')
        frame_table.pack(fill='both',expand=True)
    result_table=pandastable.Table(frame_table,dataframe=df,showtoolbar=False, showstatusbar=False,height=150)
    result_table.show()
    
def check_ddi(result): #呈現DDI結果，包含敘述與表格
    global result_original
    output_ddi.delete('1.0','end') #清掉本來的資料
    output_ddi.insert('end',result[0]) #插入新資料
    result_med=result[1]
    result_original=result[2]
    renew_frame_table(result_med)
    
def check_ddi_his(): #從HIS撈資料出來
    idno=pt_idno.get()
    result=paxlovid_ddi.final_check_idno(idno) #DDI判斷結果
    check_ddi(result)
    
def recheck_ddi(): #重新確認表格的DDI
    global result_original
    result_original=result_original.sort_values(by=['日期'],ascending=False).drop_duplicates(subset=['ATC_CODE'], keep='first')
    result=paxlovid_ddi.final_check_df(result_original)
    check_ddi(result)

def toplevel(): #跳出選取自行輸入藥品選項表
    def add_med(event): #插入藥品到現有藥品清單
        global result_original
        e=event.widget #取得被點擊欄位
        index_id=e.identify('item',event.x,event.y) #被點擊欄位的定位
        #開始逐一取值
        en_name=e.item(index_id,'value')[0]
        ch_name=e.item(index_id,'value')[1]
        chemo_name=e.item(index_id,'value')[2]
        atc_code=e.item(index_id,'value')[3]
        med_id=e.item(index_id,'value')[4]
        #做成字典，之後轉df
        new_med_dict={'身分證字號':[''],
                    '姓名':[''],
                    '日期':[''],
                    '科別':['手動加入'],
                    '醫師':['手動加入'],
                    '院內代碼':[med_id],
                    '商品名':[en_name],
                    '學名':[chemo_name],
                    '中文名':[ch_name],
                    '每次劑量':[''],
                    '天數':[''],
                    '總量':[''],
                    '途徑':[''],
                    '頻率':[''],
                    'ATC_CODE':[atc_code]}
        new_med=pd.DataFrame.from_dict(new_med_dict)
        add_med_result=paxlovid_ddi.add_new_med(result_original,new_med) #插入現有表格
        result_original=add_med_result[0] #更新全域變數
        renew_frame_table(add_med_result[1]) #更新TK的表
        sleep(0.487)
        top_level.destroy()             
    column=search_type.get()
    keyword=search_medication.get()
    if column=='All':
        result=search_nhi_med.search_contain_all(keyword)
    else:
        result=search_nhi_med.search_contain(column,keyword) #搜尋健保藥品檔
    if result.empty==True: #找不到資料的話跳彈出視窗
        tk.messagebox.showerror(title='查無資料', message='查無關鍵字藥品') 
    else:
        top_level = tk.Toplevel()
        top_level_notice=tk.Frame(top_level,background='yellow')
        top_level_notice.pack()
        top_level_table=tk.Frame(top_level)
        top_level_table.pack()
        tk.Label(top_level_notice,text='點兩下加入藥品，加入後本視窗自動關閉',background='yellow',font=('微軟正黑體','20')).pack()
        #用TK原生功能做table
        xscrollbar=tk.Scrollbar(top_level_table,orient="horizontal") #X軸滾輪
        xscrollbar.pack(side='bottom',fill='x') #X軸滾輪位置
        yscrollbar=tk.Scrollbar(top_level_table) #Y軸滾輪
        yscrollbar.pack(side='right',fill='y') #Y軸滾輪位置
        cols=result.columns.to_list() #讓資料的column名變成list，後面讓treeview使用
        tree=ttk.Treeview(top_level_table)
        tree.pack(fill='both',expand=True)
        tree['columns']=cols
        tree['show'] = 'headings' #隱藏index 
        rowcount=1
        for i in cols: #開始插入column名
            tree.column(i,anchor='center')
            tree.heading(i,text=i,anchor='center')
        tree.tag_configure('rowcolor',background='#ADD8E6') #單雙行顏色不同，目前無法啟動功能
        for index, row in result.iterrows(): #逐行插入資料
            if rowcount%2==1:
                tree.insert("",'end',text=index,values=list(row))
            else:
                tree.insert("",'end',text=index,values=list(row),tags=('rowcolor'))
            rowcount+=1
        tree.bind('<Double-1>',add_med)
        xscrollbar.config(command=tree.xview) #滾輪與文字輸入方塊綁定
        yscrollbar.config(command=tree.yview)
        tree.config(xscrollcommand=xscrollbar.set) #文字輸入方塊與滾輪綁定
        tree.config(yscrollcommand=yscrollbar.set)

def clear_all():
    #清除table區域
    global frame_table,result_original
    if frame_table!=None:
        frame_table.destroy()
        frame_table=tk.LabelFrame(frame_table_main,bg='lightyellow',text='用藥清單')
        frame_table.pack(fill='both',expand=True)
    pt_idno.delete('0','end') #清掉病歷號欄位
    output_ddi.delete('1.0','end') #清掉DDI結果欄位
    search_type.current(0) #自行輸入藥品選項回歸預設
    search_medication.delete('0','end') #清掉自行輸入藥品欄位
    #全域變數重置
    refresh_result_original()
    #result_original=pd.DataFrame(columns=['身分證字號','姓名','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE'])
    
def copy_from_cloud():
    global result_original
    copy_nhi_cloud.copy_cloud()
    cloud_df=copy_nhi_cloud.nhi_cloud()
    if type(cloud_df)==str:
        tk.messagebox.showerror(title='剪貼簿資料來源錯誤', message=cloud_df)
    else:
        add_med_result=paxlovid_ddi.add_new_med(result_original,cloud_df) #插入現有表格
        result_original=add_med_result[0] #更新全域變數
        renew_frame_table(add_med_result[1]) #更新TK的表
        recheck_ddi()
        
def download_nhi_cloud():
    global result_original
    if id_password.get()=='': #很容易會忘記輸入醫事人員卡密碼
        tk.messagebox.showerror(title='未輸入密碼', message='未輸入醫事人員卡密碼')
    else:
        cloud=download_cloud.next_page(id_password.get())
        cloud_df=cloud[0]
        if refresh_var.get()==True:
            refresh_result_original()
        add_med_result=paxlovid_ddi.add_new_med(result_original,cloud_df) #插入現有表格
        result_original=add_med_result[0] #更新全域變數
        renew_frame_table(add_med_result[1]) #更新TK的表
        recheck_ddi()

def open_record(): #打開歷史紀錄
    os.system('start EXCEL.EXE files/record.xlsx')

root = tk.Tk()
#取得螢幕的長寬
Width=root.winfo_screenwidth()
Height=root.winfo_screenheight()
root.option_add('*Font', '微軟正黑體') #更改字體
root.title('Paxlovid交互作用檢核系統')
root.geometry(str(Width)+'x'+str(int(Height*0.92)))
root.state('zoomed') #開啟時最大化
root.resizable(1, 1) #視窗可變動大小
try: #如果找不到ico檔，就忽略他
    root.iconbitmap(r'files/ico.ico')
except:
    pass

#以下開始框架建立
frame_top=tk.Frame(root,bg='lightblue')
frame_top.pack(fill='x')
frame_table_main=tk.Frame(root,bg='lightyellow')
frame_table_main.pack(fill='both',expand=True)
frame_table=tk.LabelFrame(frame_table_main,bg='lightyellow',text='用藥清單')
frame_table.pack(fill='both',expand=True)
frame_text_main=tk.LabelFrame(root,bg='lightgreen',text='藥物交互作用判斷結果')
frame_text_main.pack(fill='both',expand=True)
frame_text=tk.Frame(frame_text_main,bg='lightgreen')
frame_text.pack(fill='both',expand=True)

#以下開始頂端文字區域
title_label=tk.Label(frame_top,text='Paxlovid交互作用檢核系統',bg='lightblue',font='微軟正黑體 20 bold')
title_label.grid(row=0,column=0,columnspan=2,sticky='W')
#title_label.bind('<Double-1>',update_nhi_database)
tk.Button(frame_top,text='清除全部資料',background='salmon',command=clear_all).grid(row=0,column=3,sticky='W',padx=5,ipadx=5)
tk.Button(frame_top,text='輸出成Word',background='deepskyblue',command=export_word).grid(row=0,column=4,sticky='W',padx=5,ipadx=5)
tk.Button(frame_top,text='輸出藥品清單',background='lightgreen',command=export_excel).grid(row=0,column=5,sticky='W',padx=5,ipadx=5)
tk.Label(frame_top,text='            Design by 方志文 藥師',bg='lightblue',font='微軟正黑體 9',anchor='se').grid(row=0,column=6,columnspan=3,sticky='W')
tk.Label(frame_top,text='手動插入藥品：',bg='lightblue').grid(row=2,column=0,sticky='E')
search_type_var=pd.read_pickle(r'files\nhi_med.pkl').columns.tolist() #用檔案來源取得欄名，避免輸入錯誤
del search_type_var[5:] #刪掉不要的欄位
search_type_var.append('All') #插入全部選項
search_type=ttk.Combobox(frame_top,values=search_type_var)
search_type.grid(row=2,column=1,sticky='W')
search_type.current(0)
search_medication=tk.Entry(frame_top)
search_medication.grid(row=2,column=2,sticky='W',padx=5)
search_medication.bind('<Return>',lambda e:toplevel())
tk.Button(frame_top,text='搜尋藥品',command=toplevel).grid(row=2,column=3,sticky='W',padx=5,ipadx=5)
tk.Button(frame_top,text='再次查詢交互作用',command=recheck_ddi).grid(row=2,column=4,sticky='W',padx=5,ipadx=5)
tk.Label(frame_top,text='雲端比對：',bg='lightblue').grid(row=3,column=0,sticky='E')
tk.Button(frame_top,text='剪貼簿複製雲端藥歷',command=copy_from_cloud).grid(row=3,column=1,sticky='W',padx=5,ipadx=5)
tk.Label(frame_top,text='輸入醫事人員卡密碼：',bg='lightblue').grid(row=3,column=2,sticky='E')
id_password=tk.Entry(frame_top, show="*")
id_password.grid(row=3,column=3,sticky='E')
id_password.bind('<Return>',lambda e:download_nhi_cloud())
tk.Button(frame_top,text='爬蟲雲端藥歷',command=download_nhi_cloud).grid(row=3,column=4,sticky='W',padx=5,ipadx=5)
#做一個checkbox決定要不要清除查雲端前資料，預設要刪除
refresh_var=tk.BooleanVar() #布林值變數
refresh_checkbox=tk.Checkbutton(frame_top, text='是否清除先前資料', variable=refresh_var,bg='lightblue')
refresh_checkbox.grid(row=3,column=5,sticky='W',padx=5,ipadx=5)
refresh_checkbox.select()

frame_text_main.rowconfigure(0,weight=1)
frame_text_main.columnconfigure(0,weight=1)

#以下開始DDI結果文字方塊語法
xscrollbar=tk.Scrollbar(frame_text,orient="horizontal") #X軸滾輪
xscrollbar.pack(side='bottom',fill='x') #X軸滾輪位置
yscrollbar=tk.Scrollbar(frame_text) #Y軸滾輪
yscrollbar.pack(side='right',fill='y') #Y軸滾輪位置
output_ddi=tk.Text(frame_text) #文字輸入方塊
output_ddi.pack(fill='both',expand=True)
xscrollbar.config(command=output_ddi.xview) #滾輪與文字輸入方塊綁定
yscrollbar.config(command=output_ddi.yview)
output_ddi.config(xscrollcommand=xscrollbar.set) #文字輸入方塊與滾輪綁定
output_ddi.config(yscrollcommand=yscrollbar.set)

#建立頂層功能列
menubar=tk.Menu(root)
menubar_1=tk.Menu(menubar, tearoff=0)    
menubar_1.add_command(label="開啟歷史評估記錄", command=open_record) #新增選項
menubar_1.add_command(label="更新健保用藥品表", command=update_nhi_database) #新增選項
menubar.add_cascade(label="其他功能", menu=menubar_1) #摺疊功能列
root.config(menu=menubar)

root.mainloop()