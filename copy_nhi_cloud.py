import pandas as pd
import re

def copy_cloud():#從雲端藥歷取得資料，轉成csv
    temp_str=open(r'temp/cloud_temp.csv', "w",encoding='utf-8') #開啟暫存csv
    from_clipboard = clipboard.paste() #取得剪貼簿文字
    temp_str.write(u'\ufeff'+from_clipboard)
    temp_str.close()
    
def nhi_cloud():
    #讀取csv
    temp_csv=open(r'temp/cloud_temp.csv', "r",encoding='utf-8')
    check_csv=temp_csv.read()
    temp_csv.close()
    if ('健保代碼' in check_csv)==True:
        temp_csv=open(r'temp/cloud_temp.csv', "r",encoding='utf-8')
        str_readlines=temp_csv.readlines()
        #偵測從第幾行開始是完整的一段
        i=40
        while i<len(str_readlines):
            first_line=str_readlines[i]
            if len(first_line)>50:
                print(first_line)
                break
            else:
                i+=1
        
        #下一行的間隔是固定的，但是不同版面會有差異，決定下一個是哪一個
        l=i+1
        while l<len(str_readlines):
            first_line=str_readlines[l]
            if len(first_line)>50:
                print(first_line)
                break
            else:
                l+=1
                
        interval=l-i #間隔
        
        #偵測第幾個是健保碼
        first_line_list = re.split(r'\t+', str_readlines[i])
        print(first_line_list)
        first_line=str_readlines[i]
        j=0
        while j<len(first_line_list):
            if len(first_line_list[j])==10 and (first_line_list[j][2:9].isnumeric()==True or first_line_list[j][:6]=='XCOVID'):
                print(first_line_list[j])
                break
            else:
                j+=1
                
        #偵測第幾個是開立日期
        first_line_list = re.split(r'\t+', first_line)
        k=0
        while j<len(first_line_list):
            if len(first_line_list[k])==9 and first_line_list[k].replace('/','').isnumeric()==True:
                break
            else:
                k+=1
        print(i)
        #確認第幾行的第幾個之後開始做迴圈
        nhi_code=list()
        order_date=list()
        while i<len(str_readlines)-4:
            str_list = re.split(r'\t+', str_readlines[i])
            #print(i)
            #print(str_list)
            nhi_code.append(str_list[j])
            order_date.append(str_list[k].replace('/',''))
            i+=interval #間格
            

        nhi_code_and_date={'健保碼':nhi_code,
                           '日期':order_date}
        nhi_code_df=pd.DataFrame.from_dict(nhi_code_and_date).sort_values(by=['日期'],ascending=False).drop_duplicates(subset=['健保碼'], keep='first') #排除重複
        nhi_med=pd.read_pickle(r'files\nhi_med.pkl') #讀取健保檔，用健保碼比對ATC code
        nhi_code_df=nhi_code_df.join(nhi_med.set_index('藥品代號'), on='健保碼') 
        #表格整理，讓資料與外面相同
        nhi_code_df=nhi_code_df[['健保碼','日期','藥品英文名稱','藥品中文名稱','成份','ATC_CODE']]
        nhi_code_df=nhi_code_df.rename(columns={'健保碼':'院內代碼',
                                                '藥品英文名稱':'商品名',
                                                '藥品中文名稱':'中文名',
                                                '成份':'學名'})    
        nhi_code_df[['科別','醫師']]="雲端藥歷"
        nhi_code_df[['病歷號碼','姓名','每次劑量','天數','總量','途徑','頻率']]=''
        nhi_code_df=nhi_code_df[['病歷號碼','姓名','日期','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE']]
        #ddi_result=paxlovid_ddi.ddi_result_df(nhi_code_df) #丟回DDI比對系統
    else:
        nhi_code_df='剪貼簿來源資料錯誤\n可能是複製錯資料\n或是雲端藥歷未選擇顯示健保代碼欄位'
    return nhi_code_df