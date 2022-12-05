import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import tkinter as tk
import os

def ADtoROC(AD): #西元轉民國
    try:
        AD=datetime.datetime.strptime(str(AD),"%Y/%m/%d")
    except:
        AD=datetime.datetime.strptime(str(AD),"%Y-%m-%d")
    ROC=str(AD.year-1911)+str(AD.strftime('%m'))+str(AD.strftime('%d'))
    if len(ROC)==6:
        ROC='0'+ROC
    return ROC

def ROC_today(): #民國今日日期
    return ADtoROC(datetime.date.today())

ddi=pd.read_excel(r'files/Paxlovid_DDI.xlsx') #用Excel保留拓展空間

def output_pkl(df):
    df['評估日期']=datetime.date.today()
    try: #看看有沒有之前的紀錄檔，有的話就打開沒有的話跳過，後續直接做一個新的
        old_record=pd.read_pickle(r'files/record.pkl')
        df=pd.concat([df,old_record])
    except:
        pass
    df.to_pickle(r'files/record.pkl')
    try: #如果Excel檔被打開，會跳不能存檔的Error，這時候另存一個Excel出來，除非兩個excel都打開，不然不會錯誤，因為主要運作的都是pkl檔，不影響記錄
        df.to_excel(r'files/record.xlsx',index=False,freeze_panes=(1,0))
    except PermissionError:
        df.to_excel(r'files/record_2.xlsx',index=False,freeze_panes=(1,0))

def ddi_result_df(df): #接收藥品清單df，進行交互作用評估
    result=df.join(ddi.set_index('ATC code'), on='ATC_CODE')
    output_pkl(result)
    result=result.dropna(subset=['禁忌']) #有交互作用，禁忌那欄一定會有資料
    result=result.fillna(value={'衛福部建議處理方式':'衛福部無提供交互作用資料，但其他資料庫來源有交互作用資料'}) #衛福部資料有些沒有，那些沒有的加上註釋
    result=result.fillna('')#其他NA填空白    
    return result

def final_check(original,result,paxlovid_creatine): #接收結果df進行文字處理
    result=result.sort_values(by='禁忌', ascending=False)
    if len(original)==0:
        result_word='尚未輸入藥品清單，或病人於6個月內未於本院看診'
    else:
        result_word=('病人：'+original.iloc[len(original)-1,1]+' ('+original.iloc[len(original)-1,0]+')\n' #增加藥品時，是從前面開始加，且增加的df病人個資欄位會是空白，所以從最後一個row取病人個資。要-1是因為如果是自行輸入品項len出來會是1，iloc取值會錯誤
                +paxlovid_creatine+'\n')
        
        if len(result)==0:  
            result_word=(result_word
                    +'病人無會與Paxloivd產生交互作用藥品\n'
                    +'\n')
        else:
            result_word=(result_word
                    +'下列 '+str(len(result))+' 項藥品可能與Paxlovid有交互作用\n')
            contraindicated_result=result[result['禁忌']==1]
            if len(contraindicated_result)!=0:
                result_word=(result_word
                        +'※※※注意：有'+str(len(contraindicated_result))+'項藥品為禁忌交互作用※※※\n')
            result_word=result_word+'\n'
            i=0
            while i<len(result):
                contraindicated=''
                if int(result.iloc[i,16])==1:
                    contraindicated='※※※注意：本筆交互作用為禁忌※※※\n'
                if str(result.iloc[i,5])[0:1]=='E':
                    contraindicated=contraindicated+'外用藥，本系統採ATC code比對，外用藥是否適用交互作用情形，請依建議判斷\n'
                result_word=result_word+(str(i+1)+'.\n'
                    +'開立日期/醫師：'+str(result.iloc[i,2])+' '+result.iloc[i,4]+' '+result.iloc[i,3]+'\n'
                    +'學名：'+result.iloc[i,6]+'\n'
                    +'商品名：'+result.iloc[i,7]+'\n'
                    +'中文名：'+result.iloc[i,8]+'\n'
                    +contraindicated
                    +'可能結果：'+'\n'
                    +result.iloc[i,17]+'\n'
                    +'衛服部交互作用建議：'+'\n'
                    +result.iloc[i,18]+'\n'
                    +'Uptodate-Patient Management：'+'\n'
                    +result.iloc[i,19]+'\n\n')
                i+=1
    new_result=original[['日期','科別','醫師','院內代碼','學名','商品名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE']]
    return result_word, new_result,original

def add_new_med(original,add_on): #新增藥品
    new_med=add_on.append(original)
    new_med_modify=new_med[['科別','醫師','院內代碼','學名','商品名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE']]
    return new_med, new_med_modify

def final_check_df(df): #從前端全域變數藥品清單，再做DDI比對
    result=ddi_result_df(df)
    output_tk=final_check(df,result,str(''))
    return output_tk

def output_medication_excel(df):
    df=df[['日期','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率']]
    #儲存檔案
    file_name=tk.filedialog.asksaveasfilename(title='另存excek',initialfile=ROC_today(),defaultextension="*.xlsx", filetypes=[("Excek", ".xlsx")])
    df.to_excel(file_name,index=False,freeze_panes=(1,0))
    os.system(file_name)