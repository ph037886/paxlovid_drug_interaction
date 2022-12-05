import pandas as pd
nhi_med=pd.read_pickle(r'files\nhi_med.pkl') #來源：https://data.gov.tw/dataset/23715

def search_contain(column,keyword):
    keyword=str(keyword).strip() #去除頭尾空白
    keyword=keyword.lower() #關鍵字強制轉小寫
    mask=nhi_med[column].str.lower().str.contains(keyword) #搜尋欄位強制轉小寫，並做模糊搜尋 (contains)
    search_result=nhi_med[mask]
    return search_result

def search_contain_all(keyword):
    merga_result=pd.DataFrame(columns=nhi_med.columns.tolist())
    search_type_var=nhi_med.columns.tolist() #用檔案來源取得欄名，避免輸入錯誤
    del search_type_var[5:] #刪掉不要的欄位
    for i in search_type_var:
        new_result=search_contain(i,keyword)
        merga_result=pd.concat([merga_result,new_result])
    return merga_result

def update_nhi_database(filepath):
    try:
        nhi_database=pd.read_csv(filepath,low_memory=False)
        nhi_database=nhi_database.sort_values('有效迄日',ascending=False)
        nhi_database=nhi_database.drop_duplicates(subset=['藥品代號'],keep='first')
        nhi_database=nhi_database[['藥品英文名稱', '藥品中文名稱', '成份', 'ATC_CODE', '藥品代號', '單複方', '劑型', '製造廠名稱', '規格量', '規格單位']]
        nhi_database=nhi_database.fillna('')
        nhi_database.to_pickle(r'files\nhi_med.pkl')
        return_info=True
    except:
        return_info=False
    print(return_info)
    return return_info
    