# 標準函式庫 (不需要在 requirements.txt 中)
import urllib.request
import urllib.parse
from urllib.parse import unquote
import gzip
import io
import shutil
from http.cookiejar import CookieJar
import datetime
from io import BytesIO # 用於 Excel 輸出到記憶體

# 第三方函式庫 (已列在 requirements.txt 中)
import requests
from urllib.parse import unquote
from lxml import etree
import pandas as pd
from bs4 import BeautifulSoup
import numpy as np
import cv2
from ddddocr import DdddOcr

# 初始化 DdddOcr
ocr = DdddOcr()
def recognize_captcha(image_bytes): # 接收圖片的原始位元組
    # 這裡直接使用傳入的位元組進行 OCR
    res = ocr.classification(image_bytes)
    return res
def year_week_X_days_ago(n):# 获取6天前年週
  current_date = datetime.datetime.now()
  six_days_ago = current_date - datetime.timedelta(days=n)
  YW = six_days_ago.strftime("%Y,%W")
  return YW
def parse_table(table_html):#網頁原碼獲取表格副程式
    soup = BeautifulSoup(table_html, 'html.parser')
    table = soup.find('table')
    # 将表格转换为DataFrame
    df = pd.read_html(str(table),header=None)[0]
    return df
def try_to_login(church_detail):#登入副程式，church_detail=[church name, church id, church accunt, church pwd]
  login_payload={
  'district':'1',
  'church_id':church_detail[1],
  'account':church_detail[2],
  'pwd':church_detail[3],
  'language':'zh-tw',
  'captcha_code': 'you will know later'
  }
  are_we_in=False
  try_login_turn=0
  rs=requests.session()
  while are_we_in==False:
    captcha_len=0
    check_turn=0
    while captcha_len!=6:
      check_turn+=1
      res=rs.get('https://www.chlife-stat.org/lib/securimage/securimage_show.php',stream=True,verify=False)
      # 將圖片內容讀取到記憶體中
      captcha_image_bytes = res.content
      captcha_text = recognize_captcha(captcha_image_bytes)
      captcha_len=len(captcha_text)
    # 打印結果
    #print('驗證碼辨識次數:',check_turn)
    #print("驗證碼辨識結果:",captcha_text)
    url='https://www.chlife-stat.org/authenticate.php'
    login_payload['captcha_code']=captcha_text
    res=rs.post(url=url,data=login_payload,verify=False)
    res.encoding='utf-8'
    content=res.text
    are_we_in=not("驗證碼錯誤" in content)
    try_login_turn+=1
  return rs # 成功後返回 session object
def get_data(rs,year_week,meeting_ID,age_group_ID):
  url='https://www.chlife-stat.org/weekly_report.php'
  detial_payload={
  'year_week':year_week,
  'church_level':' 3',
  'meetings[]':meeting_ID,
  'opt_church_life':' 1',
  'opt_baptized_year':' 1',
  'opt_baptized_week':' 1',
  'show_role[]':age_group_ID
  }
  res=rs.post(url=url,data=detial_payload,verify=True)
  res.encoding='utf-8'
  content=res.text
  table = parse_table(content)
  table.set_index([('會所','會所'),('大區','大區'),('小區','小區')], inplace=True)
  table.index.names=[None,None,None]
  return table
def renew_data(specific,year_week):
  rs=requests.session()
  url={"大專":"https://www.chlife-stat.org/reports/calculate_taipei_tertiary_school.php",
  "青職":"https://www.chlife-stat.org/reports/calculate_taipei_work_saints.php",
  "青少年":"https://www.chlife-stat.org/reports/calculate_taipei_secondary.php",
  }
  Year_Week=list(map(int, year_week.split(',')))
  if specific=="大專":
    detial_payload={'year_week':year_week}
  else:
    detial_payload={'year':Year_Week[0],'month':get_week_month(year_week)}
  res=rs.post(url=url[specific],data=detial_payload,verify=True)
def get_specific_data(specific,year_week):
  rs=requests.session()
  Year_Week=list(map(int, year_week.split(',')))
  month=get_week_month(year_week)
  Specific_url={'青職':"https://www.chlife-stat.org/reports/taipei_work_saints.php",
  '青少年':"https://www.chlife-stat.org/reports/taipei_secondary_school.php",
  "大專":"https://www.chlife-stat.org/reports/taipei_tertiary_school.php",
  }
  Specific_disc_name={'青職':"文一",
  '青少年':"文山一區",
  "大專":"文山一區",
  }
  if specific=='大專':
    url=Specific_url[specific]+"?year_week="+str(Year_Week[0])+"%2C"+str(Year_Week[1])
  else:
    url=Specific_url[specific]+"?year_from="+str(Year_Week[0])+"&month_from="+str(month)+"&year_to="+str(Year_Week[0])+"&month_to="+str(month)
  res=rs.get(url=url,verify=True)
  res.encoding='utf-8'
  content=res.text
  table = parse_table(content)
  if specific=='大專':
    table.set_index([('大區','大區'),('會所','會所')], inplace=True)
    table1=table.loc[Specific_disc_name[specific],:]
    # 將合併的結果轉換為DataFrame
    sum_table1 = pd.DataFrame(table1.sum(axis=0)).T
    # 將合併的結果添加到原始表格的底部
    table1 = pd.concat([table1, sum_table1], ignore_index=False)
  else:
    table.set_index([('月份','大區','大區'),('月份','會所別','會所別')], inplace=True)
    first_day = datetime.date(Year_Week[0], 1, 1)
    # 找到第一天是星期幾
    first_day_weekday = first_day.weekday()
    # 計算第一週的星期日
    sunday = first_day + datetime.timedelta(days=(6 - first_day_weekday))
    # 計算該週的星期日
    sunday += datetime.timedelta(weeks=(Year_Week[1] - 1))
    # 找出星期日在該月份中是第幾個週日
    sunday_position = (sunday.day - 1) // 7 + 1
    columns_level0=str(Year_Week[0])+'年'+str(month)+"月"
    number_to_chinese=['零','一','二','三','四','五']
    columns_level1='第'+number_to_chinese[sunday_position]+'週'
    table1=table.loc[Specific_disc_name[specific],(columns_level0,columns_level1)]
  table1.index=['H10','H13','H27','H42','H53','H67','H73','H77','H87','H89','合計']
  return table1
from io import BytesIO

def generate_final_excel_bytes(what_we_need_10, what_we_need_87, compare_10_church, compare_10_church_87,
                               #secondary_school_data, tertiary_school_data, work_saints_data, 
                               results, year_week):
    
    # 這裡假設 church_list 已經在模組級別 (scraper_logic.py 文件頂部) 定義了
    global church_list
    
    # 創建一個字典，鍵是工作表名稱，值是對應的 DataFrame
    excel_data_to_write = {
        "表單填寫所需數據": what_we_need_10,
        "表單填寫所需數據(87分區版)": what_we_need_87,
        '文一總數據': compare_10_church,
        '文一總數據(87分區版)': compare_10_church_87,
        #"青少年專項網站數據": secondary_school_data,
        #"大專專項網站數據": tertiary_school_data,
        #"青職專項網站數據": work_saints_data,
    }
    
    # 加入每個會所的單獨數據
    for i in church_list:
        # results[i] 包含了每個會所的 results 數據
        if i in results:
            excel_data_to_write[i] = results[i]
        
    # 創建一個記憶體中的 BytesIO 物件
    output = BytesIO()
    
    # 使用 pd.ExcelWriter 將所有數據寫入這個記憶體物件
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in excel_data_to_write.items():
            # 確保數據是 DataFrame 且非空
            if isinstance(df, pd.DataFrame) and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=True, header=True)
            elif isinstance(df, dict): # 處理 results 字典的防呆
                 # 如果 results 還是字典（不應該發生），跳過
                 pass
            
    # 寫入完成後，返回 BytesIO 物件的內容 (位元組)
    output.seek(0)
    return output.getvalue()

# 在 Streamlit 的 app.py 中使用：
# excel_data = to_excel_bytes(table, "數據表")
# st.download_button(
#    label="📥 下載 Excel 檔案",
#    data=excel_data,
#    file_name=f"{church_name}.xlsx", # 使用 add_time_to_name 處理過的路徑
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# )
def add_date_to_name(name,add_time=""):
  # 将DataFrame保存为Excel文件
  add_time2=add_time.replace(',', '_')
  path=add_time2+"_"+name+".xlsx"
  return path
weekday_dict = {0:"(一)",1:"(二)",2:"(三)",3:"(四)",4:"(五)",5:"(六)",6:"(日)"}
def show_X_days_ago(n):
  # 获取当前日期和时间
  now = datetime.datetime.now()
  n_days_ago = now - datetime.timedelta(days=n)
  # 获取星期几的中文表示
  chinese_weekday = weekday_dict[n_days_ago.weekday()]
  # 格式化日期和星期几
  formatted_date = n_days_ago.strftime("%Y/%m/%d"+chinese_weekday)
  return formatted_date
def change_columns_key(table,level,old_key,new_key):
  # 获取当前的 MultiIndex 标签
  if level=="none":
    labels = table.columns
  else:
    labels = table.columns.get_level_values(level)
  # 将需要修改的标签值改为 '新标签'
  new_labels = list(labels)
  indices = [i for i, x in enumerate(new_labels) if x == old_key]
  for qq in indices:
    new_labels[qq]=new_key
  # 更新 MultiIndex 标签
  if level==0:
    table.columns = pd.MultiIndex.from_arrays([new_labels, table.columns.get_level_values(1)])
  elif level==1:
    table.columns = pd.MultiIndex.from_arrays([table.columns.get_level_values(0),new_labels])
  else:
    table.columns = new_labels
  return table

def change_index_key(table,level,old_key,new_key):
  # 获取当前的 MultiIndex 标签
  if level=="none":
    labels = table.index
  else:
    labels = table.index.get_level_values(level)
  # 将需要修改的标签值改为 '新标签'
  new_labels = list(labels)
  indices = [i for i, x in enumerate(new_labels) if x == old_key]
  for qq in indices:
    new_labels[qq]=new_key
  # 更新 MultiIndex 标签
  if level==0:
    table.index = pd.MultiIndex.from_arrays([new_labels, table.index.get_level_values(1), table.index.get_level_values(2)])
  elif level==1:
    table.index = pd.MultiIndex.from_arrays([table.index.get_level_values(0),new_labels, table.index.get_level_values(2)])
  elif level==2:
    table.index = pd.MultiIndex.from_arrays([table.index.get_level_values(0), table.index.get_level_values(1),new_labels])
  else:
    table.index = new_labels
  return table

def get_week_month(year_week):
  Year_Week=list(map(int, year_week.split(',')))
  # 找到該年的第一天
  first_day = datetime.date(Year_Week[0], 1, 1)
  # 找到第一天是星期幾
  first_day_weekday = first_day.weekday()
  # 計算第一週的星期日
  sunday = first_day + datetime.timedelta(days=(6 - first_day_weekday))
  # 計算該週的星期日
  sunday += datetime.timedelta(weeks=(Year_Week[1] - 1))
  return sunday.month
def sum_rows_with_name(table,level0,level1):
  table=pd.DataFrame(table.sum(axis=1))
  columns = pd.MultiIndex.from_tuples([(level0,level1)])
  table.columns=columns
  return table
#基本資訊
church_list=['H10','H13','H27','H42','H53','H67','H73','H77','H87','H89']
church_meetings={
"H10":["H10",'37','38','2312','39','40','1473','2026','2801','2820'],
"H13":["H13",'37','38','2312','39','40','1473','2026','2794','65'],
"H27":["H27",'37','38','2312','39','40','1473','2026','2800','2896'],
"H42":["H42",'37','38','2312','39','40','1473','2026','2759','2817'],
"H53":["H53",'37','38','2312','39','40','1473','2026','2734','2834'],
"H67":["H67",'37','38','2312','39','40','1473','2026','2786','2785'],
"H73":["H73",'37','38','2312','39','40','1473','2026','2793','2837'],
"H77":["H77",'37','38','2312','39','40','1473','2026','2812','2813','2495'],
"H87":["H87",'37','38','2312','39','40','1473','2026','2798','2833','2809'],
"H89":["H89",'37','38','2312','39','40','1473','2026','2760','2841','2720'],
}#位置：0=名字,1=主日,2=家聚會受訪,3=家聚會出訪,4=小排,5=禱告,6=福音出訪,7=晨興,8=生命讀經,9=同伴禱告,10=兒童家聚會


def run_scraper_and_process(year_week_input, church_details):
    # 1. 處理輸入 (取代原來的 input() 邏輯)
    # 由於 Streamlit 已經給了我們處理好的 year_week_input，這裡只需要處理預設邏輯
    # 雖然 app.py 已經設定了預設值，但為了程式碼健壯性，我們可以再檢查一次
    if not year_week_input:
        year_week = year_week_X_days_ago(-1)
    else:
        year_week = year_week_input

    results={}
    compare_10_church=pd.DataFrame({})
    #獲取10個會所數據
    for i in church_list:
      # 步驟 1: 登入並獲取 Session
      rs_session = try_to_login(church_detail=church_details[i])
      # 步驟 2: 使用該 Session 爬取數據 (假設 try_to_login 成功登入)
      table=get_data(rs=rs_session,year_week=year_week,meeting_ID=church_meetings[i][1:],age_group_ID=['1','2','3',"4","5"])
      table=change_index_key(table,0,'總計',i)
      # table=change_columns_key(table,0,'研讀生命讀經','生命讀經')
      # table=change_columns_key(table,0,'讀生命讀經','生命讀經')
      # table=change_columns_key(table,0,'生命讀經追求','生命讀經')
      table=change_columns_key(table,0,'人人禱告','同伴禱告')
      results[i]=table
      compare_10_church=pd.concat([compare_10_church, table.iloc[-1:]], ignore_index=False)
      column_sum = compare_10_church.sum(axis=0).to_frame().T
    new_index = pd.MultiIndex.from_arrays([['合計'],[''],['']],names=(None,None,None))
    column_sum.index = new_index
    compare_10_church=pd.concat([compare_10_church, column_sum], ignore_index=False)
    new_index = compare_10_church.index.droplevel([1,2])#刪除多餘行標籤
    compare_10_church.index = new_index

    #幫H87分區
    H87_table1=results["H87"].loc[('台北市召會第八十七會所','一大區')]
    H87_sum_1=pd.DataFrame(H87_table1.sum(axis=0)).T
    H87_sum_1.index=['H87(社區)']
    H87_table2=results["H87"].loc[('台北市召會第八十七會所','二大區')]
    H87_sum_2=pd.DataFrame(H87_table2.sum(axis=0)).T
    H87_sum_2.index=['H87(學生)']
    compare_10_church_87=compare_10_church.drop("H87")
    top_half = compare_10_church_87.iloc[:8]
    bottom_half = compare_10_church_87.iloc[8:]
    compare_10_church_87= pd.concat([top_half, H87_sum_1,H87_sum_2, bottom_half],ignore_index=False)

    #生成表單所需數據
    what_we_need_10 = pd.concat([compare_10_church.loc[:,[('今年受浸','小計'),('今年受浸','青職'),('主日','小計'),('福音出訪','小計'),('當週受浸','小計')]],
                                sum_rows_with_name(compare_10_church.loc[:,[('家聚會受訪','小計'),('家聚會出訪','小計')]],"家聚會","小計"),
                                compare_10_church.loc[:,[('晨興','小計'),('小排','小計'),('生命讀經','小計'),('同伴禱告','小計'),('召會生活','小計'),('兒童主日','小計')]],
                                sum_rows_with_name(compare_10_church.loc[:,[('召會生活','學齡前'),('召會生活','小學')]],"召會生活","兒童"),
                                sum_rows_with_name(compare_10_church.loc[:,[('家聚會出訪','學齡前'),('家聚會受訪','學齡前'),('家聚會出訪','小學'),('家聚會受訪','小學')]],"家聚會","兒童"),
                                compare_10_church.loc[:,('主日','中學')],
                                sum_rows_with_name(compare_10_church.loc[:,[('家聚會受訪','中學'),('家聚會出訪','中學')]],"家聚會","中學"),
                                compare_10_church.loc[:,[('小排','中學'),('主日','大專'),('福音出訪','大專')]],
                                sum_rows_with_name(compare_10_church.loc[:,[('家聚會受訪','大專'),('家聚會出訪','大專')]],"家聚會","大專"),
                                compare_10_church.loc[:,[('主日','青職'),('福音出訪','青職')]],
                                sum_rows_with_name(compare_10_church.loc[:,[('家聚會受訪','青職'),('家聚會出訪','青職')]],"家聚會","青職"),
                                compare_10_church.loc[:,('生命讀經','青職')],], axis=1,ignore_index=False)
    what_we_need_10=what_we_need_10.swaplevel(axis=1)
    what_we_need_10=change_columns_key(what_we_need_10,0,"小計","全會所")
    for i in ['H77','H87',"H89"]:
      try:
        if compare_10_church.loc[i,("兒童家聚會","小計")]>what_we_need_10.loc[i, ('兒童',"家聚會")]:
          what_we_need_10.loc[i, ('兒童',"家聚會")]= compare_10_church.loc[i,("兒童家聚會","小計")]
      except:
        pass
    what_we_need_10.loc['合計', ('兒童',"家聚會")]=what_we_need_10.loc['H10':'H89', ('兒童',"家聚會")].sum()
    #生成表單所需數據(H87分區版)
    what_we_need_87 = pd.concat([compare_10_church_87.loc[:,[('今年受浸','小計'),('今年受浸','青職'),('主日','小計'),('福音出訪','小計'),('當週受浸','小計')]],
                                sum_rows_with_name(compare_10_church_87.loc[:,[('家聚會受訪','小計'),('家聚會出訪','小計')]],"家聚會","小計"),
                                compare_10_church_87.loc[:,[('晨興','小計'),('小排','小計'),('生命讀經','小計'),('同伴禱告','小計'),('召會生活','小計'),('兒童主日','小計')]],
                                sum_rows_with_name(compare_10_church_87.loc[:,[('召會生活','學齡前'),('召會生活','小學')]],"召會生活","兒童"),
                                sum_rows_with_name(compare_10_church_87.loc[:,[('家聚會出訪','學齡前'),('家聚會受訪','學齡前'),('家聚會出訪','小學'),('家聚會受訪','小學')]],"家聚會","兒童"),
                                compare_10_church_87.loc[:,('主日','中學')],
                                sum_rows_with_name(compare_10_church_87.loc[:,[('家聚會受訪','中學'),('家聚會出訪','中學')]],"家聚會","中學"),
                                compare_10_church_87.loc[:,[('小排','中學'),('主日','大專'),('福音出訪','大專')]],
                                sum_rows_with_name(compare_10_church_87.loc[:,[('家聚會受訪','大專'),('家聚會出訪','大專')]],"家聚會","大專"),
                                compare_10_church_87.loc[:,[('主日','青職'),('福音出訪','青職')]],
                                sum_rows_with_name(compare_10_church_87.loc[:,[('家聚會受訪','青職'),('家聚會出訪','青職')]],"家聚會","青職"),
                                compare_10_church_87.loc[:,('生命讀經','青職')],], axis=1,ignore_index=False)
    what_we_need_87=what_we_need_87.swaplevel(axis=1)
    what_we_need_87=change_columns_key(what_we_need_87,0,"小計","全會所")
    for i in ['H77','H87(社區)','H87(學生)','H89']:
      try:
        if compare_10_church_87.loc[i,("兒童家聚會","小計")]>what_we_need_87.loc[i, ('兒童',"家聚會")]:
          what_we_need_87.loc[i, ('兒童',"家聚會")]= compare_10_church_87.loc[i,("兒童家聚會","小計")]
      except:
        pass
    what_we_need_87.loc['合計', ('兒童',"家聚會")]=what_we_need_87.loc['H10':'H89', ('兒童',"家聚會")].sum()
    # --- 呼叫生成第一個 Excel 檔案的函數 ---
    excel_bytes_total_data = generate_final_excel_bytes(
        what_we_need_10, 
        what_we_need_87, 
        compare_10_church, 
        compare_10_church_87,
        #secondary_school_data, 
        #tertiary_school_data, 
        #work_saints_data, 
        results, 
        year_week # year_week 變數
    )
    
    # --- 呼叫生成第二個 Excel 檔案的函數 ---
    weekly_report_excel_bytes = generate_weekly_report_excel(what_we_need_87, year_week)

    # 返回兩個 Excel 檔案的位元組和週數給 app.py
    # 將所有需要的返回值列出
    return excel_bytes_total_data, weekly_report_excel_bytes, year_week, what_we_need_10


def generate_weekly_report_excel(what_we_need_87, year_week):
  # ... 您的所有處理邏輯 ...
    # 最終返回 Excel 檔案的位元組
    # return dfC_excel_bytes
    dfA = what_we_need_87.reset_index(drop=True)  # 去除行標籤
    dfA.columns = [''] * len(dfA.columns)  # 去除列標籤
    #此為2025基數，分別是全會所、兒童、青少年、大專、青職
    data_B = [
        [109, 11, 2, 4, 19],
        [125, 17, 1, 9, 15],
        [105, 14, 8, 5, 13],
        [97, 20, 6, 15, 20],
        [193, 23, 16, 4, 27],
        [128, 14, 5, 10, 21],
        [147, 12, 20, 7, 31],
        [99, 21, 7, 9, 14],
        [93, 20, 1, 3, 9],
        [51, 3, 35, 2, 9],
        [137, 18, 2, 2, 25],
        [1284, 173, 103, 69, 203]
    ]
    '''
    #此為2024基數，分別是全會所、兒童、青少年、大專、青職
    data_B = [
        [128, 7, 4, 7, 24],
        [109, 10, 2, 6, 14],
        [143, 14, 9, 5, 20],
        [129, 12, 7, 24, 12],
        [193, 24, 17, 7, 24],
        [132, 13, 7, 12, 15],
        [172, 8, 24, 9, 26],
        [102, 15, 8, 13, 8],
        [91, 20, 0, 0, 8],
        [60, 0, 29, 3, 16],
        [149, 11, 2, 3, 19],
        [1408, 134, 109, 89, 186]
    ]
    '''
    # 創建 DataFrame
    dfB = pd.DataFrame(data_B)

    dfC = pd.DataFrame(np.zeros((12, 44)))

    # 填充列 [1, 2, 3, 6, 8, 9, 11, 13, 15, 17, 19, 21, 23, 25, 26, 28, 30, 31, 33, 35, 37, 39, 41, 43]
    dfC.iloc[:, [0, 1, 2, 5, 7, 8, 10, 12, 14, 16, 18, 20, 22, 24, 25, 27, 29, 30, 32, 34, 36, 38, 40, 42]] = np.round(dfA.values)

    # 第4列 (百分比顯示，四捨五入到整數位): (dfA的第3列 - dfB的第1列) / dfB的第1列
    dfC.iloc[:, 3] = np.round((dfA.iloc[:, 2] - dfB.iloc[:, 0]) / dfB.iloc[:, 0] * 100)

    # 第5列 (百分比顯示，四捨五入到整數位): dfA的第3列 / dfB的第1列
    dfC.iloc[:, 4] = np.round(dfA.iloc[:, 2] / dfB.iloc[:, 0] * 100)

    # 第7列 (百分比顯示，四捨五入到整數位): dfA的第4列 / dfB的第1列
    dfC.iloc[:, 6] = np.round(dfA.iloc[:, 3] / dfB.iloc[:, 0] * 100)

    # 第10列 (四捨五入到小數點後兩位): dfA的第6列 / dfB的第1列
    dfC.iloc[:, 9] = np.round(dfA.iloc[:, 5] / dfB.iloc[:, 0], 2)

    # 第12列 (百分比顯示，四捨五入到整數位): dfA的第7列 / dfB的第1列
    dfC.iloc[:, 11] = np.round(dfA.iloc[:, 6] / dfB.iloc[:, 0] * 100)

    # 第14列 (百分比顯示，四捨五入到整數位): dfA的第8列 / dfB的第1列
    dfC.iloc[:, 13] = np.round(dfA.iloc[:, 7] / dfB.iloc[:, 0] * 100)

    # 第16列 (百分比顯示，四捨五入到整數位): dfA的第9列 / dfB的第1列
    dfC.iloc[:, 15] = np.round(dfA.iloc[:, 8] / dfB.iloc[:, 0] * 100)

    # 第18列 (百分比顯示，四捨五入到整數位): dfA的第10列 / dfB的第1列
    dfC.iloc[:, 17] = np.round(dfA.iloc[:, 9] / dfB.iloc[:, 0] * 100)

    # 第20列 (百分比顯示，四捨五入到整數位): dfA的第11列 / dfB的第1列
    dfC.iloc[:, 19] = np.round(dfA.iloc[:, 10] / dfB.iloc[:, 0] * 100)

    # 第22列 (四捨五入到整數位): dfA的第12列 - dfB的第2列
    dfC.iloc[:, 21] = np.round(dfA.iloc[:, 11] - dfB.iloc[:, 1])

    # 第24列 (四捨五入到整數位): dfA的第13列 - dfB的第2列
    dfC.iloc[:, 23] = np.round(dfA.iloc[:, 12] - dfB.iloc[:, 1])

    # 第27列 (四捨五入到整數位): dfA的第15列 - dfB的第3列
    dfC.iloc[:, 26] = np.round(dfA.iloc[:, 14] - dfB.iloc[:, 2])

    # 第29列 (四捨五入到小數點後兩位): dfA的第16列 / dfB的第3列
    dfC.iloc[:, 28] = np.round(dfA.iloc[:, 15] / dfB.iloc[:, 2], 2)

    # 第32列 (四捨五入到整數位): dfA的第18列 - dfB的第4列
    dfC.iloc[:, 31] = np.round(dfA.iloc[:, 17] - dfB.iloc[:, 3])

    # 第34列 (百分比顯示，四捨五入到整數位): dfA的第19列 / dfB的第4列
    dfC.iloc[:, 33] = np.round(dfA.iloc[:, 18] / dfB.iloc[:, 3] * 100)

    # 第36列 (四捨五入到小數點後兩位): dfA的第20列 / dfB的第4列
    dfC.iloc[:, 35] = np.round(dfA.iloc[:, 19] / dfB.iloc[:, 3], 2)

    # 第38列 (四捨五入到整數位): dfA的第21列 - dfB的第5列
    dfC.iloc[:, 37] = np.round(dfA.iloc[:, 20] - dfB.iloc[:, 4])

    # 第40列 (百分比顯示，四捨五入到整數位): dfA的第22列 / dfB的第5列
    dfC.iloc[:, 39] = np.round(dfA.iloc[:, 21] / dfB.iloc[:, 4] * 100)

    # 第42列 (四捨五入到小數點後兩位): dfA的第23列 / dfB的第5列
    dfC.iloc[:, 41] = np.round(dfA.iloc[:, 22] / dfB.iloc[:, 4], 2)

    # 第44列 (四捨五入到小數點後兩位): dfA的第24列 / dfB的第5列
    dfC.iloc[:, 43] = np.round(dfA.iloc[:, 23] / dfB.iloc[:, 4], 2)

    #顯示百分比的先除100
    percent_columns = [3, 4, 6, 11, 13, 15, 17, 19, 33, 39]
    dfC.iloc[:, percent_columns] = dfC.iloc[:, percent_columns] / 100

    Cname_data = [
        ('受浸','受浸', '目前受浸'),
        ('受浸','受浸', '青職受浸'),
        ('全會所', '主日','主日'),
        ('全會所', '主日','繁增律(%)'),
        ('全會所', '主日','佔基數比'),
        ('全會所', '福','福音出訪'),
        ('全會所', '福','出訪比例'),
        ('全會所', '福','受浸'),
        ('全會所', '家','家聚會'),
        ('全會所', '家','家聚會倍數'),
        ('全會所', '家','晨興'),
        ('全會所', '家','佔基數比'),
        ('全會所', '排','小排'),
        ('全會所', '排','佔基數比'),
        ('全會所', '追求','生命讀經'),
        ('全會所', '追求','佔基數比'),
        ('全會所', '禱告','人人禱告'),
        ('全會所', '禱告','佔基數比'),
        ('全會所', '主日/小排','召會生活'),
        ('全會所', '主日/小排','佔基數比'),
        ('兒童', '主日','兒童主日'),
        ('兒童', '主日','與基數相比'),
        ('兒童', '主日/小排','兒童召會生活'),
        ('兒童', '主日/小排','與基數相比'),
        ('兒童', '家','家聚會'),
        ('青少年', '主日','青少年主日'),
        ('青少年', '主日','與基數相比'),
        ('青少年', '家','家聚會'),
        ('青少年', '家','倍數'),
        ('青少年', '排','小排人數'),
        ('大專', '主日','大專主日'),
        ('大專', '主日','與基數相比'),
        ('大專', '福','出訪'),
        ('大專', '福','出訪比例'),
        ('大專', '家','家聚會'),
        ('大專', '家','倍數'),
        ('青職', '主日','青職主日'),
        ('青職', '主日','與基數相比'),
        ('青職', '福','出訪'),
        ('青職', '福','出訪比例(%)'),
        ('青職', '家','家聚會'),
        ('青職', '家','倍數'),
        ('青職', '追求','生命讀經'),
        ('青職', '追求','佔基數比'),
    ]

    # 將 tuple 列表轉換為 DataFrame
    Cname_df = pd.DataFrame(Cname_data)

    # 進行行列轉置
    Cname_df_transposed = Cname_df.transpose()

    # 將 header_rows 插入到 dfC 的上方
    dfC = pd.concat([Cname_df_transposed, dfC],axis=0, ignore_index=True)

    new_index = ["","","","H10", "H13", "H27", "H42", "H53", "H67", "H73", "H77", "H87(社區)", "H87(學生)", "H89", "合 計"]
    dfC.index = pd.Index(new_index)
    dfC=dfC.replace([np.nan, np.inf, -np.inf], "-")
    # 創建一個記憶體中的 BytesIO 物件
    output = BytesIO()
    
    # 使用 pd.ExcelWriter 將數據寫入記憶體物件
    # 注意：這裡使用 with 語句確保 writer 關閉並將數據寫入 output
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfC.to_excel(writer, sheet_name="文一每週報表", header=False)

        # 獲取 workbook 和 worksheet 物件 (保持不變)
        workbook  = writer.book
        worksheet = writer.sheets['文一每週報表']

        # 定義格式 (保持不變)
        percent_fmt = workbook.add_format({'num_format': '0%'})
        decimal_fmt = workbook.add_format({'num_format': '0.00'})
          # 為指定的百分比列應用格式，例如 [4, 5, 7, 12, 14, 16, 18, 20, 34, 40] 等列
        percent_columns = [4, 5, 7, 12, 14, 16, 18, 20, 34, 40]  # 百分比列（從 0 開始計算）

        for col in percent_columns:
            worksheet.set_column(col, col, None, percent_fmt)

        # 為小數點兩位的列應用格式，例如 [10, 29, 36, 42]
        decimal_columns = [10, 29, 36, 42]  # 小數列

        for col in decimal_columns:
            worksheet.set_column(col, col, None, decimal_fmt)

        #合併儲存格
        worksheet.merge_range('B1:C2', '受浸')
        worksheet.merge_range('D1:U1', '全會所')
        worksheet.merge_range('V1:Z1', '兒童')
        worksheet.merge_range('AA1:AE1', '青少年')
        worksheet.merge_range('AF1:AK1', '大專')
        worksheet.merge_range('AL1:AS1', '青職')
        worksheet.merge_range('D2:F2', '主日')
        worksheet.merge_range('G2:I2', '福')
        worksheet.merge_range('J2:M2', '家')
        worksheet.merge_range('N2:O2', '排')
        worksheet.merge_range('P2:Q2', '追求')
        worksheet.merge_range('R2:S2', '禱告')
        worksheet.merge_range('T2:U2', '主日/小排')
        worksheet.merge_range('V2:W2', '主日')
        worksheet.merge_range('X2:Y2', '主日/小排')
        worksheet.merge_range('AA2:AB2', '主日')
        worksheet.merge_range('AC2:AD2', '家')
        worksheet.merge_range('AF2:AG2', '主日')
        worksheet.merge_range('AH2:AI2', '福')
        worksheet.merge_range('AJ2:AK2', '家')
        worksheet.merge_range('AL2:AM2', '主日')
        worksheet.merge_range('AN2:AO2', '福')
        worksheet.merge_range('AP2:AQ2', '家')
        worksheet.merge_range('AR2:AS2', '追求')

        # 創建對齊的格式，文字置中
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        # 為所有列設定居中格式
        num_columns = len(dfC.columns)
        for col_num in range(num_columns+1):
            worksheet.set_column(col_num, col_num, None, center_format)
    output.seek(0) # 重置指標到起始位置
    return output.getvalue()