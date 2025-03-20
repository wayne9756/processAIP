import pandas as pd
import numpy as np
import re
import unicodedata


def title_case_with_exceptions(text, exceptions=None):
    if pd.isna(text):
        return text  # 如果是空值，直接返回
    text = str(text)
    if exceptions is None:
        exceptions = {"to", "and", "of", "in", "on", "at", "with", "by", "for", "a", "an", "the"}  # 預設例外單字集合

    special_exceptions = {"FullText"}
    if any(exception in text for exception in special_exceptions):
        return text  # 如果包含特殊字串，則不進行處理，直接返回原字串
    words = text.split()  # 將文字拆分為單字列表
    title_cased_words = [
        word.capitalize() if (word not in exceptions or i == 0) else word for i, word in enumerate(words)
    ]
    return " ".join(title_cased_words)


# 字符正規化函數
def normalize_text(text, exceptions = None):
    if pd.isna(text):
        return text  # 如果是空值，直接返回
    # 如果 text 中包含 "FullText"，則保持不變
    text = str(text).strip()
    text = unicodedata.normalize('NFKC', text)  # 全形轉半形
    text = re.sub(r'\s+', ' ', text)  # 去掉多餘的空格
    text = text.replace('and','&')
    text = text.replace('eJournal','')
    text = text.replace('eJournals','')
    return text

# 读取数据
client = pd.read_excel('./AIP_Client_Pivot.xlsx')
master = pd.read_excel('./AIP_Master_List.xlsx')
client.dropna(how='any', inplace=True)

# 确保 'Sub Year' 是数字类型
client["Sub Year"] = pd.to_numeric(client["Sub Year"], errors="coerce")

# 正規化需要比較的字段
client['Product Name'] = client['Product Name'].apply(normalize_text).apply(title_case_with_exceptions)
master['Subject Collection'] = master['Subject Collection'].apply(normalize_text).apply(title_case_with_exceptions)
master['Title'] = master['Title'].apply(normalize_text).apply(title_case_with_exceptions)

# 按 'Sub Year' 升序排列
sorted_client = client.sort_values(by='Sub Year', ascending=True)

# 其他分类的产品名称
other_of_subjection = [
    'Management Core',
    'Industry & Public Sector',
    'Community College Collection',
    'Further Education College Collection',

    'Management eJoural Portfolio',
    'Engineering, Computing & Technology eJournal Portfolio',
    'Emerald Full Text + FullText',
    'Emerald Fulltext Plus',
    
    'Emerald Management 111',
    'Emerald Management 120',
    'Emerald Management 125',
    'Emerald Management 140',

    'Emerald Management 150',
    'Emerald Management 160',
    'Emerald Management 175',
    'Emerald Management 200',

    'Premier',
]

backfiles = [
    'Emerald Backfiles',
    'Emerald Backfiles Additions',
]

subjections = master['Subject Collection'].unique() 

journals = master['Title'].unique()

subjections_backfiles = []

def checkTheBackfiles( products: pd.Series, backfiles: list):
    for product in products:
        if(type(product) != float and ('Backfiles' in product or 'Backfile' in product)):
            backfiles.append(product)

def checkTheRangeBackfile(sub_master: pd.DataFrame, product_row: pd.Series):
    sub_master = sub_master.copy()  # 避免 SettingWithCopyWarning
    
    # 确保日期字段是数字类型
    sub_master['Online date, start'] = pd.to_numeric(sub_master['Online date, start'], errors="coerce")
    sub_master['Online date, End'] = pd.to_numeric(sub_master['Online date, End'], errors="coerce")
    # 筛选符合条件的记录
    filtered_df = sub_master.copy()
    product_name = product_row['Product Name']
    if 'Additions' in product_name:
        start_year =  1994
        end_year = 2013
    else:
        start_year = 0
        end_year = 2006
    start_year = pd.Series(np.maximum(start_year, sub_master['Online date, start']))
    end_year = pd.Series(np.minimum(end_year, sub_master['Online date, End']))
    filtered_df['start_year'] = start_year
    filtered_df['end_year'] = end_year
    result = filtered_df.apply(
        lambda row: pd.DataFrame({
            **{col: [row[col]] * (row['end_year'] - row['start_year'] + 1) for col in filtered_df.columns},
            'Sub Year': range(row['start_year'], row['end_year']+1),
            'Product Name': [product_name] * (row['end_year'] - row['start_year'] + 1)
        }), axis=1).reset_index(drop=True)
    result = pd.concat(result.values).reset_index(drop=True)
    result = result.drop('start_year', axis= 1)
    result = result.drop('end_year', axis=1)
    return result



# 检查范围的函数
def checkTheRange(sub_master: pd.DataFrame, product_row: pd.Series):
    sub_master = sub_master.copy()  # 避免 SettingWithCopyWarning
    sub_year = product_row['Sub Year']
    
    # 确保日期字段是数字类型
    sub_master['Online date, start'] = pd.to_numeric(sub_master['Online date, start'], errors="coerce")
    sub_master['Online date, End'] = pd.to_numeric(sub_master['Online date, End'], errors="coerce")
    
    # 筛选符合条件的记录
    filtered_df = sub_master[
        (sub_master['Online date, start'] <= sub_year) &
        (sub_year <= sub_master['Online date, End'])
    ].copy()
    
    if not filtered_df.empty:
        filtered_df.loc[:, 'Sub Year'] = product_row['Sub Year']
        filtered_df.loc[:, 'Product Name'] = product_row['Product Name']
    return filtered_df

valid_list = []
wrong_list = []

checkTheBackfiles(client['Product Name'], backfiles=subjections_backfiles)

for _, row in sorted_client.iterrows():
    normalized_product_name = row['Product Name']
    if normalized_product_name in backfiles:
        sub_master = master[master[normalized_product_name] == '*'] 
    elif normalized_product_name in subjections_backfiles:
        # Remove the "backfile" in the normalized_product_name
        normalized_product_name = " ".join(normalized_product_name.split()[:-1])
        sub_master = master[master['Subject Collection'] == normalized_product_name]
    else:
        if normalized_product_name in other_of_subjection: 
            sub_master = master[master[normalized_product_name] == '*'] 
        elif normalized_product_name in subjections:
            sub_master = master[master['Subject Collection'] == normalized_product_name]
        elif normalized_product_name in journals:
            sub_master = master[master['Title'] == normalized_product_name]
        else:
            wrong_list.append(row['Product Name'])
            continue  # 若匹配不到，跳過後續處理
        if not sub_master.empty:
            valid_list.append(checkTheRange(sub_master, row))
            continue
    if not sub_master.empty:
        valid_list.append(checkTheRangeBackfile(sub_master, row))




# 保存匹配結果
if valid_list:  
    result_df = pd.concat(valid_list, ignore_index=True)
    result_df = result_df.sort_values(by="Title")
    result_df.to_excel('test.xlsx', index=False)
else:
    print("没有符合条件的数据。")

if wrong_list:
    wrong_df = pd.DataFrame({'Product Name': wrong_list})
    wrong_df.to_excel("wrong_product.xlsx", index=False)

# 合併年份範圍函數
def merge_sub_years(df: pd.DataFrame):
    """將年份列表合併為範圍格式，如 [2005~2006, 2025]"""
    years = sorted(df['Sub Year'].dropna().astype(int).unique())  # 排序並去重
    ranges = []
    start = years[0]
    for i in range(1, len(years)):
        if years[i] != years[i-1] + 1:  # 判斷年份是否連續
            end = years[i-1]
            ranges.append(f"{start}~{end}" if start != end else f"{start}")
            start = years[i]
    ranges.append(f"{start}~{years[-1]}" if start != years[-1] else f"{start}")
    return f"[{', '.join(ranges)}]"

# 合併相同 Title 的數據
titles = result_df['Title'].unique()
merge_result = []

for title in titles:
    sub_df = result_df[result_df['Title'] == title]  # 選擇對應 title 的數據
    during = merge_sub_years(sub_df)  # 合併年份範圍
    product_names = " #".join(sub_df['Product Name'].dropna().unique())  # 合併唯一的 Product Names
  
    # 建立合併後的 DataFrame，保留必要的列
    merged_df = pd.DataFrame({
        'Status': [sub_df['Status'].iloc[0]],  # 假設 Status 對於相同 Title 是一致的
        'Title': [title],  # Title 是分組基準
        'Acronymn':[sub_df['Acronymn'].iloc[0]],
        'Platform ISSN': [sub_df['Platform ISSN'].iloc[0]],  # 假設 ISSN 是一致的
        'During': [during],  # 新增合併的年份範圍
        'Product Names': [product_names]  # 合併的 Product Names
    })
    merge_result.append(merged_df)

final_df = pd.concat(merge_result, ignore_index=True)
final_df.to_excel("merged_titles.xlsx", index=False)
