
import pandas as pd
import numpy as np
import re
import unicodedata

class BaseProduct:
    def __init__(self, product_name: str = "", years:list = [] ):
        self.product_type = type(self) 
        self.set_name(product_name)
        self.set_order_period(years)

    def set_name(self, name: str):
        if not isinstance(name, str):
            raise ValueError("Product name must be a string.")
        self.product_name = name

    def set_order_period(self, years: list|int):
        if self.product_type == Backfile:
            if 'Additions' in self.product_name:
                start_year =  1994
                end_year = 2013
            else:
                start_year = 0
                end_year = 2006    
            self.years = np.arange(start= start_year,stop=end_year+1)
        elif self.product_type == Package or self.product_type == Subscription:
            self.years =  years
        else:
            raise ValueError("UNKNOWN product type.")
    def get_order_period(self):
        return self.years
    def get_name(self):
        return self.product_name



    def __str__(self):
        """Readable representation for the user."""
        return f"{self.product_type.capitalize()}(Product Name: {self.product_name}, Years: {self.years}-{self.years})"


class Backfile(BaseProduct):
    def __init__(self, product_name: str, years:list = [] ):
        super().__init__(product_name, years)

class Subscription(BaseProduct):
    def __init__(self, product_name: str, years:list = [] ):
        super().__init__(product_name, years)

class Package(BaseProduct):
    def __init__(self, product_name: str, years:list = [] ):
        super().__init__(product_name, years)

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


def createPeriod(names:pd.Series, df:pd.DataFrame, kind):
    the_set_of_ordered_product = {}
    for name in names:
        sub_df = df[df['Product Name'] == name]
        if sub_df.empty:  # 检查是否有数据
            raise ValueError("Product name DO NOT EXIST.")
        product = kind(name, sub_df['Sub Year'].values )
        the_set_of_ordered_product[f'{name}'] = product

    return the_set_of_ordered_product

def processBackfile(row: pd.Series, backfiles_dict: dict, backfile_name: set, df: pd.DataFrame, idx: int):
    name = str(row['Subject Collection']) + " Backfiles"
    if name in backfile_name:
        # Update the original DataFrame using df.at
        row['Backfiles'] = row['Backfiles'] + f'@{backfiles_dict[name].get_name()} '
        df.at[idx, 'Backfiles'] = row['Backfiles']
        df.at[idx, 'Backfiles Period'] = backfiles_dict[name].get_order_period()
    # Additional backfiles need to process.

def processPackage(row: pd.Series, package_dict: dict, package_name: list, df: pd.DataFrame, idx: int):

    columns_of_subscribe = [col for col in df.columns if row[col] == "*"]
    columns_of_subscribe = list(set(columns_of_subscribe) & set(package_name))
    for col in columns_of_subscribe:
        period = package_dict.get(col).get_order_period()
        df.at[idx, col] = period

def processSubcribe(row: pd.Series, sub_dict: dict, sub_name: list, df: pd.DataFrame, idx: int):
    name = row['Subject Collection']
    if name in sub_name:
        df.at[idx, 'Single Subscribe'] = sub_dict[name].get_order_period()

def to_array(x):
    if isinstance(x, (list, np.ndarray)):  # 如果是列表或陣列，直接用
        return np.array(x)
    if pd.isna(x):  # 如果是 NaN，返回空陣列
        return np.array([])
    return np.array([x])  # 如果是純量，包裝成單元素陣列

# 獨立功能：將年份陣列轉為連續區間
def get_period_ranges(years, start_year, end_year):
    if not years.size:  # 如果陣列為空
        return []
    
    # 排序並去重
    sorted_years = np.unique(years).astype(int)
    start_year = max(sorted_years[0], start_year)  # 取較大的起始年
    end_year = min(sorted_years[-1], end_year)     # 取較小的結束年
    filtered_years = sorted_years[(sorted_years >= start_year) & (sorted_years <= end_year)]
    # 找出連續區間
    period_ranges = []
    if filtered_years.size == 0:
        return []
    start = filtered_years[0]
    prev = start
    for year in filtered_years[1:]:
        if year != prev + 1:  # 不連續時結束一個區間
            if start == prev:
                period_ranges.append(f"{start}")
            else:
                period_ranges.append(f"{start}~{prev}")
            start = year
        prev = year
    # 處理最後一個區間
    if start == prev:
        period_ranges.append(f"{start}")
    else:
        period_ranges.append(f"{start}~{prev}")
    
    return period_ranges

def flatten_array(arr):
    # 如果是純量且為 NaN，返回空陣列
    if not isinstance(arr, (list, np.ndarray)) and pd.isna(arr):
        return np.array([])
    
    # 如果是列表或陣列，扁平化處理
    if isinstance(arr, (list, np.ndarray)):
        flat_list = []
        for item in np.array(arr, dtype=object).flatten():
            if isinstance(item, (list, np.ndarray)):
                flat_list.extend(flatten_array(item))
            else:
                flat_list.append(item)
        return np.array(flat_list)
    
    # 如果是純量（非 NaN），包裝成單元素陣列
    return np.array([arr])



def mergeTheValidRange(row: pd.Series, df:pd.DataFrame, package_name, idx):
    start_year = row['Online date, start']
    end_year = row['Online date, End']
    sub_df = df.loc[idx, package_name]  # Scalar if package_name is str, Series if list
    package_period = sub_df[(sub_df.notna())].values
    subscription_period = df.loc[idx,'Single Subscribe']
    backfiles_period = df.loc[idx,'Backfiles Period']
    sub_arrays = [flatten_array(subscription_period), flatten_array(backfiles_period), flatten_array(package_period)]
    all_years = np.concatenate([arr for arr in sub_arrays if arr.size > 0]) if any(arr.size > 0 for arr in sub_arrays) else np.array([])
    period_ranges = get_period_ranges(all_years, start_year, end_year)
    df.at[idx, 'Subscribe Period'] = period_ranges


p = [
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

master = pd.read_excel('./AIP_Master_List.xlsx')
client = pd.read_excel('./AIP_Client_Pivot.xlsx')

client.dropna(how='any', inplace=True)

client["Sub Year"] = pd.to_numeric(client["Sub Year"], errors="coerce")
client['Product Name'] = client['Product Name'].apply(normalize_text).apply(title_case_with_exceptions)
master['Subject Collection'] = master['Subject Collection'].apply(normalize_text).apply(title_case_with_exceptions)
master['Title'] = master['Title'].apply(normalize_text).apply(title_case_with_exceptions)
master.columns = master.columns.map(normalize_text)

sorted_client = client.sort_values(by=['Product Name', 'Sub Year'], ascending=True)
client_product_name = sorted_client['Product Name'].unique()


backfiles_name =  list(filter(lambda name: "backfiles" in name.lower() or "backfile" in name.lower() ,client_product_name))
package_name = list(filter(lambda name: name in p ,client_product_name))
subscription_name = set(client_product_name) - set(backfiles_name) - set(package_name)

package_dict = createPeriod(package_name, sorted_client, Package)
subscription_dict = createPeriod(subscription_name, sorted_client, Subscription) # return the product's subscription period @np.array
backfile_dict = createPeriod(backfiles_name,sorted_client, Backfile)

master['Backfiles'] =  ""

master['Single Subscribe'] = pd.Series([np.array([]) for _ in range(len(master))])
master['Backfiles Period'] = pd.Series([np.array([]) for _ in range(len(master))])
master['Subscribe Period'] = pd.Series([[] for _ in range(len(master))])


for index, row in master.iterrows():
    # process the package
    processPackage(row=row, package_dict=package_dict, package_name=package_name, df=master, idx=index)
    # process bakckfiles
    processBackfile(row=row, backfiles_dict=backfile_dict, backfile_name=backfiles_name, df=master, idx=index)
    # process the subscription
    processSubcribe(row=row, sub_dict=subscription_dict, sub_name=subscription_name,df=master,idx=index)
    # merge the valid range of subscription
    mergeTheValidRange(row=row, df=master, package_name=package_name, idx=index)


master = master[master['Subscribe Period'].apply(lambda x: not not x)]
cols = master.columns.tolist()
cols.remove('Subscribe Period')
cols.insert(5, 'Subscribe Period')  # 插入到索引 4（第五個位置）
master = master[cols]
master.to_excel('TTT.xlsx')
