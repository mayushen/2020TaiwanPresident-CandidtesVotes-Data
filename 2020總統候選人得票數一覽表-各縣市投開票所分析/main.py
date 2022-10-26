import pandas as pd
from string import ascii_uppercase


def get_tidy_data(file_path):
    # 讀取excel檔案路徑
    data = pd.read_excel(file_path, skiprows=[0, 1, 3, 4])

    # 擷取候選人人數與欄位資訊
    column_names = data.columns
    n_candidates = len(column_names) - 11
    candidates = list(column_names[3: 3 + n_candidates])

    # 重置欄位名稱、填補行政區之NaN遺漏值
    columns = ['district', 'village', 'office'] + candidates + list(ascii_uppercase[: 8])
    data.columns = columns
    data['district'] = data['district'].fillna(method='ffill')

    # 清除「district」行政區欄位的空字串
    data['district'] = data['district'].str.replace('\u3000', '').str.strip()

    # 清除總計、行政區小總計的水平列
    data = data.dropna(axis=0).reset_index(drop=True)
    data = data.drop(list(ascii_uppercase[:8]), axis=1)

    # 將候選人欄位轉製至rows中
    candidates_info = data.columns[3: 3 + n_candidates]
    formatted_data = pd.melt(data, id_vars=['district', 'village', 'office'],
                             value_vars=candidates_info, var_name="candidates", value_name="votes")

    return formatted_data


admin_areas = ["臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市", "新竹縣", "苗栗縣", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
               "臺東縣", "澎湖縣", "基隆市", "新竹市", "嘉義市", "金門縣", "連江縣"]
file_paths = [
    f"/Users/mayushen0406/PycharmProjects/2020總統候選人得票數一覽表-各縣市投開票所分析/總統-各投票所得票明細及概況(Excel檔)/" \
    f"總統-A05-4-候選人得票數一覽表-各投開票所({area_name}).xls" for area_name in admin_areas
]
df_dict = {}
votes_data_final = pd.DataFrame()
for path, area in zip(file_paths, admin_areas):
    tidy_df = get_tidy_data(path)
    # 加入area「縣市」資料至新的欄位裡
    tidy_df["admin_area"] = area
    df_dict[area] = tidy_df
    # 將整理好的Tidy Data加入至votes_data_final資料匡裡
    votes_data_final = pd.concat([votes_data_final, tidy_df],
                                 axis=0, ignore_index=True)
    print('現在正在處理{}的資料......'.format(area))
    print('資料的大小為', tidy_df.shape)

# 將原candidates欄位資訊分別整理成「party_number」與「candidates」兩欄，並新增party欄位
can_name_info = votes_data_final['candidates'].str.split('\n', expand=True)
votes_data_final['party_number'] = can_name_info[0].str.replace('\(', '').str.replace('\)', '')
votes_data_final['candidates'] = can_name_info[1].str.cat(can_name_info[2], '/')
party_dict = {'1': '親民黨',
              '2': '中國國民黨',
              '3': '民主進步黨'}
votes_data_final['party'] = votes_data_final['party_number'].map(party_dict)

# 將final_data的資料型別轉換成適合的type
votes_data_final['party_number'] = votes_data_final['party_number'].astype(int)
votes_data_final['office'] = votes_data_final['office'].astype(int)
votes_data_final['votes'] = votes_data_final['votes'].fillna(0)
votes_data_final['votes'] = votes_data_final['votes'].astype(str).str.replace(',', '')
votes_data_final['votes'] = votes_data_final['votes'].astype(int)

print(votes_data_final.head(15))

# 確認最終資料無誤後，將DataFrame轉換成csv檔
votes_data_final.to_csv('2020年台灣總統候選人各開票所之得票數一覽表.csv', index=False)
