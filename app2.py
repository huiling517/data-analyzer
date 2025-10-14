import warnings  # 用於隱藏警告
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.ticker import ScalarFormatter  # 用於格式化數字
from matplotlib import rcParams  # 用於設置全局字型
import os  # 用於路徑操作
import matplotlib.font_manager as fm  # 用於字體管理

# 隱藏 Matplotlib 圖例相關的警告
warnings.filterwarnings("ignore", message="No artists with labels found to put in legend")

# --- 字體設定：載入中文字體 ---
# 指定字體檔案名稱
font_file_name = 'NotoSansTC-Regular.ttf'
font_path = os.path.join(os.path.dirname(__file__), 'fonts', font_file_name)  # 假設字體檔案位於 `fonts` 資料夾
font_name_for_matplotlib = 'Noto Sans TC'

# 檢查字體檔案是否存在
if os.path.exists(font_path):
    try:
        fm.fontManager.addfont(font_path)
        rcParams['font.family'] = [font_name_for_matplotlib, 'sans-serif']  # 設置 Matplotlib 的全局字型
    except Exception as e:
        st.error(f"從 {font_path} 載入字體時發生錯誤: {e}。")
        rcParams['font.family'] = ['Arial Unicode MS', 'sans-serif']  # 回退到通用字體
else:
    st.warning(f"警告: 中文字體檔案 '{font_file_name}' 未在 '{os.path.join('fonts', font_file_name)}' 找到。使用備用字體。")
    rcParams['font.family'] = ['Arial Unicode MS', 'sans-serif']  # 回退到通用字體

# 確保負號 '-' 正常顯示
rcParams['axes.unicode_minus'] = False

# 自訂 CSS 讓頁面靠左對齊並縮小段落間距
def set_page_style():
    st.markdown(
        """
        <style>
            .appview-container .main .block-container {
                max-width: 80%;
                margin-left: 0;
                padding-left: 2rem;
                padding-right: 2rem;
                text-align: left;
            }
            p, ul, ol {
                margin-top: 0.2rem;
                margin-bottom: 0.2rem;
                line-height: 1.5;
            }
            h1, h2, h3 {
                margin-top: 0.5rem;
                margin-bottom: 0.5rem;
            }
        </style>
        """,
        unsafe_allow_html=True
    )

# 1. 讀取 Excel 檔案
@st.cache_data
def load_data(uploaded_file):
    data = pd.read_excel(uploaded_file, sheet_name=0)  # 預設讀取第一個工作表
    return data

# 2. 資料處理：篩選出年月和指定欄位
def process_data(data, selected_items, selected_columns):
    # 篩選出符合選定項目的資料，並建立副本
    filtered_data = data[data['項目'].isin(selected_items)].copy()

    # 確保 '年月' 欄位為字串格式
    filtered_data['年月'] = filtered_data['年月'].astype(str)

    # 提取年份和月份
    filtered_data['年份'] = filtered_data['年月'].str[:3]  # 提取前三位作為年份
    filtered_data['月份'] = filtered_data['年月'].str[3:].astype(int)  # 提取後兩位作為月份

    # 只保留必要欄位
    columns_to_keep = ['項目', '年份', '月份'] + selected_columns
    filtered_data = filtered_data[columns_to_keep]
    filtered_data = filtered_data.sort_values(['項目', '年份', '月份'])  # 按項目、年份和月份排序
    return filtered_data

# 3. 視覺化函數：繪製趨勢圖
def plot_data(data, selected_items, selected_columns, separate_by_year, combine_plots):
    if combine_plots:  # 多個項目畫在同一張圖
        plt.figure(figsize=(12, 6))
        for selected_column in selected_columns:  # 迴圈處理每個欄位
            for item in selected_items:  # 迴圈處理每個項目
                # 篩選出對應項目和欄位的數據
                item_data = data[data['項目'] == item].copy()
                item_data['年月'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)  # 合併年份和月份

                if separate_by_year:  # 按年度分開繪製
                    for year in item_data['年份'].unique():  # 針對每個年份分別繪製
                        year_data = item_data[item_data['年份'] == year]
                        sns.lineplot(
                            x='月份',
                            y=selected_column,
                            data=year_data,
                            marker='o',
                            label=f"{item} - {selected_column} ({year}年)"
                        )
                    plt.xlabel("月份", fontsize=12)
                    plt.xticks(range(1, 13))  # 確保 X 軸只顯示 1 到 12 月
                else:  # 繪製整體趨勢
                    sns.lineplot(x='年月', y=selected_column, data=item_data, marker='o', label=f"{item} - {selected_column}")
                    plt.xlabel("年月", fontsize=12)
                    plt.xticks(rotation=45)

        plt.ylabel("金額 (新台幣)", fontsize=12)
        plt.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, -0.2),
            ncol=3,
            fontsize=10
        )
        plt.grid(True)
        st.pyplot(plt)

    else:  # 每個項目分別畫在不同的圖中
        for selected_column in selected_columns:  # 迴圈處理每個欄位
            for item in selected_items:  # 迴圈處理每個項目
                plt.figure(figsize=(12, 6))
                item_data = data[data['項目'] == item].copy()

                if separate_by_year:  # 按年度分開繪製
                    for year in item_data['年份'].unique():  # 針對每個年份分別繪製
                        year_data = item_data[item_data['年份'] == year]
                        sns.lineplot(
                            x='月份',
                            y=selected_column,
                            data=year_data,
                            marker='o',
                            label=f"{year}年"
                        )
                    plt.xlabel("月份", fontsize=12)
                    plt.xticks(range(1, 13))
                else:  # 整體趨勢
                    item_data['年月'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)
                    sns.lineplot(x='年月', y=selected_column, data=item_data, marker='o', label=f"{item}")
                    plt.xlabel("年月", fontsize=12)
                    plt.xticks(rotation=45)

                plt.title(f"{item} - {selected_column} 趨勢圖", fontsize=16)
                plt.ylabel("金額 (新台幣)", fontsize=12)
                plt.legend(
                    loc="upper center",
                    bbox_to_anchor=(0.5, -0.2),
                    ncol=5,
                    fontsize=10
                )
                plt.grid(True)
                st.pyplot(plt)

# 4. Streamlit App 主程式
def main():
    set_page_style()  # 設置頁面樣式

    st.title("數據分析工具")
    st.write("1. 上傳一個 Excel 檔案，並選擇一個或多個項目和欄位，分析其在不同月份或年度的趨勢變化")
    st.write("2. 欄位規則：A欄為年月，A1請填入文字:年月；B欄為分析項目，B1請填入文字:分析項目；C欄以後請自行定義名稱")
    st.write("3. A欄為年月，A1之後欄位請填入文字，例如114年5月，請填入文字:11405")
    st.write("4. B欄為分析項目，A2之後欄位請填入文字，例如：醫療收入或人事費用...")
    st.write("5. C2往右及往下儲存格為數值資料")

    # 讓使用者上傳 Excel 檔案
    uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx"], label_visibility="hidden")

    if uploaded_file is not None:
        # 載入資料數據
        data = load_data(uploaded_file)

        # 讓使用者選擇分析的「項目」
        item_options = data['項目'].unique().tolist()  # 動態取得項目選項
        selected_items = st.multiselect("請選擇要分析的項目：", item_options, default=[item_options[0]])

        # 讓使用者選擇分析的列欄位（多選）
        column_options = data.columns[2:].tolist()  # 動態取得 C1、D1、E1...等欄位名稱
        selected_columns = st.multiselect(
            "請選擇要分析的欄位（可以多選）：", column_options, default=[column_options[0]]
        )

        # 新增選項：是否按年度分開繪製
        separate_by_year = st.checkbox("是否按年度分開繪製？", value=False)

        # 新增選項：是否將多個項目畫在同一張圖
        combine_plots = st.checkbox("將多個項目畫在同一張圖？", value=True)

        # 資料處理
        filtered_data = process_data(data, selected_items, selected_columns)

        # 如果有資料，顯示圖表
        if not filtered_data.empty:
            plot_data(filtered_data, selected_items, selected_columns, separate_by_year, combine_plots)
        else:
            st.write("目前尚無符合的資料，請重新選擇項目或欄位。")
    else:
        st.write("請上傳 Excel 檔案以進行分析。")

if __name__ == "__main__":
    main()

