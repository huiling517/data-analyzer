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
font_file_name = 'NotoSansTC-Regular.ttf'
font_path = os.path.join(os.path.dirname(__file__), 'fonts', font_file_name)
font_name_for_matplotlib = 'Noto Sans TC'

if os.path.exists(font_path):
    try:
        fm.fontManager.addfont(font_path)
        rcParams['font.family'] = [font_name_for_matplotlib, 'sans-serif']
    except Exception as e:
        st.error(f"從 {font_path} 載入字體時發生錯誤: {e}。")
        rcParams['font.family'] = ['Arial Unicode MS', 'sans-serif']
else:
    st.warning(f"警告: 中文字體檔案 '{font_file_name}' 未找到，使用備用字體。")
    rcParams['font.family'] = ['Arial Unicode MS', 'sans-serif']

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


# 1. 讀取 Excel 檔案並允許選擇工作表
@st.cache_data
def load_data(uploaded_file, sheet_name):
    data = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    return data


# 2. 資料處理：適配五碼和六碼的年月格式
def process_data(data, selected_items, selected_columns):
    filtered_data = data[data['項目'].isin(selected_items)].copy()
    filtered_data['年月'] = filtered_data['年月'].astype(str)

    # 適配五碼或六碼的年月格式
    def parse_year_month(ym):
        if len(ym) == 5:  # 五碼格式，形如 20191
            return ym[:3], int(ym[3:])  # 前三位為年份，後兩位為月份
        elif len(ym) == 6:  # 六碼格式，形如 201901
            return ym[:4], int(ym[4:])  # 前四位為年份，後兩位為月份
        else:
            return None, None  # 如果格式不符，返回空值

    filtered_data['年份'], filtered_data['月份'] = zip(*filtered_data['年月'].map(parse_year_month))

    # 移除解析失敗的行
    filtered_data = filtered_data.dropna(subset=['年份', '月份'])

    # 只保留必要欄位
    columns_to_keep = ['項目', '年份', '月份'] + selected_columns
    filtered_data = filtered_data[columns_to_keep]
    filtered_data = filtered_data.sort_values(['項目', '年份', '月份'])
    return filtered_data


# 3. 視覺化函數：繪製趨勢圖
def plot_data(data, selected_items, selected_columns, separate_by_year, combine_plots, custom_labels):
    x_label, y_label = custom_labels

    if combine_plots:
        plt.figure(figsize=(12, 6))
        for selected_column in selected_columns:
            for item in selected_items:
                item_data = data[data['項目'] == item].copy()
                item_data['年月'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)

                if separate_by_year:
                    for year in item_data['年份'].unique():
                        year_data = item_data[item_data['年份'] == year]
                        sns.lineplot(x='月份', y=selected_column, data=year_data, marker='o', label=f"{item} - {selected_column} ({year}年)")
                    plt.xlabel(x_label, fontsize=12)
                    plt.xticks(range(1, 13))
                else:
                    sns.lineplot(x='年月', y=selected_column, data=item_data, marker='o', label=f"{item} - {selected_column}")
                    plt.xlabel(x_label, fontsize=12)
                    plt.xticks(rotation=45)

        plt.ylabel(y_label, fontsize=12)
        plt.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=3, fontsize=10)
        plt.grid(True)
        st.pyplot(plt)

    else:
        for selected_column in selected_columns:
            for item in selected_items:
                plt.figure(figsize=(12, 6))
                item_data = data[data['項目'] == item].copy()

                if separate_by_year:
                    for year in item_data['年份'].unique():
                        year_data = item_data[item_data['年份'] == year]
                        sns.lineplot(
                            x='月份',
                            y=selected_column,
                            data=year_data,
                            marker='o',
                            label=f"{year}年"
                        )
                    plt.xlabel(x_label, fontsize=12)
                    plt.xticks(range(1, 13))
                    plt.title(f"{item} - {selected_column} 趨勢圖", fontsize=16)
                else:
                    item_data['年月'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)
                    sns.lineplot(x='年月', y=selected_column, data=item_data, marker='o', label=f"{item}")
                    plt.xlabel(x_label, fontsize=12)
                    plt.xticks(rotation=45)

                plt.ylabel(y_label, fontsize=12)
                plt.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=5, fontsize=10)
                plt.grid(True)
                st.pyplot(plt)


# 4. Streamlit App 主程式
def main():
    set_page_style()

    st.title("數據分析工具")
    st.write("1. 上傳一個 Excel 檔案，並選擇一個或多個項目和欄位，分析其在不同月份或年度的趨勢變化")
    st.write("2. 欄位規則：A欄為年月，A1請填入文字:年月；B欄為分析項目，B1請填入文字:分析項目；C欄以後請自行定義名稱")
    st.write("3. A欄為年月，A1之後欄位請填入文字，例如114年5月，請填入文字:11405 或 201905")
    st.write("4. B欄為分析項目，A2之後請填入文字，例如：醫療收入或人事費用...")
    st.write("5. C2往右及往下儲存格為數值資料")

    uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx"], label_visibility="hidden")

    if uploaded_file is not None:
        # 獲取工作表名稱
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        selected_sheet = st.selectbox("請選擇工作表：", sheet_names)

        # 載入選定的工作表數據
        data = load_data(uploaded_file, selected_sheet)

        if '項目' not in data.columns or '年月' not in data.columns:
            st.error("上傳的工作表中缺少必要欄位：'項目' 或 '年月'，請檢查數據格式。")
            return

        item_options = data['項目'].unique().tolist()
        selected_items = st.multiselect("請選擇要分析的項目：", item_options, default=[item_options[0]] if item_options else [])

        column_options = data.columns[2:].tolist()
        selected_columns = st.multiselect("請選擇要分析的欄位（可以多選）：", column_options, default=[column_options[0]] if column_options else [])

        separate_by_year = st.checkbox("是否按年度分開繪製？", value=False)
        combine_plots = st.checkbox("將多個項目畫在同一張圖？", value=True)

        # 新增選項：自定義 X 軸與 Y 軸標籤
        x_label = st.text_input("請輸入 X 軸標籤：", value="月份")
        y_label = st.text_input("請輸入 Y 軸標籤：", value="金額 (新台幣)")

        if selected_items and selected_columns:
            filtered_data = process_data(data, selected_items, selected_columns)

            if not filtered_data.empty:
                plot_data(filtered_data, selected_items, selected_columns, separate_by_year, combine_plots, (x_label, y_label))
            else:
                st.write("目前尚無符合的資料，請重新選擇項目或欄位。")
        else:
            st.write("請選擇至少一個項目和一個欄位以進行分析。")
    else:
        st.write("請上傳 Excel 檔案以進行分析。")


if __name__ == "__main__":
    main()
