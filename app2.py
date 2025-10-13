
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from matplotlib.ticker import ScalarFormatter  # 用於格式化數字
from matplotlib import rcParams  # 用於設置全局字型
import os # 導入 os 模組用於路徑操作
import matplotlib.font_manager as fm # 導入 font_manager 模組

# --- 設置 Matplotlib 字型為支持中文的字型 ---
# 1. 指定字體檔案的路徑
# 根據您提供的截圖，字體檔案名稱為 'NotoSansTC-Regular.ttf'
font_file_name = 'NotoSansTC-Regular.ttf'
font_path = os.path.join(os.path.dirname(__file__), 'fonts', font_file_name)

# 2. 檢查字體檔案是否存在，並添加到 Matplotlib 的字體管理器中
# 根據您提供的預覽，字體的顯示名稱為 'Noto Sans TC'
font_name_for_matplotlib = 'Noto Sans TC'

if os.path.exists(font_path):
    try:
        fm.fontManager.addfont(font_path)
        rcParams['font.family'] = [font_name_for_matplotlib, 'sans-serif']
        # 在 Streamlit 界面顯示成功訊息，方便部署後確認
        st.success(f"成功載入字體: {font_name_for_matplotlib}。")
    except Exception as e:
        # 如果載入字體時發生錯誤，則回退到通用字體並顯示錯誤訊息
        st.error(f"從 {font_path} 載入字體時發生錯誤: {e}。")
        st.warning("字體載入失敗，使用備用字體。")
        rcParams['font.family'] = ['Arial Unicode MS', 'sans-serif'] # 回退到通用字體
else:
    # 如果字體檔案不存在，則回退到通用字體並顯示警告訊息
    st.warning(f"警告: 中文字體檔案 '{font_file_name}' 未在 '{os.path.join('fonts', font_file_name)}' 找到。使用備用字體。")
    rcParams['font.family'] = ['Arial Unicode MS', 'sans-serif'] # 回退到通用字體

# 確保負號 '-' 正常顯示，避免中文環境下顯示為方塊
rcParams['axes.unicode_minus'] = False

# 自訂 CSS 讓頁面靠左對齊並縮小段落間距
def set_page_style():
    st.markdown(
        """
        <style>
            /* 修改 Streamlit 頁面內容容器的樣式 */
            .appview-container .main .block-container {
                max-width: 80%;
                margin-left: 0;
                padding-left: 2rem;
                padding-right: 2rem;
                text-align: left;
            }

            /* 調整段落之間的間距 */
            p, ul, ol {
                margin-top: 0.2rem;
                margin-bottom: 0.2rem;
                line-height: 1.5;
            }

            /* 調整標題與段落之間的間距 */
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
def process_data(data, selected_items, selected_column):
    # 篩選出符合選定項目的資料，並建立副本
    filtered_data = data[data['項目'].isin(selected_items)].copy()

    # 確保 '年月' 欄位為字串格式
    filtered_data['年月'] = filtered_data['年月'].astype(str)

    # 提取年份和月份
    # 假設 '年月' 格式為 'YYYMM'，例如 '11405'
    filtered_data['年份'] = filtered_data['年月'].str[:3]  # 提取前三位作為年份
    filtered_data['月份'] = filtered_data['年月'].str[3:].astype(int)  # 提取後兩位作為月份

    # 只保留必要欄位
    filtered_data = filtered_data[['項目', '年份', '月份', selected_column]]
    filtered_data = filtered_data.sort_values(['項目', '年份', '月份'])  # 按項目、年份和月份排序
    return filtered_data

# 3. 視覺化函數：繪製趨勢圖
def plot_data(data, selected_items, selected_column, separate_by_year, combine_plots):
    if combine_plots:  # 多個項目畫在同一張圖
        plt.figure(figsize=(12, 6))
        if separate_by_year:  # 如果選擇按年度分開繪圖
            for year in data['年份'].unique():
                year_data = data[data['年份'] == year]
                sns.lineplot(x='月份', y=selected_column, data=year_data, marker='o', label=f"{year} 年")
            plt.xlabel("月份", fontsize=12)
            plt.xticks(range(1, 13))  # 確保 X 軸只顯示 1 到 12 月
        else:  # 否則繪製整體趨勢
            full_date_data = pd.DataFrame()
            for item in selected_items:
                item_data_temp = data[data['項目'] == item].copy()
                item_data_temp['年月_label'] = item_data_temp['年份'] + item_data_temp['月份'].astype(str).str.zfill(2)
                full_date_data = pd.concat([full_date_data, item_data_temp])

            sns.lineplot(x='年月_label', y=selected_column, hue='項目', data=full_date_data, marker='o') # 使用hue='項目'來區分不同項目
            plt.xlabel("年月", fontsize=12)
            plt.xticks(rotation=45)

        plt.ylabel("金額 (新台幣)", fontsize=12)

        # 關閉科學記數法並隱藏 Offset Text
        ax = plt.gca()  # 獲取當前的軸對象
        ax.yaxis.set_major_formatter(ScalarFormatter())  # 強制使用完整數字格式
        ax.get_yaxis().get_offset_text().set_visible(False)  # 隱藏 Offset Text

        # 將圖例移到圖外下方
        if separate_by_year: # 如果是按年度分開繪製 (此時圖例是年份)
             plt.legend(
                loc="upper center",
                bbox_to_anchor=(0.5, -0.2),
                ncol=5,
                fontsize=10
            )
        else: # 如果是整體趨勢 (此時圖例是項目)
             plt.legend(
                loc="upper center",
                bbox_to_anchor=(0.5, -0.2),
                ncol=5,
                title="項目",
                fontsize=10
            )


        plt.grid(True)
        st.pyplot(plt)
        plt.close() # 關閉圖形，避免Streamlit重複顯示
    else:  # 每個項目分別畫在不同的圖中
        for item in selected_items:
            plt.figure(figsize=(12, 6))
            item_data = data[data['項目'] == item].copy()
            if separate_by_year:  # 按年度分開繪製
                for year in item_data['年份'].unique():
                    year_data = item_data[item_data['年份'] == year]
                    sns.lineplot(x='月份', y=selected_column, data=year_data, marker='o', label=f"{year} 年")
                plt.xlabel("月份", fontsize=12)
                plt.xticks(range(1, 13))
                plt.legend(
                    loc="upper center",
                    bbox_to_anchor=(0.5, -0.2),
                    ncol=5,
                    fontsize=10
                )
            else:  # 整體趨勢
                item_data['年月_label'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)
                sns.lineplot(x='年月_label', y=selected_column, data=item_data, marker='o', label=item)
                plt.xlabel("年月", fontsize=12)
                plt.xticks(rotation=45)
                plt.legend(
                    loc="upper center",
                    bbox_to_anchor=(0.5, -0.2),
                    ncol=1,
                    fontsize=10
                )

            plt.title(f"{item} 趨勢圖", fontsize=16)
            plt.ylabel("金額 (新台幣)", fontsize=12)

            # 關閉科學記數法並隱藏 Offset Text
            ax = plt.gca()
            ax.yaxis.set_major_formatter(ScalarFormatter())
            ax.get_yaxis().get_offset_text().set_visible(False)

            plt.grid(True)
            st.pyplot(plt)
            plt.close() # 關閉圖形，避免Streamlit重複顯示

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
        selected_items = st.multiselect("請選擇要分析的項目：", item_options, default=[item_options[0]] if item_options else [])

        # 讓使用者選擇分析的列欄位（如 C1、D1 等）
        column_options = data.columns[2:].tolist()  # 動態取得 C1、D1、E1...等欄位名稱
        selected_column = st.selectbox("請選擇要分析的欄位：", column_options)

        # 新增選項：是否按年度分開繪製
        separate_by_year = st.checkbox("是否按年度分開繪製？", value=False)

        # 新增選項：是否將多個項目畫在同一張圖
        combine_plots = st.checkbox("將多個項目畫在同一張圖？", value=True)

        # 資料處理
        if selected_items and selected_column:
            filtered_data = process_data(data, selected_items, selected_column)

            # 如果有資料，顯示圖表
            if not filtered_data.empty:
                plot_data(filtered_data, selected_items, selected_column, separate_by_year, combine_plots)
            else:
                st.write("目前尚無符合的資料，請重新選擇項目或欄位。")
        else:
            st.write("請選擇至少一個項目和一個欄位以進行分析。")
    else:
        st.write("請上傳 Excel 檔案以進行分析。")

if __name__ == "__main__":
    main()
