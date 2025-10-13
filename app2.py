import streamlit as st
import matplotlib.pyplot as plt
from matplotlib import rcParams

# 設定 Matplotlib 字型
def set_chinese_font():
    rcParams['font.family'] = ['SimHei', 'Noto Sans CJK SC', 'Arial Unicode MS']  # 嘗試常見中文字型
    rcParams['axes.unicode_minus'] = False  # 解決負號顯示為方塊的問題

# 主程式
def main():
    set_chinese_font()  # 設置中文字型

    st.title("中文顯示測試")
    st.write("這是一個測試應用，確認 Streamlit Cloud 是否正確顯示中文。")

    # 繪製包含中文的圖表
    fig, ax = plt.subplots()
    ax.plot([1, 2, 3], [4, 5, 6])
    ax.set_title("中文標題測試")
    ax.set_xlabel("月份")
    ax.set_ylabel("金額 (新台幣)")
    st.pyplot(fig)

if __name__ == "__main__":
    main()
