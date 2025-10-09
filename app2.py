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
            for item in selected_items:
                # 篩選出對應項目的數據
                item_data = data[data['項目'] == item]
                item_data['年月'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)  # 合併年份和月份
                sns.lineplot(x='年月', y=selected_column, data=item_data, marker='o', label=item)  # 傳遞項目的名稱
            plt.xlabel("年月", fontsize=12)
            plt.xticks(rotation=45)

        plt.ylabel("金額 (新台幣)", fontsize=12)

        # 關閉科學記數法並隱藏 Offset Text
        ax = plt.gca()  # 獲取當前的軸對象
        ax.yaxis.set_major_formatter(ScalarFormatter())  # 強制使用完整數字格式
        ax.get_yaxis().get_offset_text().set_visible(False)  # 隱藏 Offset Text

        # 將圖例移到圖外下方，並不顯示「項目」
        plt.legend(
            loc="upper center",  # 放置在圖形下方
            bbox_to_anchor=(0.5, -0.2),  # 調整位置（水平置中，垂直向下）
            ncol=5,  # 每行顯示 5 個圖例
            fontsize=10
        )

        plt.grid(True)
        st.pyplot(plt)
    else:  # 每個項目分別畫在不同的圖中
        for item in selected_items:
            plt.figure(figsize=(12, 6))
            item_data = data[data['項目'] == item]
            if separate_by_year:  # 按年度分開繪製
                for year in item_data['年份'].unique():
                    year_data = item_data[item_data['年份'] == year]
                    sns.lineplot(x='月份', y=selected_column, data=year_data, marker='o', label=f"{year} 年")
                plt.xlabel("月份", fontsize=12)
                plt.xticks(range(1, 13))
            else:  # 整體趨勢
                item_data['年月'] = item_data['年份'] + item_data['月份'].astype(str).str.zfill(2)
                sns.lineplot(x='年月', y=selected_column, data=item_data, marker='o', label=item)  # 傳遞項目的名稱
                plt.xlabel("年月", fontsize=12)
                plt.xticks(rotation=45)

            plt.title(f"{item} 趨勢圖", fontsize=16)
            plt.ylabel("金額 (新台幣)", fontsize=12)

            # 關閉科學記數法並隱藏 Offset Text
            ax = plt.gca()
            ax.yaxis.set_major_formatter(ScalarFormatter())
            ax.get_yaxis().get_offset_text().set_visible(False)

            plt.grid(True)
            st.pyplot(plt)
