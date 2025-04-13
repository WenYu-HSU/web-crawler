import pandas as pd
import jieba
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import emoji
import re
import os

def analyze_excel_file(
    excel_path=r"C:\Users\msi\Desktop\霸凌負面.xlsx",
    column_name="content",
    font_path=r"C:\Windows\Fonts\msjh.ttc"
):
    """分析 Excel 文件並生成詞頻分析與文字雲"""

    # 顯示當前工作目錄
    current_dir = os.getcwd()
    print(f"\n當前工作目錄：{current_dir}")

    try:
        print(f"\n正在讀取 Excel 檔案：{excel_path}")
        df = pd.read_excel(excel_path)

        # 文本清理與初始化
        texts = df[column_name].dropna()
        word_freq = Counter()

        df['關鍵詞列表'] = ''
        df['關鍵詞頻率'] = ''

        # 停用詞
        stopwords = {'的', '了', '是', '在', '我', '有', '和', '就', '不', '人', 'http', 'https', '翻譯'}

        # 分析每一列
        for idx, text in enumerate(texts):
            if pd.isna(text):
                continue

            # 清理文字：去 emoji、連結、@、#、非中英文字符
            text = emoji.replace_emoji(str(text), '')
            text = re.sub(r'http\S+|www.\S+', '', text)
            text = re.sub(r'@\S+|#\S+', '', text)
            text = re.sub(r'[^\w\s\u4e00-\u9fff]', '', text)

            # 分詞
            words = jieba.cut(text)
            words = [w for w in words if w not in stopwords and len(w) > 1]

            # 詞頻統計
            row_word_freq = Counter(words)
            word_freq.update(words)

            # 寫入當前行詞頻結果
            df.at[idx, '關鍵詞列表'] = '、'.join(row_word_freq.keys())
            df.at[idx, '關鍵詞頻率'] = '、'.join(f'{word}({count})' for word, count in row_word_freq.most_common())

        if not word_freq:
            print("⚠️ 沒有找到有效的文字進行分析")
            return None

        # 建立文字雲
        wordcloud = WordCloud(
            font_path=font_path,
            width=800,
            height=400,
            background_color='white',
            max_words=200
        ).generate_from_frequencies(word_freq)

        output_file = f'wordcloud_{column_name}_分析結果.png'
        wordcloud.to_file(output_file)

        # 顯示圖形
        plt.figure(figsize=(12, 6))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.show()

        # 儲存 Excel 結果
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='原始資料', index=False)

            # 詞頻表
            word_freq_sorted = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
            df_word_freq = pd.DataFrame(word_freq_sorted, columns=['關鍵詞', '出現次數'])
            df_word_freq.to_excel(writer, sheet_name='詞頻統計', index=False)

            # 摘要統計
            summary_data = {
                '統計項目': ['總詞數', '不重複詞數', '分析文本數'],
                '數值': [sum(word_freq.values()), len(word_freq), len(texts)]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='摘要統計', index=False)

        print("\n✅ 分析完成！結果已更新至原始 Excel 檔案中：")
        print(f"🔸 原始資料表：新增了「關鍵詞列表」與「關鍵詞頻率」欄位")
        print(f"🔸 詞頻統計表：包含所有關鍵詞與次數")
        print(f"🔸 摘要統計表：詞彙量、文本數")
        print(f"🔸 文字雲圖片儲存於：{os.path.join(current_dir, output_file)}")

        # 印出前 50 熱門關鍵詞
        print("\n📊 最常出現的關鍵詞 Top 50：")
        for word, count in word_freq_sorted[:50]:
            print(f"{word}: {count}")

        return df

    except Exception as e:
        print(f"❌ 發生錯誤：{str(e)}")
        return None

# 執行主流程
if __name__ == "__main__":
    result_df = analyze_excel_file()
