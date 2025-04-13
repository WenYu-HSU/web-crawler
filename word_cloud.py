import pandas as pd
import jieba
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import emoji
import re
import os

def analyze_excel_file(excel_path=r"C:\Users\user\OneDrive\桌面\bullying 1.xlsx", column_name="content", font_path=r"C:\Windows\Fonts\msjh.ttc"):
    """分析Excel文件並生成結果"""
    # 顯示當前工作目錄
    current_dir = os.getcwd()
    print(f"\n當前工作目錄：{current_dir}")
    
    try:
        # 讀取Excel檔案
        print(f"\n正在讀取Excel檔案：{excel_path}")
        df = pd.read_excel(excel_path)
        
        # 文本清理和分詞
        texts = df[column_name].dropna()
        word_freq = Counter()

        # 為每一行建立一個關鍵詞列表
        df['關鍵詞列表'] = ''
        df['關鍵詞頻率'] = ''
        
        # 處理每一行文本
        for idx, text in enumerate(texts):
            if pd.isna(text):
                continue
        
            # 清理文本
            text = emoji.replace_emoji(str(text), '')
            text = re.sub(r'http\S+|www.\S+', '', text)
            text = re.sub(r'@\S+|#\S+', '', text)
            text = re.sub(r'[^\w\s\u4e00-\u9fff]', '', text)
            
            # 分詞
            words = jieba.cut(text,cut_all=False)
            
            # 過濾停用詞
            stopwords = {'的', '了', '是', '在', '我', '有', '和', '就', '不', '人', 'http', 'https','翻譯','自己'}
            words = [w for w in words if w not in stopwords and len(w) > 1]

            # 計算當前行的詞頻
            row_word_freq = Counter(words)
            
            word_freq.update(words)

            # 將該行的關鍵詞和頻率存入DataFrame
            df.at[idx, '關鍵詞列表'] = '、'.join(row_word_freq.keys())
            df.at[idx, '關鍵詞頻率'] = '、'.join(f'{word}({count})' for word, count in row_word_freq.most_common())
        
        if not word_freq:
            print("沒有找到有效的文字進行分析！")
            return None, None, None
        
        # 生成文字雲
        wordcloud = WordCloud(
            font_path=font_path,
            width=800,
            height=400,
            background_color='white',
            max_words=200
        )
        
        wordcloud.generate_from_frequencies(word_freq)
        
        # 儲存結果
        output_file = f'wordcloud_{column_name}_分析結果.png'
        wordcloud.to_file(output_file)
        
        # 顯示文字雲
        plt.figure(figsize=(12, 6))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.show()
        
        # 將整體詞頻統計添加到新工作表
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # 儲存原始資料（包含新增的關鍵詞列表和頻率）
            df.to_excel(writer, sheet_name='原始資料', index=False)
            
            # 建立詞頻統計工作表
            word_freq_sorted = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
            df_word_freq = pd.DataFrame(word_freq_sorted, columns=['關鍵詞', '出現次數'])
            df_word_freq.to_excel(writer, sheet_name='詞頻統計', index=False)
            
            # 建立摘要統計工作表
            summary_data = {
                '統計項目': ['總詞數', '不重複詞數', '分析文本數'],
                '數值': [sum(word_freq.values()), len(word_freq), len(texts)]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='摘要統計', index=False)
        
        print("\n分析完成！結果已更新到原始Excel檔案中：")
        print(f"1. 原始資料表：新增了「關鍵詞列表」和「關鍵詞頻率」欄位")
        print(f"2. 詞頻統計表：包含所有關鍵詞的出現次數")
        print(f"3. 摘要統計表：包含整體統計資訊")
        print(f"4. 文字雲圖片已儲存至：{os.path.join(current_dir, output_file)}")
        
        # 顯示前20個關鍵詞
        print("\n最常出現的關鍵詞及次數：")
        for word, count in word_freq_sorted[:20]:
            print(f"{word}: {count}")
            
        return df
        
    except Exception as e:
        print(f"發生錯誤：{str(e)}")
        return None

# 呼叫函數
if __name__ == "__main__":
    result_df = analyze_excel_file()