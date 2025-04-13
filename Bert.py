# 匯入必要套件
import pandas as pd  
from transformers import BertTokenizer, BertForSequenceClassification  
import torch 
from torch.nn.functional import softmax  
import warnings  
warnings.filterwarnings('ignore') 

# 載入模型 
def load_model():
    """
    載入預訓練的中文 BERT 模型與對應的 tokenizer
    模型為二分類情緒分析（0 = 負面，1 = 正面）
    """
    model_name = "bert-base-chinese"
    tokenizer = BertTokenizer.from_pretrained(model_name)
    model = BertForSequenceClassification.from_pretrained(model_name, num_labels=2)
    return tokenizer, model

#  對單一文本做情感分析 
def analyze_sentiment(text, tokenizer, model):
    """
    使用 BERT 模型對單句文字做情緒分析，輸出正面機率
    """
    # 將文字轉換為模型輸入格式
    inputs = tokenizer(
        text,
        return_tensors="pt",    
        truncation=True,         # 超過最大長度時截斷
        max_length=512,
        padding=True             # 補齊不足長度
    )

    # 執行模型預測
    outputs = model(**inputs)

    # 將 raw logits 轉換為機率分佈
    probs = softmax(outputs.logits, dim=1)

    # 取得「正面」這一類的機率值
    positive_score = probs[0][1].item()

    return positive_score  # 回傳正面情緒機率（0~1）

# 批量處理 Excel 檔案 
def process_excel(input_file, output_file, text_column):
    """
    載入 Excel 檔案，針對指定欄位進行情感分析，並存回結果
    """
    print("載入模型中...")
    tokenizer, model = load_model()

    print(f"讀取文件: {input_file}")
    df = pd.read_excel(input_file)

    print("進行情緒分析...")
    sentiment_scores = []  # 用來儲存每一列的結果
    total = len(df)

    # 遍歷每一列
    for idx, row in df.iterrows():
        text = str(row[text_column])  # 讀取指定欄位
        score = analyze_sentiment(text, tokenizer, model)
        sentiment_scores.append(score)

        # 顯示進度每10筆印一次
        if (idx + 1) % 10 == 0:
            print(f"已處理 {idx + 1}/{total} 條數據")

    # 加入新的一欄：情緒分數（正面機率）
    df['sentiment_score'] = sentiment_scores

    print(f"保存結果到: {output_file}")
    df.to_excel(output_file, index=False)
    print("處理完成！")

# 主程式執行區 
if __name__ == "__main__":
    input_file = r"C:\Users\user\OneDrive\桌面\emotion.xlsx"  # 輸入檔案路徑
    output_file = r"C:\Users\user\OneDrive\桌面\threads_analysis.xlsx"  # 輸出檔案路徑
    text_column = "content"  # 欲分析的欄位名稱

    process_excel(input_file, output_file, text_column)
