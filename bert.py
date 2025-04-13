import pandas as pd
from transformers import BertTokenizer, BertForSequenceClassification
import torch
from torch.nn.functional import softmax
import warnings
warnings.filterwarnings('ignore')

def load_model():
    """
    載入預訓練的中文BERT模型
    """
    # 使用預訓練的中文BERT模型
    model_name = "bert-base-chinese"
    tokenizer = BertTokenizer.from_pretrained(model_name)
    model = BertForSequenceClassification.from_pretrained(model_name, num_labels=2)
    return tokenizer, model

def analyze_sentiment(text, tokenizer, model):
    """
    對單一文本進行情緒分析
    """
    # 預處理文本
    inputs = tokenizer(text, 
                      return_tensors="pt",
                      truncation=True,
                      max_length=512,
                      padding=True)
    
    # 進行預測
    outputs = model(**inputs)
    probs = softmax(outputs.logits, dim=1)
    
    # 獲取正面情緒的概率（0-1之間）
    positive_score = probs[0][1].item()
    
    return positive_score

def process_excel(input_file, output_file, text_column):
    """
    處理Excel文件並保存結果
    """
    # 載入模型
    print("載入模型中...")
    tokenizer, model = load_model()
    
    # 讀取Excel文件
    print(f"讀取文件: {input_file}")
    df = pd.read_excel(input_file)
    
    # 進行情緒分析
    print("進行情緒分析...")
    sentiment_scores = []
    total = len(df)
    
    for idx, row in df.iterrows():
        text = str(row[text_column])
        score = analyze_sentiment(text, tokenizer, model)
        sentiment_scores.append(score)
        
        # 顯示進度
        if (idx + 1) % 10 == 0:
            print(f"已處理 {idx + 1}/{total} 條數據")
    
    # 添加結果列
    df['sentiment_score'] = sentiment_scores
    
    # 保存結果
    print(f"保存結果到: {output_file}")
    df.to_excel(output_file, index=False)
    print("處理完成！")

# 使用示例
if __name__ == "__main__":
    input_file = r"C:\Users\user\OneDrive\桌面\emotion.xlsx"  # 輸入文件名
    output_file = r"C:\Users\user\OneDrive\桌面\threads_analysis.xlsx"  # 輸出文件名
    text_column = "content"  # 要分析的文本列名
    
    process_excel(input_file, output_file, text_column)