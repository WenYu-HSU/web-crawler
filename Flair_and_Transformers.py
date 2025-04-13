# 匯入所需套件
import openpyxl  
from flair.data import Sentence  
from flair.models import TextClassifier 
from transformers import pipeline 

# 開啟 Excel 檔案 
# 設定 Excel 檔案路徑
file_path = r'/Users/p1ayer0ne/threads.xlsx'

# 使用 openpyxl 載入 Excel 檔案
wb = openpyxl.load_workbook(file_path)

# 選擇要操作的工作表
sheet = wb['Sheet1']


# 初始化兩個情感分析模型 
# 初始化 transformers 的多語言情感分析模型
transformers_classifier = pipeline(
    "sentiment-analysis",
    model="nlptown/bert-base-multilingual-uncased-sentiment",
    top_k=None  # 回傳所有 1~5 星的預測分數
)

# 初始化 Flair 的英文情感分類器
flair_classifier = TextClassifier.load('en-sentiment')

# 遍歷 Excel 的每一列進行情感分析 
# 從第 2 列開始
for row in range(2, sheet.max_row + 1):

    # 讀取第 C 欄
    c_cell = sheet[f'C{row}']

    # Flair 結果寫入 I 欄
    i_cell = sheet[f'I{row}']

    # transformers 結果寫入 J 欄
    j_cell = sheet[f'J{row}']

    # 取得儲存格內容
    c_text = c_cell.value

    # 確保有文字可分析
    if c_text:

        # 使用 transformers 進行情感分析
        transformers_result = transformers_classifier(c_text)

        # 建立星級與數值的對應（1~5星）
        star_mapping = {
            "1 star": 1,
            "2 stars": 2,
            "3 stars": 3,
            "4 stars": 4,
            "5 stars": 5
        }

        # 加權計算每個星級的分數，取出第 0 組預測結果
        weighted_score = sum(
            star_mapping[res['label']] * res['score']
            for res in transformers_result[0]
        )

        # 將分數標準化為 0 ~ 1（1 星 = 0.0，5 星 = 1.0）
        transformers_normalized = (weighted_score - 1) / 4

        # 使用 Flair 進行情感分析
        sentence = Sentence(c_text)
        flair_classifier.predict(sentence)

        # Flair 回傳標籤：POSITIVE 或 NEGATIVE
        flair_label = sentence.labels[0].value

        # Flair 回傳信心分數（越接近 1 越有信心）
        flair_score = sentence.labels[0].score

        # 將 Flair 結果轉換成 0~1 分數：POSITIVE = 分數本身；NEGATIVE = 1 - 分數
        flair_normalized = flair_score if flair_label == "POSITIVE" else (1 - flair_score)

        # 寫入結果到 Excel
        i_cell.value = f"{flair_label} ({flair_normalized:.2f})"
        j_cell.value = f"加權星級: {weighted_score:.2f}, 標準化分數: {transformers_normalized:.2f}"

        #  結果輸出到終端 
        print(f"C{row} 的 Flair 標籤: {flair_label} ({flair_normalized:.2f}) -> 寫入 I{row}")
        print(f"C{row} 的 Transformers 標準化: {transformers_normalized:.2f} -> 寫入 J{row}")

# 儲存修改後的 Excel 檔案 
wb.save(file_path)

# 顯示完成訊息
print("所有情感分數已成功寫入 I 列和 J 列中。")
