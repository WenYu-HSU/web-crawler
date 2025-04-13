# 匯入必要的套件
import openpyxl                      # 用來讀取與寫入 Excel 檔案
from snownlp import SnowNLP          # 用來進行情感分析的套件（適用中文）

# 設定 Excel 檔案的路徑
file_path = r'C:\Users\OneDrive\threads1.xlsx'  

# 開啟 Excel 檔案
wb = openpyxl.load_workbook(file_path)

# 選擇要操作的工作表
sheet = wb['Sheet1'] 

# 從第2列開始，遍歷 C 列的每一列文字內容
for row in range(2, sheet.max_row + 1):
    
    # 取得 C 欄的儲存格
    c_cell = sheet[f'C{row}']

    # 取得 I 欄的儲存格
    i_cell = sheet[f'I{row}']

    # 取得 C 欄的實際文字內容
    c_text = c_cell.value

    # 檢查該儲存格是否有文字（避免空白內容）
    if c_text:
        # 使用 SnowNLP 執行中文情感分析
        s = SnowNLP(c_text)

        # 獲取情感分數（介於 0~1，越接近 1 越正面）
        sentiment_score = s.sentiments

        # 將情感分數寫入對應的 I 欄儲存格
        i_cell.value = sentiment_score

        # 輸出處理結果到終端機
        print(f"C{row} 的情感分數: {sentiment_score} 已寫入 I{row}")

# 最後儲存修改後的 Excel 檔案
wb.save(file_path)

# 顯示完成訊息
print("所有情感分數已經成功寫入 I 列中。")