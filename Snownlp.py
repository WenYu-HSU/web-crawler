import openpyxl
from snownlp import SnowNLP

try:
    # 打開 Excel 文件
    file_path = r"C:\Users\user\OneDrive\桌面\文字雲分析１.xlsx"
    
    # 先確認檔案可以被開啟
    wb = openpyxl.load_workbook(file_path)
    sheet = wb['生活不含負面']
    
    # 執行情感分析
    for row in range(2, sheet.max_row + 1):
        a_cell = sheet[f'A{row}']
        b_cell = sheet[f'B{row}']
        a_text = a_cell.value
        if a_text:
            s = SnowNLP(str(a_text))
            sentiment_score = s.sentiments
            b_cell.value = sentiment_score
            print(f"A{row} 的情感分數: {sentiment_score} 已寫入 B{row}")


        d_cell = sheet[f'D{row}']
        e_cell = sheet[f'E{row}']
        d_text = d_cell.value
        if d_text:
            s = SnowNLP(str(d_text))
            sentiment_score = s.sentiments
            e_cell.value = sentiment_score
            print(f"D{row} 的情感分數: {sentiment_score} 已寫入 E{row}")
        
    # 嘗試儲存到不同的檔案名稱
    save_path = r'C:\Users\user\OneDrive\桌面\threads_analysis.xlsx'
    wb.save(save_path)
    print("所有情感分數已經成功寫入新檔案中。")

except openpyxl.utils.exceptions.InvalidFileException:
    print("錯誤：請確認 Excel 檔案格式是否為 .xlsx、.xlsm、.xltx 或 .xltm")
except PermissionError:
    print("錯誤：檔案可能正在被使用中。請關閉 Excel 後再試一次")
except Exception as e:
    print(f"發生錯誤：{str(e)}")