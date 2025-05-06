from batch_pdf_to_excel import BatchPDFConverter

if __name__ == "__main__":
    batch = BatchPDFConverter()  # 建立 BatchPDFConverter 物件
    batch.run()  # 呼叫 run 方法來執行批次處理，會自動選擇檔案和輸入密碼
