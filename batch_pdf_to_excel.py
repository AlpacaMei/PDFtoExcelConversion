# batch_pdf_to_excel.py

import tkinter as tk
from tkinter import filedialog, simpledialog
from pdf_to_excel import PDFToExcel  # 假設這是 PDFToExcel 類別的模組

class BatchPDFConverter:
    def __init__(self):
        self.pdf_files = []
        self.password = None

    def select_files(self):
        """讓使用者選擇多個 PDF 檔案"""
        root = tk.Tk()
        root.withdraw()  # 隱藏主視窗
        self.pdf_files = filedialog.askopenfilenames(title="選擇 PDF 檔案", filetypes=[("PDF Files", "*.pdf")])

        if not self.pdf_files:
            print("未選擇檔案，程式結束。")
            exit()

    def prompt_password(self):
        """讓使用者輸入密碼"""
        self.password = simpledialog.askstring("輸入密碼", "請輸入 PDF 密碼：", show="*")

        if not self.password:
            print("未輸入密碼，程式結束。")
            exit()

    def run_batch(self):
        """執行批次處理"""
        for pdf_file in self.pdf_files:
            print(f"🔄 處理中：{pdf_file}")
            converter = PDFToExcel(pdf_file, self.password)  # 傳遞檔案和密碼
            try:
                converter.run()  # 執行單一 PDF 的處理
                print(f"✅ 完成：{pdf_file}")
            except Exception as e:
                print(f"❌ 錯誤處理 {pdf_file}：{e}")

    def run(self):
        """執行批次處理的完整流程"""
        self.select_files()  # 選擇檔案
        self.prompt_password()  # 輸入密碼
        self.run_batch()  # 開始批次處理


if __name__ == "__main__":
    batch = BatchPDFConverter()  # 建立 BatchPDFConverter 物件
    batch.run()  # 呼叫 run 方法來執行批次處理，會自動選擇檔案和輸入密碼