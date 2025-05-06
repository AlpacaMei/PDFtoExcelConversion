import os
import tkinter as tk
from tkinter import filedialog, simpledialog
from pypdf import PdfReader, PdfWriter
import pdfplumber
import pandas as pd
import xlsxwriter
import win32com.client as win32


class PDFToExcel:
    def __init__(self, pdf_path, pdf_password):
        """初始化 PDFToExcel 物件，並設置 PDF 檔案路徑與密碼"""
        self.pdf_path = pdf_path
        self.pdf_password = pdf_password
        self.unlocked_pdf = "unlocked.pdf"
        self.save_folder = os.path.join(os.getcwd(), 'datasets')
        if not os.path.exists(self.save_folder):
            os.makedirs(self.save_folder)  # 如果資料夾不存在，則建立

    def decrypt_pdf(self):
        """解密 PDF 並儲存成新的無密碼 PDF"""
        try:
            reader = PdfReader(self.pdf_path)
            reader.decrypt(self.pdf_password)

            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            with open(self.unlocked_pdf, "wb") as f:
                writer.write(f)
        except Exception as e:
            print(f"發生錯誤: {e}")
            exit()

    def make_unique_columns(self, columns):
        """檢查欄位名稱是否重複，如果有重複則重新命名"""
        seen = set()
        unique_columns = []
        for col in columns:
            new_col = col
            count = 1
            while new_col in seen:
                new_col = f"{col}_{count}"
                count += 1
            seen.add(new_col)
            unique_columns.append(new_col)
        return unique_columns

    def extract_tables(self):
        """提取 PDF 表格內容為 DataFrame"""
        all_tables = []

        with pdfplumber.open(self.unlocked_pdf) as pdf:
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue  # 跳過空表格或缺少標題/資料列的表格

                    # 檢查欄位是否唯一，不唯一就重新命名
                    raw_columns = table[0]
                    if len(set(raw_columns)) != len(raw_columns):
                        columns = self.make_unique_columns(raw_columns)
                    else:
                        columns = raw_columns

                    try:
                        df = pd.DataFrame(table[1:], columns=columns)
                        df["頁碼"] = i + 1

                        # 確保每個表格的索引唯一
                        df = df.reset_index(drop=True)

                        all_tables.append(df)
                    except Exception as e:
                        print(f"⚠️ 第 {i+1} 頁表格格式異常，略過：{e}")
                        continue

        if not all_tables:
            raise ValueError("未從 PDF 中擷取到任何有效表格。")

        # 合併所有表格並重設索引
        final_df = pd.concat(all_tables, ignore_index=True)

        # 確保合併後的 DataFrame 有唯一索引
        final_df = final_df.reset_index(drop=True)

        return final_df

    def save_to_excel(self, df):
        """儲存 DataFrame 到 Excel，並加密"""
        base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        excel_output = os.path.join(self.save_folder, f"{base_name}.xlsx")

        # 使用 pandas 的 ExcelWriter 輸出為 Excel
        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        # 使用 pywin32 來加密 Excel
        excel = win32.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(os.path.abspath(excel_output))
        workbook.Password = self.pdf_password  # 設置密碼
        workbook.SaveAs(os.path.abspath(excel_output))  # 儲存並保護密碼
        workbook.Close()
        excel.Quit()

        print(f"✅ 已成功轉換為加密的 Excel：{excel_output}")
        return excel_output

    def clean_up(self):
        """清除中間檔（可選）"""
        if os.path.exists(self.unlocked_pdf):
            os.remove(self.unlocked_pdf)

    def run(self):
        """執行 PDF 轉換到 Excel 的流程"""
        self.decrypt_pdf()
        df = self.extract_tables()
        self.save_to_excel(df)
        self.clean_up()


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
    batch = BatchPDFConverter()
    batch.run()  # 呼叫 run 方法來執行批次處理，會自動選擇檔案和輸入密碼
