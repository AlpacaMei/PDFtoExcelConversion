import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from pypdf import PdfReader, PdfWriter
import pdfplumber
import pandas as pd
import os
import win32com.client as win32

class PDFToExcel:
    def __init__(self):
        self.pdf_path = None
        self.pdf_password = None
        self.unlocked_pdf = "unlocked.pdf"
        self.save_folder = os.path.join(os.getcwd(), 'datasets')
        if not os.path.exists(self.save_folder):
            os.makedirs(self.save_folder)  # 如果資料夾不存在，則建立

    def select_pdf(self):
        """讓使用者選擇 PDF 檔案"""
        root = tk.Tk()
        root.withdraw()  # 隱藏主視窗
        self.pdf_path = filedialog.askopenfilename(title="選擇 PDF 檔案", filetypes=[("PDF Files", "*.pdf")])

        if not self.pdf_path:
            print("未選擇檔案，程式結束。")
            exit()

    def enter_password(self):
        """讓使用者輸入密碼"""
        self.pdf_password = simpledialog.askstring("輸入密碼", "請輸入 PDF 密碼：", show="*")

        if not self.pdf_password:
            print("未輸入密碼，程式結束。")
            exit()

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

    def extract_tables(self):
        """使用 pdfplumber 讀取表格並儲存為 DataFrame"""
        all_tables = []

        try:
            with pdfplumber.open(self.unlocked_pdf) as pdf:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        df = pd.DataFrame(table[1:], columns=table[0])
                        df["頁碼"] = i + 1  # 加上頁碼欄位（可省略）
                        all_tables.append(df)
        except Exception as e:
            print(f"發生錯誤: {e}")
            exit()

        return pd.concat(all_tables, ignore_index=True)

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
        self.select_pdf()
        self.enter_password()
        self.decrypt_pdf()
        df = self.extract_tables()
        self.save_to_excel(df)
        self.clean_up()


# 執行程式
if __name__ == "__main__":
    converter = PDFToExcel()
    converter.run()
