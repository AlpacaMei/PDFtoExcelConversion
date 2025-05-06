import tkinter as tk
from tkinter import filedialog, simpledialog
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

    def select_pdf(self):
        root = tk.Tk()
        root.withdraw()
        self.pdf_path = filedialog.askopenfilename(
            title="選擇 PDF 檔案", filetypes=[("PDF Files", "*.pdf")]
        )
        if not self.pdf_path:
            print("未選擇檔案，程式結束。")
            exit()

    def enter_password(self):
        self.pdf_password = simpledialog.askstring("輸入密碼", "請輸入 PDF 密碼：", show="*")
        if not self.pdf_password:
            print("未輸入密碼，程式結束。")
            exit()

    def decrypt_pdf(self):
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
        all_tables = []
        try:
            with pdfplumber.open(self.unlocked_pdf) as pdf:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        headers = table[0]
                        # 自動處理欄位名稱重複
                        seen = {}
                        new_headers = []
                        for h in headers:
                            if h in seen:
                                seen[h] += 1
                                new_headers.append(f"{h}_{seen[h]}")
                            else:
                                seen[h] = 0
                                new_headers.append(h)
                        df = pd.DataFrame(table[1:], columns=new_headers)
                        df["頁碼"] = i + 1
                        all_tables.append(df)
        except Exception as e:
            print(f"發生錯誤: {e}")
            exit()

        if not all_tables:
            print("⚠️ PDF 中未找到任何表格。")
            exit()

        return pd.concat(all_tables, ignore_index=True)

    def save_to_excel(self, df):
        base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        folder_path = os.path.dirname(self.pdf_path)
        excel_output = os.path.join(folder_path, f"{base_name}.xlsx")

        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        # 使用 pywin32 設定密碼
        excel = win32.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(os.path.abspath(excel_output))
        workbook.Password = self.pdf_password
        workbook.SaveAs(os.path.abspath(excel_output))
        workbook.Close()
        excel.Quit()

        print(f"✅ 已成功轉換為加密的 Excel：{excel_output}")
        return excel_output

    def clean_up(self):
        if os.path.exists(self.unlocked_pdf):
            os.remove(self.unlocked_pdf)

    def run(self):
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
