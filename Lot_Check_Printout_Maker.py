import os
import tkinter as tk
import tkinter.filedialog as fd
import shutil
import openpyxl

class App(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.master.title("Excel File Selector")
        self.master.geometry("300x200")
        self.create_widgets()
        self.filepath = None

    def create_widgets(self):
        self.select_file_button = tk.Button(self.master, text="Select Excel File", command=self.select_file)
        self.select_file_button.pack(pady=20)
        self.run_button = tk.Button(self.master, text="Run Comparison", command=self.run_comparison)
        self.run_button.pack(pady=20)

    def select_file(self):
        self.filepath = fd.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
        if self.filepath:
            print(f"Selected file: {self.filepath}")

    def run_comparison(self):
        if not self.filepath:
            print("Please select an Excel file first.")
            return

        # Copy the selected file to Lot_Number_Copy.xlsx or Lot_Number_Copy_Old.xlsx
        old_file = "Lot_Number_Copy.xlsx"
        new_file = "Lot_Number_Copy_Old.xlsx"
        if os.path.exists(old_file):
            shutil.copy(old_file, new_file)
            os.remove(old_file)
            os.rename(new_file, old_file)
        shutil.copy(self.filepath, old_file)

        # Compare the first row of each column for Lot_Number_Copy.xlsx and Lot_Number_Copy_Old.xlsx
        old_wb = openpyxl.load_workbook(old_file)
        new_wb = openpyxl.load_workbook(new_file)
        old_sheet = old_wb.active
        new_sheet = new_wb.active
        column_diff = set(old_sheet.columns) - set(new_sheet.columns)

        # Create a new excel file named Lot_Number_Check_Print with columns that exist in Lot_Number_Copy.xlsx but not in Lot_Number_Copy_Old.xlsx
        new_filename = "Lot_Number_Check.xlsx"
        if os.path.exists(new_filename):
            os.rename(new_filename, "Lot_Number_Check_Old.xlsx")
        if column_diff:
            new_wb = openpyxl.Workbook()
            new_sheet = new_wb.active
            for col in column_diff:
                new_sheet.column_dimensions[col[0].column_letter].width = old_sheet.column_dimensions[col[0].column_letter].width
                for i, cell in enumerate(old_sheet[col[0].column_letter]):
                    new_sheet.cell(row=i+1, column=col[0].column_index, value=cell.value)
            new_wb.save(new_filename)
            print(f"Created new file: {new_filename} with columns: {column_diff}")
        else:
            print("No differences found between files.")

root = tk.Tk()
app = App(root)
app.pack()
root.mainloop()
