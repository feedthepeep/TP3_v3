import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class SimpleExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Table Viewer")
        self.root.geometry("1300x700")
        self.root.resizable(False, False)
        
        main_container = ttk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        #Таблица 
        left_frame = ttk.Frame(main_container, width=1000, height=900)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        left_frame.pack_propagate(False)  
        
        # Верхняя панель
        top_frame = ttk.Frame(left_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(top_frame, text="Открыть Excel", command=self.load_excel).pack(side=tk.LEFT, padx=5)
        self.info_label = ttk.Label(top_frame, text="Готов к загрузке")
        self.info_label.pack(side=tk.LEFT, padx=10)
        
        # Фрейм для таблицы
        table_frame = ttk.Frame(left_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Скроллы
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(table_frame, yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.df = None
        
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            try:
                #Загружаем данные
                self.df = pd.read_excel(file_path, header=1)
                
                column_names = [
                    'Регион', 
                    'Код Региона',
                    '2011', '2012', '2013', '2014', '2015', '2016', '2017'
                ]
                self.df.columns = column_names
                
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                columns = [str(col) for col in self.df.columns]
                self.tree["columns"] = columns
                self.tree["show"] = "headings"
                
                # Установка фиксорованной ширины колонок
                for i, col in enumerate(columns):
                    self.tree.heading(col, text=col)
                    if i == 0:  
                        self.tree.column(col, width=350, minwidth=350, anchor='w')
                    elif i == 1: 
                        self.tree.column(col, width=120, minwidth=120, anchor='center')
                    else:  
                        self.tree.column(col, width=100, minwidth=100, anchor='center')
                
                # Вставка данных
                for _, row in self.df.iterrows():
                    values = [str(val) if pd.notna(val) else "" for val in row]
                    self.tree.insert("", "end", values=values)
                
                self.info_label.config(text=f"Загружено: {len(self.df)} строк, {len(self.df.columns)} колонок")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleExcelApp(root)
    root.mainloop()